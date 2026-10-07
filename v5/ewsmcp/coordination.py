"""Per-DATA_DIR coordination shared by all local stdio processes.

The control database stores leases, cooldown and idempotent operation state.
Small byte-range lock files enforce the parts that must span the duration of
blocking EWS calls.  No third-party daemon is required, so this is also safe
for the existing stdio deployment; a single HTTP server remains the preferred
long-term topology.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import sqlite3
import time
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Callable, Dict, Iterator, Optional, Tuple

from .errors import ToolError


def _lock_byte(fh: Any, blocking: bool) -> bool:
    if os.name == "nt":
        import msvcrt

        fh.seek(0)
        mode = msvcrt.LK_LOCK if blocking else msvcrt.LK_NBLCK
        try:
            msvcrt.locking(fh.fileno(), mode, 1)
            return True
        except OSError:
            return False
    import fcntl

    flags = fcntl.LOCK_EX | (0 if blocking else fcntl.LOCK_NB)
    try:
        fcntl.flock(fh.fileno(), flags)
        return True
    except OSError:
        return False


def _unlock_byte(fh: Any) -> None:
    try:
        if os.name == "nt":
            import msvcrt

            fh.seek(0)
            msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl

            fcntl.flock(fh.fileno(), fcntl.LOCK_UN)
    except OSError:
        pass


class FileMutex:
    def __init__(self, path: Path):
        self.path = path

    def acquire(self, timeout_s: float = 0.0) -> Optional[Any]:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        fh = self.path.open("a+b")
        if self.path.stat().st_size == 0:
            fh.write(b"0")
            fh.flush()
        deadline = time.monotonic() + max(0.0, timeout_s)
        while True:
            if _lock_byte(fh, blocking=False):
                return fh
            if time.monotonic() >= deadline:
                fh.close()
                return None
            time.sleep(0.05)

    @staticmethod
    def release(fh: Any) -> None:
        _unlock_byte(fh)
        fh.close()

    @contextmanager
    def held(self, timeout_s: float) -> Iterator[None]:
        fh = self.acquire(timeout_s)
        if fh is None:
            raise TimeoutError("interprocess lock wait timed out")
        try:
            yield
        finally:
            self.release(fh)


class ProcessCoordinator:
    ACTIVE_STATES = {"queued", "running", "rate_limited", "unknown_outcome"}

    def __init__(self, data_dir: str, mailbox: str):
        self.root = Path(data_dir) / "coordination"
        self.root.mkdir(parents=True, exist_ok=True)
        self.db_path = self.root / "control.db"
        self.instance_id = uuid.uuid4().hex[:12]
        self.pid = os.getpid()
        self.mailbox = (mailbox or "default").strip().lower()
        self._leader_fh: Optional[Any] = None
        self._waiters = 0
        self._ensure_schema()
        self._register_instance()
        self.try_become_sync_leader()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path), timeout=10.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        return conn

    def _ensure_schema(self) -> None:
        with self._connect() as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS instances (
                    instance_id TEXT PRIMARY KEY, pid INTEGER, mailbox TEXT,
                    started REAL, heartbeat REAL
                );
                CREATE TABLE IF NOT EXISTS state (
                    key TEXT PRIMARY KEY, value TEXT, updated REAL
                );
                CREATE TABLE IF NOT EXISTS operations (
                    operation_id TEXT PRIMARY KEY, idempotency_key TEXT NOT NULL,
                    kind TEXT NOT NULL, args_hash TEXT NOT NULL,
                    status TEXT NOT NULL, result_json TEXT, error TEXT,
                    created REAL NOT NULL, updated REAL NOT NULL,
                    UNIQUE(kind, idempotency_key)
                );
                CREATE INDEX IF NOT EXISTS ix_operations_updated
                    ON operations(updated DESC);
            """)

    def _register_instance(self) -> None:
        now = time.time()
        with self._connect() as conn:
            conn.execute(
                "INSERT OR REPLACE INTO instances VALUES (?, ?, ?, ?, ?)",
                (self.instance_id, self.pid, self.mailbox, now, now),
            )
            conn.execute("DELETE FROM instances WHERE heartbeat < ?", (now - 3600,))
            # Every start used to leave a waiters:<id> row behind forever;
            # 160+ of them had accumulated. Drop the ones whose instance is
            # gone — the state table is a live view, not a journal.
            conn.execute(
                "DELETE FROM state WHERE key LIKE 'waiters:%' AND "
                "substr(key, 9) NOT IN (SELECT instance_id FROM instances)"
            )

    def heartbeat(self) -> None:
        with self._connect() as conn:
            conn.execute(
                "UPDATE instances SET heartbeat=? WHERE instance_id=?",
                (time.time(), self.instance_id),
            )

    def try_become_sync_leader(self) -> bool:
        if self._leader_fh is not None:
            return True
        fh = FileMutex(self.root / "sync-leader.lock").acquire(0)
        if fh is None:
            return False
        self._leader_fh = fh
        self._set_state(
            "sync_leader",
            json.dumps(
                {
                    "instance_id": self.instance_id,
                    "pid": self.pid,
                }
            ),
        )
        return True

    @property
    def is_sync_leader(self) -> bool:
        return self._leader_fh is not None

    def _set_state(self, key: str, value: str) -> None:
        with self._connect() as conn:
            conn.execute(
                "INSERT OR REPLACE INTO state(key, value, updated) VALUES (?, ?, ?)",
                (key, value, time.time()),
            )

    def folder_revision(self) -> int:
        """Shared generation of the mailbox folder tree.

        Every stdio process has its own exchangelib folder cache.  This value
        lets one process announce a mutation so the others invalidate their
        in-memory trees before resolving another folder.
        """
        with self._connect() as conn:
            row = conn.execute(
                "SELECT value FROM state WHERE key=?",
                (f"folder_revision:{self.mailbox}",),
            ).fetchone()
        try:
            return int(row["value"]) if row is not None else 0
        except (TypeError, ValueError):
            return 0

    def bump_folder_revision(self) -> int:
        key = f"folder_revision:{self.mailbox}"
        with self._connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute(
                "SELECT value FROM state WHERE key=?", (key,)
            ).fetchone()
            try:
                revision = int(row["value"]) + 1 if row is not None else 1
            except (TypeError, ValueError):
                revision = 1
            conn.execute(
                "INSERT OR REPLACE INTO state(key, value, updated) VALUES (?, ?, ?)",
                (key, str(revision), time.time()),
            )
            conn.commit()
        return revision

    def set_cooldown(self, seconds: float, reason: str = "Exchange backoff") -> None:
        until = time.time() + max(0.0, float(seconds))
        self._set_state(
            f"cooldown:{self.mailbox}",
            json.dumps(
                {
                    "until": until,
                    "reason": reason[:500],
                    "instance_id": self.instance_id,
                }
            ),
        )

    def cooldown(self) -> Optional[Dict[str, Any]]:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT value FROM state WHERE key=?",
                (f"cooldown:{self.mailbox}",),
            ).fetchone()
        if row is None:
            return None
        try:
            value = json.loads(row["value"])
            remaining = max(0, int(float(value["until"]) - time.time() + 0.999))
            if remaining <= 0:
                return None
            return {**value, "retry_after_s": remaining}
        except Exception:
            return None

    @staticmethod
    def backoff_seconds(exc: Exception) -> Optional[float]:
        name = type(exc).__name__.lower()
        message = str(exc)
        if "ratelimit" not in name and "throttl" not in name and "back off" not in message.lower():
            return None
        matches = re.findall(
            r"(?:back off|retry after)[^0-9]*(\d+(?:\.\d+)?)", message, flags=re.IGNORECASE
        )
        return float(matches[-1]) if matches else 60.0

    @contextmanager
    def live_call(self, lane: str, timeout_s: float) -> Iterator[None]:
        cooldown = self.cooldown()
        if cooldown:
            raise ToolError(
                "rate_limited",
                "Exchange cooldown is active for this mailbox",
                retry_after_s=cooldown["retry_after_s"],
            )
        lane = lane if lane in {"calendar", "read", "write", "sync"} else "read"
        mutex = FileMutex(self.root / f"live-{lane}.lock")
        self._waiters += 1
        self._set_state(f"waiters:{self.instance_id}", str(self._waiters))
        try:
            fh = mutex.acquire(timeout_s)
        finally:
            self._waiters -= 1
            self._set_state(f"waiters:{self.instance_id}", str(self._waiters))
        if fh is None:
            raise ToolError(
                "upstream_timeout",
                f"Timed out waiting for the shared Exchange {lane} lane",
                retry_after_s=1,
            )
        self._set_state(
            f"lock:{lane}",
            json.dumps(
                {
                    "instance_id": self.instance_id,
                    "pid": self.pid,
                    "since": time.time(),
                }
            ),
        )
        try:
            yield
        finally:
            self._set_state(f"lock:{lane}", "")
            mutex.release(fh)

    @staticmethod
    def args_hash(payload: Dict[str, Any]) -> str:
        canonical = json.dumps(
            payload, sort_keys=True, ensure_ascii=False, separators=(",", ":"), default=str
        )
        return hashlib.sha256(canonical.encode()).hexdigest()

    def reserve_operation(
        self, kind: str, key: str, payload: Dict[str, Any]
    ) -> Tuple[Dict[str, Any], bool]:
        digest = self.args_hash(payload)
        now = time.time()
        op_id = f"op_{uuid.uuid4().hex}"
        with self._connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute(
                "SELECT * FROM operations WHERE kind=? AND idempotency_key=?",
                (kind, key),
            ).fetchone()
            if row is not None:
                conn.commit()
                if row["args_hash"] != digest:
                    raise ToolError(
                        "idempotency_conflict",
                        "idempotency_key was already used with different arguments",
                        hint="Use a fresh idempotency_key for a distinct create.",
                    )
                return self._operation_dict(row), False
            conn.execute(
                "INSERT INTO operations VALUES (?, ?, ?, ?, 'queued', NULL, NULL, ?, ?)",
                (op_id, key, kind, digest, now, now),
            )
            conn.commit()
        return self.operation(op_id) or {}, True

    def _update_operation(
        self, op_id: str, status: str, *, result: Any = None, error: Optional[str] = None
    ) -> None:
        result_json = (
            json.dumps(result, ensure_ascii=False, default=str) if result is not None else None
        )
        with self._connect() as conn:
            conn.execute(
                "UPDATE operations SET status=?, result_json=COALESCE(?, result_json), "
                "error=?, updated=? WHERE operation_id=?",
                (status, result_json, error, time.time(), op_id),
            )

    def execute_operation(self, op_id: str, fn: Callable[[], Any]) -> Any:
        self._update_operation(op_id, "running")
        try:
            result = fn()
        except Exception as exc:
            backoff = self.backoff_seconds(exc)
            if backoff is not None:
                self.set_cooldown(backoff, f"{type(exc).__name__}: {exc}")
                self._update_operation(op_id, "rate_limited", error=str(exc)[:500])
            else:
                self._update_operation(op_id, "failed", error=f"{type(exc).__name__}: {exc}"[:500])
            raise
        self._update_operation(op_id, "completed", result=result)
        return result

    def mark_unknown(self, op_id: str, error: str) -> None:
        op = self.operation(op_id)
        if op and op["status"] in {"queued", "running"}:
            self._update_operation(op_id, "unknown_outcome", error=error[:500])

    @staticmethod
    def _operation_dict(row: sqlite3.Row, include_result: bool = True) -> Dict[str, Any]:
        out = {
            key: row[key]
            for key in ("operation_id", "kind", "status", "created", "updated", "error")
        }
        if include_result and row["result_json"]:
            try:
                out["result"] = json.loads(row["result_json"])
            except json.JSONDecodeError:
                pass
        return out

    def operation(self, op_id: str) -> Optional[Dict[str, Any]]:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM operations WHERE operation_id=?",
                (op_id,),
            ).fetchone()
        return self._operation_dict(row) if row else None

    def recent_operations(self, limit: int = 10) -> list[Dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM operations ORDER BY updated DESC LIMIT ?",
                (int(limit),),
            ).fetchall()
        return [self._operation_dict(row, include_result=False) for row in rows]

    def status(self) -> Dict[str, Any]:
        self.heartbeat()
        locks: Dict[str, Any] = {}
        leader_holder = None
        with self._connect() as conn:
            for row in conn.execute("SELECT key, value FROM state WHERE key LIKE 'lock:%'"):
                if row["value"]:
                    try:
                        locks[row["key"].split(":", 1)[1]] = json.loads(row["value"])
                    except json.JSONDecodeError:
                        pass
            row = conn.execute("SELECT value FROM state WHERE key='sync_leader'").fetchone()
            if row and row["value"]:
                try:
                    leader_holder = json.loads(row["value"])
                except json.JSONDecodeError:
                    pass
            instance_count = conn.execute(
                "SELECT COUNT(*) AS c FROM instances WHERE heartbeat >= ?",
                (time.time() - 120,),
            ).fetchone()["c"]
            waiting_total = 0
            for row in conn.execute(
                "SELECT value FROM state WHERE key LIKE 'waiters:%' AND updated >= ?",
                (time.time() - 120,),
            ):
                try:
                    waiting_total += int(row["value"] or 0)
                except ValueError:
                    pass
        return {
            "instance_id": self.instance_id,
            "pid": self.pid,
            "sync_leader": self.is_sync_leader,
            "sync_leader_holder": leader_holder,
            "active_instances": int(instance_count),
            "lock_holders": locks,
            "cooldown": self.cooldown(),
            "waiting_live_calls": waiting_total,
            "recent_operations": self.recent_operations(),
        }

    def close(self) -> None:
        if self._leader_fh is not None:
            FileMutex.release(self._leader_fh)
            self._leader_fh = None
        try:
            with self._connect() as conn:
                conn.execute("DELETE FROM instances WHERE instance_id=?", (self.instance_id,))
        except sqlite3.Error:
            pass
