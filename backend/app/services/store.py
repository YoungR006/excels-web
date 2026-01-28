from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Dict, List, Optional
from uuid import uuid4

import pandas as pd


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


@dataclass
class Commit:
    id: str
    message: str
    timestamp: str
    changed_sheets: List[str]
    snapshot: Dict[str, pd.DataFrame]
    format_rules: List[dict]


@dataclass
class WorkbookState:
    filename: str
    sheets: Dict[str, pd.DataFrame]
    format_rules: List[dict] = field(default_factory=list)
    commits: List[Commit] = field(default_factory=list)

    def snapshot(self) -> Dict[str, pd.DataFrame]:
        return {name: df.copy(deep=True) for name, df in self.sheets.items()}


@dataclass
class Session:
    id: str
    name: Optional[str]
    workbooks: Dict[str, WorkbookState] = field(default_factory=dict)


class InMemoryStore:
    def __init__(self) -> None:
        self.sessions: Dict[str, Session] = {}

    def create_session(self, name: Optional[str] = None) -> Session:
        session = Session(id=str(uuid4()), name=name)
        self.sessions[session.id] = session
        return session

    def get_session(self, session_id: str) -> Session:
        if session_id not in self.sessions:
            raise KeyError("session_not_found")
        return self.sessions[session_id]

    def add_workbook(self, session: Session, filename: str, sheets: Dict[str, pd.DataFrame]) -> WorkbookState:
        workbook = WorkbookState(filename=filename, sheets=sheets)
        session.workbooks[filename] = workbook
        self._commit(workbook, message="init")
        return workbook

    def get_workbook(self, session: Session, filename: str) -> WorkbookState:
        if filename not in session.workbooks:
            raise KeyError("workbook_not_found")
        return session.workbooks[filename]

    def _commit(self, workbook: WorkbookState, message: str, changed_sheets: Optional[List[str]] = None) -> Commit:
        commit = Commit(
            id=str(uuid4()),
            message=message or "update",
            timestamp=_now_iso(),
            changed_sheets=changed_sheets or list(workbook.sheets.keys()),
            snapshot=workbook.snapshot(),
            format_rules=[rule.copy() for rule in workbook.format_rules],
        )
        workbook.commits.append(commit)
        return commit

    def commit(self, workbook: WorkbookState, message: str, changed_sheets: List[str]) -> Commit:
        return self._commit(workbook, message=message, changed_sheets=changed_sheets)

    def history(self, workbook: WorkbookState) -> List[Commit]:
        return list(workbook.commits)

    def rollback(self, workbook: WorkbookState, commit_id: str) -> Commit:
        target = None
        for commit in workbook.commits:
            if commit.id == commit_id:
                target = commit
                break
        if target is None:
            raise KeyError("commit_not_found")
        workbook.sheets = {name: df.copy(deep=True) for name, df in target.snapshot.items()}
        workbook.format_rules = [rule.copy() for rule in target.format_rules]
        return self._commit(workbook, message=f"rollback:{commit_id}", changed_sheets=list(workbook.sheets.keys()))


STORE = InMemoryStore()
