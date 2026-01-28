from typing import Any, Dict, List, Literal, Optional
from pydantic import BaseModel, Field


class CreateSessionRequest(BaseModel):
    name: Optional[str] = None


class CreateEmptyWorkbookRequest(BaseModel):
    filename: str = Field(default="blank.xlsx")
    sheet_name: str = Field(default="Sheet1")


class PreviewRequest(BaseModel):
    sheet: Optional[str] = None
    limit: int = Field(default=100, ge=1, le=1000)


class Operation(BaseModel):
    type: Literal[
        "set_cell",
        "set_range",
        "add_column",
        "swap_columns",
        "rename_column",
        "round_column",
        "format_lt",
        "t_test",
        "rename_sheet",
        "add_sheet",
        "delete_rows",
        "update_cells",
        "sort",
    ]
    sheet: Optional[str] = None
    cell: Optional[str] = None
    range: Optional[str] = None
    value: Optional[Any] = None
    column: Optional[str] = None
    column_name: Optional[str] = None
    column_a: Optional[str] = None
    column_b: Optional[str] = None
    column_index_a: Optional[int] = None
    column_index_b: Optional[int] = None
    decimals: Optional[int] = None
    threshold: Optional[float] = None
    color: Optional[str] = None
    equal_var: Optional[bool] = False
    output: Optional[Literal["sheet", "column"]] = "sheet"
    output_prefix: Optional[str] = None
    new_name: Optional[str] = None
    rows: Optional[List[int]] = None
    where: Optional[Dict[str, Any]] = None
    set: Optional[Dict[str, Any]] = None
    by: Optional[str] = None
    ascending: Optional[bool] = True
    from_sheet: Optional[str] = Field(default=None, alias="from")
    to_sheet: Optional[str] = Field(default=None, alias="to")


class ApplyOperationsRequest(BaseModel):
    message: Optional[str] = None
    operations: List[Operation]


class CommitSummary(BaseModel):
    id: str
    message: str
    timestamp: str
    changed_sheets: List[str]


class AnalysisResult(BaseModel):
    type: Literal["t_test"]
    sheet: str
    column_a: str
    column_b: str
    n_a: int
    n_b: int
    mean_a: float
    mean_b: float
    t_stat: float
    p_value: float
    df: float


class ApplyOperationsResponse(BaseModel):
    commit: CommitSummary
    analysis: Optional[List[AnalysisResult]] = None


class HistoryResponse(BaseModel):
    commits: List[CommitSummary]


class RollbackRequest(BaseModel):
    commit_id: str


class BatchApplyRequest(BaseModel):
    message: Optional[str] = None
    operations: List[Operation]


class ParseRequest(BaseModel):
    message: str
    sheet: Optional[str] = None


class ParseResponse(BaseModel):
    message: str
    operations: List[Operation]
