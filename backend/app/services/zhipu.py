from __future__ import annotations

import json
import os
import re
from typing import Any, Dict

import httpx
from dotenv import load_dotenv


load_dotenv()


DEFAULT_MODEL = "glm-4.7-flash"
DEFAULT_BASE_URL = "https://open.bigmodel.cn/api/anthropic"


def _extract_json(text: str) -> Dict[str, Any]:
    text = text.strip()
    if text.startswith("{") and text.endswith("}"):
        return json.loads(text)
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise ValueError("no_json_found")
    return json.loads(match.group(0))


def parse_to_operations(message: str, sheet: str | None = None) -> Dict[str, Any]:
    api_key = os.getenv("ZHIPU_API_KEY")
    if not api_key:
        raise RuntimeError("missing_api_key")
    model = os.getenv("ZHIPU_MODEL", DEFAULT_MODEL)
    base_url = os.getenv("ZHIPU_BASE_URL", DEFAULT_BASE_URL).rstrip("/")

    system_prompt = (
        "You are an Excel operation parser. Convert the user request into JSON with keys: "
        "message (string) and operations (array). Each operation must match one of: "
        "set_cell {type, sheet, cell, value}; set_range {type, sheet, range, value}; "
        "add_column {type, sheet, column_name, value}; rename_column {type, sheet, column, new_name}; "
        "swap_columns {type, sheet, column_a, column_b} or {type, sheet, column_index_a, column_index_b}; "
        "round_column {type, sheet, column, decimals}; format_lt {type, sheet, column, threshold, color}; "
        "t_test {type, sheet, column_a, column_b, equal_var, output, output_prefix}; "
        "rename_sheet {type, from, to}; "
        "add_sheet {type, to}; delete_rows {type, sheet, rows}; "
        "update_cells {type, sheet, where:{column,value}, set:{col:value}}; "
        "sort {type, sheet, by, ascending}. "
        "Only return strict JSON."
    )
    if sheet:
        system_prompt += f" Default sheet is {sheet} unless user specifies otherwise."

    payload = {
        "model": model,
        "max_tokens": 1024,
        "messages": [{"role": "user", "content": message}],
        "system": system_prompt,
    }

    headers = {"x-api-key": api_key, "content-type": "application/json"}
    with httpx.Client(timeout=30) as client:
        resp = client.post(f"{base_url}/v1/messages", headers=headers, json=payload)
        if resp.status_code >= 400:
            raise RuntimeError(f"zhipu_error:{resp.status_code}")
        data = resp.json()

    content = data.get("content", "")
    if isinstance(content, list) and content:
        content = content[0].get("text", "")
    if not isinstance(content, str):
        raise RuntimeError("invalid_response")
    return _extract_json(content)
