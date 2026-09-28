import json
import os
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from openpyxl import load_workbook


SOURCE_FILE = Path(r"C:\Users\xlin075\Documents\mapping\Book1.xlsx")
OUTPUT_DIR = SOURCE_FILE.parent
LOG_PATH = OUTPUT_DIR / "processing_error_log.txt"
ERROR_PATH = OUTPUT_DIR / "mapping_errors.txt"
SUMMARY_LIST_PATH = OUTPUT_DIR / "summary_keywords.json"

API_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com").rstrip("/")
API_KEY = os.getenv("DEEPSEEK_API_KEY")
MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
BATCH_SIZE = int(os.getenv("MAPPING_BATCH_SIZE", "50"))

SUMMARY_KEYWORDS = [
    "\u7968\u636e\u53ca\u6e05\u7b97",
    "\u624b\u7eed\u8d39/\u670d\u52a1\u8d39/\u4ea4\u6613\u8d39",
    "\u6c34\u7535\u53ca\u901a\u8baf\u8d39\u7528",
    "\u884c\u653f\u53ca\u529e\u516c\u8d39\u7528",
    "\u79df\u91d1",
    "\u5229\u606f",
    "\u5b58\u6b3e",
    "\u4eba\u4e8b\u8d39\u7528",
    "\u5de5\u7a0b\u53ca\u88c5\u4fee\u6b3e",
    "\u91c7\u8d2d\u6b3e",
    "\u4e13\u4e1a\u670d\u52a1\u8d39",
    "\u4ea4\u901a\u53ca\u5dee\u65c5\u8d39",
    "\u62bc\u91d1\u4fdd\u8bc1\u91d1",
    "\u4e1a\u52a1\u8d27\u6b3e",
    "\u4fdd\u9669\u57fa\u91d1\u6536\u5165",
    "\u501f\u6b3e",
    "\u5907\u7528\u91d1",
    "\u62a5\u9500\u6b3e",
    "\u7a0e\u52a1\u8d39\u7528",
    "\u5f80\u6765\u6b3e",
]


def log_error(message: str, exc: Exception | None = None) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    text = f"[{timestamp}] {message}\n"
    if exc is not None:
        text += traceback.format_exc() + "\n"
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        handle.write(text)


def make_copy_path() -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return OUTPUT_DIR / f"Book_copy_{timestamp}.xlsx"


def get_header_lookup(ws: Any) -> tuple[int, dict[str, int]]:
    scan_columns = min(ws.max_column, 100)
    for row_number in range(1, min(ws.max_row, 20) + 1):
        values = [ws.cell(row_number, column).value for column in range(1, scan_columns + 1)]
        headers = {
            str(value).strip(): column
            for column, value in enumerate(values, 1)
            if value is not None and str(value).strip()
        }
        if "\u6458\u8981mapping" in headers and "\u6237\u540dmapping" in headers:
            return row_number, headers
    return 0, {}


def request_json(system_prompt: str, user_prompt: str) -> dict[str, Any]:
    if not API_KEY:
        raise RuntimeError(
            "DEEPSEEK_API_KEY is not set. Set it in the process environment; "
            "the key must not be written to source files."
        )

    payload = {
        "model": MODEL,
        "temperature": 0,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    }
    request = Request(
        f"{API_URL}/chat/completions",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {API_KEY}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urlopen(request, timeout=180) as response:
            response_data = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"API request failed with HTTP {exc.code}: {body}") from exc
    except URLError as exc:
        raise RuntimeError(f"API request failed: {exc.reason}") from exc

    try:
        content = response_data["choices"][0]["message"]["content"]
        result = json.loads(content)
    except (KeyError, IndexError, TypeError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"API returned an invalid JSON response: {response_data}") from exc
    if not isinstance(result, dict):
        raise RuntimeError("API returned JSON that is not an object.")
    return result


def batched(values: list[str], size: int) -> list[list[str]]:
    return [values[index : index + size] for index in range(0, len(values), size)]


def map_summaries(values: list[str]) -> dict[str, str]:
    result: dict[str, str] = {}
    system = (
        "You classify Chinese bank transaction summaries. Choose exactly one category "
        "from the supplied keyword list for every input. Use the transaction meaning, "
        "not merely one matching character. Return JSON with an object named "
        "'mappings'; each input string must be a key and its value must be exactly one "
        "keyword from the list."
    )
    for batch in batched(values, BATCH_SIZE):
        user = json.dumps(
            {"keywords": SUMMARY_KEYWORDS, "summaries": batch},
            ensure_ascii=False,
        )
        response = request_json(system, user)
        mappings = response.get("mappings")
        if not isinstance(mappings, dict):
            raise RuntimeError(f"Summary response has no valid mappings object: {response}")
        for source in batch:
            target = mappings.get(source)
            if target not in SUMMARY_KEYWORDS:
                raise RuntimeError(
                    f"Summary mapping for {source!r} is not a supplied keyword: {target!r}"
                )
            result[source] = target
        time.sleep(0.1)
    return result


def map_accounts(values: list[str]) -> dict[str, str]:
    result: dict[str, str] = {}
    system = (
        "You normalize Chinese bank counterparty account names. Return JSON with an "
        "object named 'mappings'. For each input, map it to a canonical account name. "
        "Only merge names when differences are clearly typographical, punctuation, "
        "spacing, legal suffix, or an obvious minor writing error. Do not merge "
        "different companies merely because they are similar. Preserve the original "
        "input as the canonical value when unsure. Every input must be a key."
    )
    for batch in batched(values, BATCH_SIZE):
        user = json.dumps({"account_names": batch}, ensure_ascii=False)
        response = request_json(system, user)
        mappings = response.get("mappings")
        if not isinstance(mappings, dict):
            raise RuntimeError(f"Account response has no valid mappings object: {response}")
        for source in batch:
            target = mappings.get(source)
            if not isinstance(target, str) or not target.strip():
                raise RuntimeError(f"Account mapping for {source!r} is invalid: {target!r}")
            result[source] = target.strip()
        time.sleep(0.1)
    return result


def process_workbook(source_path: Path, copy_path: Path) -> None:
    if not source_path.exists():
        raise FileNotFoundError(source_path)
    if not API_KEY:
        raise RuntimeError("DEEPSEEK_API_KEY is required.")

    workbook = load_workbook(source_path, data_only=False)
    errors: list[str] = []
    sheets: list[tuple[Any, int, dict[str, int]]] = []
    summary_values: set[str] = set()
    account_values: set[str] = set()

    for sheet in workbook.worksheets:
        header_row, headers = get_header_lookup(sheet)
        if not headers:
            errors.append(f"{sheet.title}: required headers were not found")
            continue
        sheets.append((sheet, header_row, headers))
        summary_column = headers["\u6458\u8981mapping"]
        account_column = headers["\u6237\u540dmapping"]
        summary_source_column = headers.get("\u6458\u8981\u63cf\u8ff0")
        account_source_column = headers.get("\u5bf9\u65b9\u6237\u540d")
        if summary_source_column is None or account_source_column is None:
            errors.append(f"{sheet.title}: summary or account source header was not found")
            continue
        for row in range(header_row + 1, sheet.max_row + 1):
            if sheet.cell(row, summary_column).value not in (None, ""):
                errors.append(f"{sheet.title}!{sheet.cell(row, summary_column).coordinate} is not empty")
            if sheet.cell(row, account_column).value not in (None, ""):
                errors.append(f"{sheet.title}!{sheet.cell(row, account_column).coordinate} is not empty")
            summary = sheet.cell(row, summary_source_column).value
            account = sheet.cell(row, account_source_column).value
            if summary not in (None, ""):
                summary_values.add(str(summary))
            if account not in (None, ""):
                account_values.add(str(account))

    if errors:
        ERROR_PATH.write_text("\n".join(errors) + "\n", encoding="utf-8")
        raise RuntimeError(
            f"Input validation failed with {len(errors)} error(s); see {ERROR_PATH}."
        )

    SUMMARY_LIST_PATH.write_text(
        json.dumps(SUMMARY_KEYWORDS, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    summary_mappings = map_summaries(sorted(summary_values))
    account_mappings = map_accounts(sorted(account_values))

    output = load_workbook(source_path, data_only=False)
    output_sheets = {sheet.title: sheet for sheet in output.worksheets}
    for source_sheet, header_row, headers in sheets:
        sheet = output_sheets[source_sheet.title]
        summary_column = headers["\u6458\u8981mapping"]
        account_column = headers["\u6237\u540dmapping"]
        summary_source_column = headers["\u6458\u8981\u63cf\u8ff0"]
        account_source_column = headers["\u5bf9\u65b9\u6237\u540d"]
        for row in range(header_row + 1, sheet.max_row + 1):
            summary = sheet.cell(row, summary_source_column).value
            account = sheet.cell(row, account_source_column).value
            if summary not in (None, ""):
                sheet.cell(row, summary_column, summary_mappings[str(summary)])
            if account not in (None, ""):
                sheet.cell(row, account_column, account_mappings[str(account)])
    output.save(copy_path)
    print(f"Created: {copy_path}")
    print(f"Summary keywords: {len(SUMMARY_KEYWORDS)}")
    print(f"Unique summaries mapped: {len(summary_mappings)}")
    print(f"Unique account names mapped: {len(account_mappings)}")


def main() -> None:
    try:
        process_workbook(SOURCE_FILE, make_copy_path())
    except Exception as exc:
        log_error("Fatal script error.", exc)
        raise


if __name__ == "__main__":
    main()
