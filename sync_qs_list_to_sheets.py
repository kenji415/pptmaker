import csv
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from google.oauth2 import service_account
from googleapiclient.discovery import build


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


@dataclass(frozen=True)
class Tag:
    text: str
    needs_fix: bool
    is_fix_tag: bool
    commenter: str = ""


def normalize_path(p: str) -> str:
    p = (p or "").replace("\\", "/")
    while "//" in p:
        p = p.replace("//", "/")
    return p.strip("/")


def strip_pdf_ext(name: str) -> str:
    return name[:-4] if name.lower().endswith(".pdf") else name


def uniq(items: List[str]) -> List[str]:
    seen = set()
    out = []
    for x in items:
        if x in seen:
            continue
        seen.add(x)
        out.append(x)
    return out


def load_text_mapping_csv(csv_path: Path) -> Dict[str, List[Tag]]:
    """
    Read printviewer text_mapping.csv.
    Expected columns (newer): file_path,juku_name,text_name,commenter,needs_fix,is_fix_tag
    Older formats are handled as best-effort.
    """
    if not csv_path.exists():
        return {}

    out: Dict[str, List[Tag]] = {}
    with csv_path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            file_path = normalize_path((row.get("file_path") or "").strip())
            if not file_path:
                continue
            juku_name = (row.get("juku_name") or "").strip()
            text_name = (row.get("text_name") or "").strip()
            tag_text = juku_name if not text_name else f"{juku_name}:{text_name}"
            if not tag_text:
                continue

            needs_fix_raw = (row.get("needs_fix") or "").strip().lower()
            needs_fix = needs_fix_raw in ("1", "true", "yes", "y", "on")
            is_fix_tag_raw = (row.get("is_fix_tag") or "").strip().lower()
            is_fix_tag = (
                (is_fix_tag_raw in ("1", "true", "yes", "y", "on"))
                if is_fix_tag_raw != ""
                else needs_fix
            )
            commenter = (row.get("commenter") or "").strip()

            out.setdefault(file_path, []).append(
                Tag(text=tag_text, needs_fix=needs_fix, is_fix_tag=is_fix_tag, commenter=commenter)
            )
    return out


def list_pdf_keys(pdf_dir: Path, subject: str = "算数") -> List[Tuple[str, str]]:
    """
    Returns list of (key, filename) where key is like 算数/6年/xxx.pdf.
    """
    subject_dir = pdf_dir / subject
    if not subject_dir.exists():
        return []

    results: List[Tuple[str, str]] = []
    for p in subject_dir.rglob("*.pdf"):
        if not p.is_file():
            continue
        rel = p.relative_to(pdf_dir).as_posix()
        results.append((normalize_path(rel), p.name))
    results.sort(key=lambda x: x[0])
    return results


def build_rows(
    keys_and_names: List[Tuple[str, str]],
    mapping: Dict[str, List[Tag]],
) -> Dict[str, List[str]]:
    """
    key -> [A,B,C,D]
    A: key
    B: title
    C: fix contents (is_fix_tag==True)
    D: other tags (is_fix_tag==False)
    """
    out: Dict[str, List[str]] = {}
    for key, fname in keys_and_names:
        tags = mapping.get(key, [])
        fix = uniq(
            [
                (f"{t.text}({t.commenter})" if t.commenter else t.text)
                for t in tags
                if t.is_fix_tag
            ]
        )
        other = uniq([t.text for t in tags if not t.is_fix_tag])
        title = strip_pdf_ext(fname)
        out[key] = [key, title, "/".join(fix), "/".join(other)]
    return out


def sheet_to_index(values: List[List[str]]) -> Dict[str, int]:
    """
    values: rows from A2:...
    return key->rowIndex (0-based into values)
    """
    out = {}
    for i, row in enumerate(values):
        key = (row[0] if len(row) > 0 else "").strip()
        if key:
            out[key] = i
    return out


def get_sheet_id(service, spreadsheet_id: str, sheet_name: str) -> int:
    meta = (
        service.spreadsheets()
        .get(spreadsheetId=spreadsheet_id)
        .execute()
    )
    for s in meta.get("sheets", []):
        props = s.get("properties", {})
        if props.get("title") == sheet_name:
            return int(props["sheetId"])
    raise RuntimeError(f"Sheet not found: {sheet_name}")


def a1_(sheet_name: str, rng: str) -> str:
    """
    Build A1 notation with a safely quoted sheet name.
    Google Sheets expects quotes when sheet name contains spaces/special chars.
    """
    safe = (sheet_name or "").replace("'", "''")
    return f"'{safe}'!{rng}"


def col_to_a1_(col_1based: int) -> str:
    """1-based column index to A1 letters (A, B, ..., Z, AA, AB, ...)"""
    if col_1based < 1:
        raise ValueError("col must be >= 1")
    n = col_1based
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(ord("A") + rem) + s
    return s


def find_header_columns(service, spreadsheet_id: str, sheet_name: str, header_row: int) -> Dict[str, int]:
    """
    Read header_row and return mapping from header text to 1-based column index.
    Scans A..ZZ (702 cols) which is plenty and still cheap for one row.
    """
    rng = a1_(sheet_name, f"A{header_row}:ZZ{header_row}")
    resp = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=rng).execute()
    row = (resp.get("values") or [[]])[0]
    out: Dict[str, int] = {}
    for i, v in enumerate(row, start=1):
        text = (v or "").strip()
        if text and text not in out:
            out[text] = i
    return out

def sync(
    spreadsheet_id: str,
    sheet_name: str,
    pdf_dir: Path,
    text_mapping_csv: Path,
    creds_json_path: Path,
):
    creds = service_account.Credentials.from_service_account_file(
        str(creds_json_path), scopes=SCOPES
    )
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    mapping = load_text_mapping_csv(text_mapping_csv)
    rows_map = build_rows(list_pdf_keys(pdf_dir, "算数"), mapping)
    current_keys = list(rows_map.keys())
    current_key_set = set(current_keys)

    header_row = int(os.environ.get("QS_HEADER_ROW", "1"))
    data_start_row = header_row + 1  # default: 2

    headers = find_header_columns(service, spreadsheet_id, sheet_name, header_row)
    required = {
        "相対パス": "key",
        "タイトル": "title",
        "修正内容": "fix",
        "その他タグ": "tags",
    }
    missing = [h for h in required.keys() if h not in headers]
    if missing:
        raise SystemExit(f"Header not found in row {header_row}: {', '.join(missing)}")

    col_key = headers["相対パス"]
    col_title = headers["タイトル"]
    col_fix = headers["修正内容"]
    col_tags = headers["その他タグ"]

    # Read existing keys column only (from data_start_row down)
    key_col_letter = col_to_a1_(col_key)
    resp = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=a1_(sheet_name, f"{key_col_letter}{data_start_row}:{key_col_letter}"),
    ).execute()
    key_values: List[List[str]] = resp.get("values", [])
    values = [[(r[0] if r else "")] for r in key_values]
    idx = sheet_to_index(values)

    # Compute deletions (sheet keys missing in current)
    to_delete_row_indices = []
    for k, i in idx.items():
        if k not in current_key_set:
            to_delete_row_indices.append(i)
    # Delete from bottom (convert to sheet row number: +2, and to 0-based for API)
    if to_delete_row_indices:
        sheet_id = get_sheet_id(service, spreadsheet_id, sheet_name)
        requests = []
        for i in sorted(to_delete_row_indices, reverse=True):
            start_index = (data_start_row - 1) + i  # 0-based sheet index
            requests.append(
                {
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": start_index,
                            "endIndex": start_index + 1,
                        }
                    }
                }
            )
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": requests}
        ).execute()

        # refresh after deletions (keys only)
        resp = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=a1_(sheet_name, f"{key_col_letter}{data_start_row}:{key_col_letter}"),
        ).execute()
        key_values = resp.get("values", [])
        values = [[(r[0] if r else "")] for r in key_values]
        idx = sheet_to_index(values)

    # Updates for existing keys (write into each target column; columns may be sparse)
    data_updates = []
    for k, i in idx.items():
        if k in rows_map:
            row_num = data_start_row + i
            a, b, c, d = rows_map[k]
            data_updates.append({"range": a1_(sheet_name, f"{col_to_a1_(col_key)}{row_num}"), "values": [[a]]})
            data_updates.append({"range": a1_(sheet_name, f"{col_to_a1_(col_title)}{row_num}"), "values": [[b]]})
            data_updates.append({"range": a1_(sheet_name, f"{col_to_a1_(col_fix)}{row_num}"), "values": [[c]]})
            data_updates.append({"range": a1_(sheet_name, f"{col_to_a1_(col_tags)}{row_num}"), "values": [[d]]})

    if data_updates:
        # batch in chunks
        for chunk_start in range(0, len(data_updates), 500):
            chunk = data_updates[chunk_start : chunk_start + 500]
            service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    "valueInputOption": "RAW",
                    "data": chunk,
                },
            ).execute()

    # Appends for new keys (in key order)
    to_append = [rows_map[k] for k in current_keys if k not in idx]
    if to_append:
        # Append rows at the bottom by inserting empty rows, then fill target columns.
        # Determine current last row based on key column length
        existing_count = len(values)
        first_new_row = data_start_row + existing_count
        sheet_id = get_sheet_id(service, spreadsheet_id, sheet_name)
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": [
                    {
                        "insertDimension": {
                            "range": {
                                "sheetId": sheet_id,
                                "dimension": "ROWS",
                                "startIndex": first_new_row - 1,
                                "endIndex": first_new_row - 1 + len(to_append),
                            },
                            "inheritFromBefore": True,
                        }
                    }
                ]
            },
        ).execute()

        append_updates = []
        for offset, row in enumerate(to_append):
            row_num = first_new_row + offset
            a, b, c, d = row
            append_updates.append({"range": a1_(sheet_name, f"{col_to_a1_(col_key)}{row_num}"), "values": [[a]]})
            append_updates.append({"range": a1_(sheet_name, f"{col_to_a1_(col_title)}{row_num}"), "values": [[b]]})
            append_updates.append({"range": a1_(sheet_name, f"{col_to_a1_(col_fix)}{row_num}"), "values": [[c]]})
            append_updates.append({"range": a1_(sheet_name, f"{col_to_a1_(col_tags)}{row_num}"), "values": [[d]]})

        for chunk_start in range(0, len(append_updates), 500):
            chunk = append_updates[chunk_start : chunk_start + 500]
            service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    "valueInputOption": "RAW",
                    "data": chunk,
                },
            ).execute()


def main():
    spreadsheet_id = os.environ.get("QS_SPREADSHEET_ID", "").strip()
    sheet_name = os.environ.get("QS_SHEET_NAME", "QSリスト").strip()
    pdf_dir = Path(os.environ.get("QS_PDF_DIR", "")).expanduser()
    creds_json = Path(os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")).expanduser()
    mapping_csv = Path(os.environ.get("QS_TEXT_MAPPING_CSV", "text_mapping.csv")).expanduser()

    if not spreadsheet_id:
        raise SystemExit("QS_SPREADSHEET_ID is required")
    if not pdf_dir.exists():
        raise SystemExit(f"QS_PDF_DIR not found: {pdf_dir}")
    if not creds_json.exists():
        raise SystemExit(f"GOOGLE_SERVICE_ACCOUNT_JSON not found: {creds_json}")

    sync(
        spreadsheet_id=spreadsheet_id,
        sheet_name=sheet_name,
        pdf_dir=pdf_dir,
        text_mapping_csv=mapping_csv if mapping_csv.is_absolute() else (Path(__file__).parent / mapping_csv),
        creds_json_path=creds_json,
    )


if __name__ == "__main__":
    main()

