import csv
import re
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import clean_claude_comments2 as base


RAW_PATH = Path(r"D:\AI\comment\test\input.txt")
OUT_DIR = Path(r"D:\AI\comment\test")
CSV_PATH = OUT_DIR / "test_cleaned_facebook_comments_structured.csv"
XLSX_PATH = OUT_DIR / "test_cleaned_facebook_comments_structured.xlsx"
SUMMARY_PATH = OUT_DIR / "test_cleaned_facebook_comments_structured_summary.txt"

COLUMNS = base.COLUMNS


FACEBOOK_METADATA_LINES = {
    "·",
    "•",
    "ผู้เขียน",
    "แฟนตัวยง",
    "top fan",
    "author",
}

FACEBOOK_UI_NOISE = {
    "like",
    "reply",
    "share",
    "edited",
    "see more",
    "ตอบกลับ",
    "แชร์",
    "ถูกใจ",
    "แก้ไขแล้ว",
    "ดูเพิ่มเติม",
}


def normalize_line(line: str) -> str:
    return base.normalize_line(line).replace("\xa0", " ").strip()


def is_facebook_metadata(line: str) -> bool:
    low = normalize_line(line).lower()
    return low in FACEBOOK_METADATA_LINES or bool(re.fullmatch(r"[·•]+", low))


def is_facebook_time_line(line: str) -> bool:
    low = re.sub(r"\s+", " ", normalize_line(line).lower())
    if low in {"just now", "yesterday"}:
        return True
    if re.match(r"^\d+\s*(?:s|m|h|d|w|mo|y|yr|yrs)$", low):
        return True
    if re.match(
        r"^\d+\s*(?:sec|secs|second|seconds|min|mins|minute|minutes|hr|hrs|hour|hours|day|days|week|weeks|month|months|year|years)(?: ago)?$",
        low,
    ):
        return True
    return bool(
        re.match(
            r"^\d+\s*(?:วินาที|นาที|ชม\.?|ชั่วโมง|วัน|สัปดาห์|เดือน|ปี)(?:ที่แล้ว)?$",
            low,
        )
    )


def is_facebook_ui_noise(line: str) -> bool:
    low = re.sub(r"\s+", " ", normalize_line(line).lower())
    if not low:
        return True
    if is_facebook_metadata(low) or low in FACEBOOK_UI_NOISE:
        return True
    checks = [
        r"^\d+$",
        r"^see \d+ repl(?:y|ies)$",
        r"^view \d+ repl(?:y|ies)$",
        r"^\d+\s+repl(?:y|ies)$",
        r"^ดูคำตอบเพิ่มเติม.*",
    ]
    return any(re.search(pattern, low) for pattern in checks)


def find_timestamp(lines, start_index: int, max_scan: int = 80):
    stop = min(len(lines), start_index + max_scan)
    for index in range(start_index, stop):
        if is_facebook_time_line(lines[index]):
            return index
    return None


def clean_comment_text(body_lines):
    body_lines = [line for line in body_lines if not is_facebook_ui_noise(line)]
    return re.sub(r"\s+", " ", " ".join(body_lines)).strip()


def extract_comments(text: str):
    """Parse Facebook copied comments without changing the shared classifier.

    Facebook clipboard text usually appears as author, optional badge/separator,
    comment body, timestamp, then action labels such as Reply/Share. This parser
    uses the timestamp as the end marker for each comment and skips UI labels.
    """
    lines = [normalize_line(line) for line in text.splitlines()]
    lines = [line for line in lines if line]

    comments = []
    index = 0
    while index < len(lines):
        if is_facebook_ui_noise(lines[index]) or is_facebook_time_line(lines[index]):
            index += 1
            continue

        author = lines[index]
        body_start = index + 1
        while body_start < len(lines) and is_facebook_metadata(lines[body_start]):
            body_start += 1

        timestamp_index = find_timestamp(lines, body_start)
        if timestamp_index is None:
            index += 1
            continue

        comment_text = clean_comment_text(lines[body_start:timestamp_index])
        if comment_text:
            comments.append(
                {
                    "author": author,
                    "timestamp": lines[timestamp_index],
                    "text": comment_text,
                }
            )

        index = timestamp_index + 1
        while index < len(lines) and is_facebook_ui_noise(lines[index]):
            index += 1

    return comments


def write_csv(rows):
    with CSV_PATH.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def write_xlsx(rows):
    original_path = base.XLSX_PATH
    try:
        base.XLSX_PATH = XLSX_PATH
        base.write_xlsx(rows)
    finally:
        base.XLSX_PATH = original_path


def write_summary(comments, rows):
    original_path = base.SUMMARY_PATH
    try:
        base.SUMMARY_PATH = SUMMARY_PATH
        base.write_summary(comments, rows)
    finally:
        base.SUMMARY_PATH = original_path


def main():
    base.validate_category_schema()
    raw = RAW_PATH.read_text(encoding="utf-8-sig")
    comments = extract_comments(raw)
    rows = [base.classify_comment(comment) for comment in comments]
    write_csv(rows)
    write_xlsx(rows)
    write_summary(comments, rows)
    print(f"Parsed Facebook comments: {len(comments)}")
    print(f"Wrote: {CSV_PATH}")
    print(f"Wrote: {XLSX_PATH}")
    print(f"Wrote: {SUMMARY_PATH}")


if __name__ == "__main__":
    main()
