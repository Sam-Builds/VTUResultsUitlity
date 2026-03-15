import io
import re
from datetime import datetime
from pathlib import Path

CODE_RE = re.compile(r"\b1B[A-Z]{2,6}\d{3,}[A-Z]?\b")
RESULT_RE = re.compile(r"\b(P|F|A|W|X|NE)\b", re.IGNORECASE)
USN_RE = re.compile(r"\b1[A-Z]{2}\d{2}[A-Z]{2,3}\d{3}\b")




def _open_pdf(pdf_path: str) -> io.BytesIO:
    import pikepdf
    buf = io.BytesIO()
    try:
        with pikepdf.open(pdf_path) as pdf:
            pdf.save(buf)
    except pikepdf.PasswordError:
        with pikepdf.open(pdf_path, password="") as pdf:
            pdf.save(buf)
    buf.seek(0)
    return buf


def _extract_text(pdf_bytes: io.BytesIO) -> str:
    import pdfplumber
    parts = []
    with pdfplumber.open(pdf_bytes) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=3, y_tolerance=3)
            if text:
                parts.append(text)
    return "\n".join(parts)


def _extract_meta(text: str):
 
    seat_m = re.search(
        r"(?:University\s+Seat\s+Number|Seat\s*No|Register\s*No|USN|Enrollment\s*No)\s*[:\-]?\s*([A-Z0-9]+)",
        text, re.IGNORECASE,
    )
    usn_m = USN_RE.search(text)
    name_m = re.search(
        r"(?:Student\s+Name|Name\s+of\s+Student|Student)\s*[:\-]?\s*([A-Z][A-Za-z .'-]+)",
        text, re.IGNORECASE,
    )
    usn = seat_m.group(1).strip() if seat_m else (usn_m.group(0) if usn_m else "Unknown")
    name = name_m.group(1).strip() if name_m else "Unknown"
    return usn, name


def _infer_marks(nums: list):
    if not nums:
        return "", "", ""
    for i in range(len(nums) - 2):
        a, b, c = nums[i], nums[i + 1], nums[i + 2]
        if c <= 200 and abs((a + b) - c) <= 1:
            return str(a), str(b), str(c)
    # Fallback: last three numbers as-is
    if len(nums) >= 3:
        return str(nums[-3]), str(nums[-2]), str(nums[-1])
    if len(nums) == 2:
        a, b = nums[0], nums[1]
        return str(a), str(b), str(a + b)
    return "", "", str(nums[0])


def _parse_subject_lines(text: str) -> list:
    rows = []
    for line in text.splitlines():
        line = line.strip()
        code_m = CODE_RE.search(line)
        if not code_m:
            continue

        code = code_m.group(0)
        after = line[code_m.end():].strip()

        result_m = RESULT_RE.search(after)
        result = result_m.group(1).upper() if result_m else ""
        if result_m:
            after = after[:result_m.start()] + " " + after[result_m.end():]

        first_num = re.search(r"\d", after)
        if first_num:
            subject_name = after[:first_num.start()].strip(" |:-")
            marks_text = after[first_num.start():]
        else:
            subject_name = after.strip(" |:-")
            marks_text = ""

        subject_name = re.sub(r"\s+", " ", subject_name).strip()
        marks_text = re.sub(r"\b\d{4}-\d{2}-\d{2}\b", " ", marks_text)
        nums = [int(m.group()) for m in re.finditer(r"\b\d{1,3}\b", marks_text)
                if int(m.group()) <= 200]
        internal, external, total = _infer_marks(nums)

        rows.append({
            "Subject Code": code,
            "Subject Name": subject_name,
            "Internal": internal,
            "External": external,
            "Total": total,
            "Result": result,
        })
    return rows


def parse_vtu_pdf(pdf_path: str) -> list:
    pdf_bytes = _open_pdf(pdf_path)
    text = _extract_text(pdf_bytes)
    usn, name = _extract_meta(text)
    subject_rows = _parse_subject_lines(text)

    cols = ["USN", "Name", "Subject Code", "Subject Name", "Internal", "External", "Total", "Result"]
    result = []
    for row in subject_rows:
        row["USN"] = usn
        row["Name"] = name
        result.append({col: row.get(col, "") for col in cols})
    return result


parse_scanned_vtu = parse_vtu_pdf

