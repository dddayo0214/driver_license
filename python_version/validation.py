import re
from datetime import date


EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
PHONE_RE = re.compile(r"^(?:09\d{8}|0\d{1,2}-?\d{6,8})$")
TAIWAN_ID_RE = re.compile(r"^[A-Z][12]\d{8}$")
ID_LETTER_CODES = {
    letter: code
    for letter, code in zip(
        "ABCDEFGHJKLMNPQRSTUVXYWZIO", range(10, 36), strict=True
    )
}


def is_valid_taiwan_id(value: str) -> bool:
    value = value.strip().upper()
    if not TAIWAN_ID_RE.fullmatch(value):
        return False
    code = ID_LETTER_CODES[value[0]]
    digits = [code // 10, code % 10, *(int(char) for char in value[1:])]
    weights = [1, 9, 8, 7, 6, 5, 4, 3, 2, 1, 1]
    return sum(number * weight for number, weight in zip(digits, weights, strict=True)) % 10 == 0


def validate_form(data: dict, birth_date: date, exam_date: date) -> list[str]:
    errors = []
    required = ("姓名", "電話", "電子郵件", "身分證字號", "目的地區", "目的監理所")
    for field in required:
        if not str(data.get(field, "")).strip():
            errors.append(f"請輸入{field}")

    phone = str(data.get("電話", "")).replace(" ", "")
    if phone and not PHONE_RE.fullmatch(phone):
        errors.append("電話格式不正確")
    email = str(data.get("電子郵件", "")).strip()
    if email and not EMAIL_RE.fullmatch(email):
        errors.append("電子郵件格式不正確")
    identity = str(data.get("身分證字號", "")).strip()
    if identity and not is_valid_taiwan_id(identity):
        errors.append("身分證字號格式或檢查碼不正確")
    if birth_date >= date.today():
        errors.append("生日必須早於今天")
    if exam_date < date.today():
        errors.append("考試日期不可早於今天")
    return errors
