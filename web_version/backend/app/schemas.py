import re
from datetime import date

from pydantic import BaseModel, EmailStr, Field, field_validator, model_validator

from .config import LICENSE_TYPES, STATIONS

PHONE_RE = re.compile(r"^(?:09\d{8}|0\d{1,2}-?\d{6,8})$")
ID_RE = re.compile(r"^[A-Z][12]\d{8}$")
ID_CODES = {letter: code for letter, code in zip("ABCDEFGHJKLMNPQRSTUVXYWZIO", range(10, 36), strict=True)}


def valid_taiwan_id(value: str) -> bool:
    if not ID_RE.fullmatch(value):
        return False
    code = ID_CODES[value[0]]
    digits = [code // 10, code % 10, *(int(char) for char in value[1:])]
    return sum(number * weight for number, weight in zip(digits, (1, 9, 8, 7, 6, 5, 4, 3, 2, 1, 1), strict=True)) % 10 == 0


class RegistrationData(BaseModel):
    license_type: str
    name: str = Field(min_length=1, max_length=50)
    birth_date: date
    phone: str
    email: EmailStr
    identity_number: str
    exam_date: date
    region: str
    station: str
    keep_browser: bool = True

    @field_validator("name", "phone", "identity_number", "region", "station", mode="before")
    @classmethod
    def strip_text(cls, value):
        return value.strip() if isinstance(value, str) else value

    @field_validator("phone")
    @classmethod
    def check_phone(cls, value):
        if not PHONE_RE.fullmatch(value.replace(" ", "")):
            raise ValueError("電話格式不正確")
        return value

    @field_validator("identity_number")
    @classmethod
    def check_identity(cls, value):
        value = value.upper()
        if not valid_taiwan_id(value):
            raise ValueError("身分證字號格式或檢查碼不正確")
        return value

    @model_validator(mode="after")
    def check_relationships(self):
        if self.license_type not in LICENSE_TYPES:
            raise ValueError("不支援的駕照類型")
        if self.birth_date >= date.today():
            raise ValueError("生日必須早於今天")
        if self.exam_date < date.today():
            raise ValueError("考試日期不可早於今天")
        if self.region not in STATIONS or self.station not in STATIONS[self.region]:
            raise ValueError("監理站與區域不相符")
        return self

    def selenium_payload(self):
        return {
            "駕照類型": self.license_type, "姓名": self.name,
            "生日": f"{self.birth_date.year - 1911}{self.birth_date.month:02d}{self.birth_date.day:02d}",
            "電話": self.phone, "電子郵件": str(self.email), "身分證字號": self.identity_number,
            "考試日期": f"{self.exam_date.year - 1911}{self.exam_date.month:02d}{self.exam_date.day:02d}",
            "目的地區": self.region, "目的監理所": self.station,
        }


class JobStatus(BaseModel):
    state: str = "idle"
    message: str = "就緒"
