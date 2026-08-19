import json
import os
from pathlib import Path

from cryptography.fernet import Fernet, InvalidToken


class EncryptedStore:
    def __init__(self, data_dir: Path):
        self.data_dir = data_dir
        self.key_path = data_dir / "secret.key"
        self.data_path = data_dir / "user_info.enc"

    def _fernet(self):
        self.data_dir.mkdir(parents=True, exist_ok=True)
        if not self.key_path.exists():
            temporary = self.key_path.with_suffix(".tmp")
            temporary.write_bytes(Fernet.generate_key())
            os.replace(temporary, self.key_path)
        return Fernet(self.key_path.read_bytes())

    def save(self, data: dict):
        encrypted = self._fernet().encrypt(json.dumps(data, ensure_ascii=False).encode("utf-8"))
        temporary = self.data_path.with_suffix(".tmp")
        temporary.write_bytes(encrypted)
        os.replace(temporary, self.data_path)

    def load(self):
        if not self.data_path.exists():
            return None
        try:
            raw = self._fernet().decrypt(self.data_path.read_bytes())
            return json.loads(raw.decode("utf-8"))
        except (InvalidToken, OSError, json.JSONDecodeError) as exc:
            raise ValueError("本機加密資料無法解密或已損毀") from exc
