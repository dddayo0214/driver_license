from datetime import date, timedelta
from pathlib import Path

from fastapi.testclient import TestClient

from app.main import app
from app.schemas import valid_taiwan_id
from app.secure_store import EncryptedStore

client = TestClient(app)


def test_health():
    assert client.get("/api/health").json() == {"status": "ok"}


def test_options_are_available():
    result = client.get("/api/options")
    assert result.status_code == 200
    assert result.json()["license_types"]
    assert result.json()["stations"]


def test_taiwan_identity_checksum():
    assert valid_taiwan_id("A123456789")
    assert not valid_taiwan_id("A123456788")


def test_rejects_past_exam_date():
    payload = {
        "license_type": "普通重型機車", "name": "測試者", "birth_date": "2000-01-01",
        "phone": "0912345678", "email": "test@example.com", "identity_number": "A123456789",
        "exam_date": str(date.today() - timedelta(days=1)), "region": "臺北市區監理所（含金門馬祖）",
        "station": "士林監理站(臺北市士林區承德路5段80號)", "keep_browser": True,
    }
    assert client.put("/api/profile", json=payload).status_code == 422


def test_encrypted_store_does_not_write_plaintext():
    test_directory = Path(__file__).parent / ".encrypted-test"
    store = EncryptedStore(test_directory)
    try:
        original = {"identity_number": "A123456789", "name": "測試者"}
        store.save(original)
        assert b"A123456789" not in store.data_path.read_bytes()
        assert store.load() == original
    finally:
        store.data_path.unlink(missing_ok=True)
        store.key_path.unlink(missing_ok=True)
        test_directory.rmdir()
