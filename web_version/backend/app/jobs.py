import logging
import threading

from .automation import NoAvailableSession, RegistrationBot, RegistrationCancelled
from .schemas import JobStatus, RegistrationData

LOGGER = logging.getLogger(__name__)


class JobManager:
    def __init__(self):
        self._lock = threading.Lock()
        self._stop_event = threading.Event()
        self._thread = None
        self._status = JobStatus()

    def status(self):
        with self._lock:
            return self._status.model_copy()

    def _update(self, state, message):
        with self._lock:
            self._status = JobStatus(state=state, message=message)

    def start(self, registration: RegistrationData):
        with self._lock:
            if self._thread and self._thread.is_alive():
                return False
            self._stop_event.clear()
            self._status = JobStatus(state="running", message="正在準備報名…")
            self._thread = threading.Thread(target=self._run, args=(registration,), daemon=True)
            self._thread.start()
            return True

    def stop(self):
        with self._lock:
            if not self._thread or not self._thread.is_alive():
                return False
            self._stop_event.set()
            self._status = JobStatus(state="stopping", message="正在停止…")
            return True

    def _run(self, registration):
        bot = RegistrationBot(self._stop_event, lambda message: self._update("running", message), registration.keep_browser)
        try:
            bot.run(registration.selenium_payload())
            self._update("success", "資料已送出，請在瀏覽器確認結果")
        except RegistrationCancelled as exc:
            self._update("cancelled", str(exc))
        except NoAvailableSession as exc:
            self._update("unavailable", str(exc))
        except Exception as exc:
            LOGGER.exception("Registration failed")
            self._update("error", str(exc))
