import logging
import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import messagebox, ttk

from automation import LicenseRegistrationBot, NoAvailableSession, RegistrationCancelled
from stations import STATIONS
from storage import load_json, save_json
from validation import validate_form

LICENSE_TYPES = ("普通重型機車", "普通輕型機車 (50cc 以下)", "普通小型車", "職業小型車", "普通大貨車", "職業大貨車", "普通大客車", "職業大客車", "普通聯結車", "職業聯結車")


class LicenseRegistrationForm:
    def __init__(self, root):
        self.root = root
        root.title("駕照報名工具")
        root.geometry("560x720")
        root.resizable(False, False)
        self.data_file = Path(__file__).resolve().parent / "user_info.json"
        self.stop_event = threading.Event()
        self.worker = None
        self.password_visible = False
        self.status_var = tk.StringVar(value="就緒")
        self.keep_browser_var = tk.BooleanVar(value=True)
        self._build_ui()
        self.load_data()
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.grid(sticky="nsew")
        ttk.Label(frame, text="駕照報名表單", font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=(0, 18))
        self.license_type = self._combo(frame, 1, "駕照類型:", LICENSE_TYPES)
        self.name_entry = self._entry(frame, 2, "姓名:")
        self.birth_year, self.birth_month, self.birth_day = self._date_fields(frame, 3, "生日:", True)
        self.phone_entry = self._entry(frame, 4, "電話:")
        self.email_entry = self._entry(frame, 5, "電子郵件:")
        self.id_entry = self._entry(frame, 6, "身分證字號:", "●")
        self.toggle_id_btn = ttk.Button(frame, text="顯示", width=5, command=self.toggle_identity)
        self.toggle_id_btn.grid(row=6, column=2, padx=(5, 0))
        self.exam_year, self.exam_month, self.exam_day = self._date_fields(frame, 7, "考試日期:")
        self.destination_region = self._combo(frame, 8, "監理所區域:", tuple(STATIONS))
        self.destination_region.bind("<<ComboboxSelected>>", self.update_destination_station)
        self.destination_station = self._combo(frame, 9, "監理所:", tuple(next(iter(STATIONS.values()))))
        ttk.Checkbutton(frame, text="完成後保留瀏覽器", variable=self.keep_browser_var).grid(row=10, column=1, sticky="w", pady=(8, 0))
        buttons = ttk.Frame(frame)
        buttons.grid(row=11, column=0, columnspan=3, pady=(24, 12))
        self.save_btn = ttk.Button(buttons, text="儲存資料", command=self.save_data)
        self.clear_btn = ttk.Button(buttons, text="清除", command=self.clear_form)
        self.start_btn = ttk.Button(buttons, text="開始報名", command=self.start_registration)
        self.stop_btn = ttk.Button(buttons, text="停止", command=self.stop_registration, state="disabled")
        for column, button in enumerate((self.save_btn, self.clear_btn, self.start_btn, self.stop_btn)):
            button.grid(row=0, column=column, padx=5)
        ttk.Separator(frame).grid(row=12, column=0, columnspan=3, sticky="ew", pady=8)
        ttk.Label(frame, textvariable=self.status_var, foreground="#1558a6", wraplength=510).grid(row=13, column=0, columnspan=3, sticky="w", pady=4)

    def _entry(self, parent, row, label, show=None):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=6)
        entry = ttk.Entry(parent, width=36, show=show)
        entry.grid(row=row, column=1, pady=6, padx=(10, 0), sticky="w")
        return entry

    def _combo(self, parent, row, label, values):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=6)
        combo = ttk.Combobox(parent, width=34, state="readonly", values=values)
        combo.grid(row=row, column=1, pady=6, padx=(10, 0), sticky="w")
        if values:
            combo.current(0)
        return combo

    def _date_fields(self, parent, row, label, birth=False):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=6)
        holder = ttk.Frame(parent)
        holder.grid(row=row, column=1, pady=6, padx=(10, 0), sticky="w")
        today = date.today()
        years = list(range(today.year, 1919, -1) if birth else range(today.year, today.year + 3))
        widgets = (
            ttk.Combobox(holder, width=7, state="readonly", values=years),
            ttk.Combobox(holder, width=4, state="readonly", values=list(range(1, 13))),
            ttk.Combobox(holder, width=4, state="readonly", values=list(range(1, 32))),
        )
        for index, (widget, suffix) in enumerate(zip(widgets, ("年", "月", "日"), strict=True)):
            widget.grid(row=0, column=index * 2, padx=(0, 3))
            ttk.Label(holder, text=suffix).grid(row=0, column=index * 2 + 1, padx=(0, 5))
        widgets[0].set(today.year - 20 if birth else today.year)
        widgets[1].set(today.month)
        widgets[2].set(today.day)
        return widgets

    def update_destination_station(self, _event=None):
        stations = STATIONS.get(self.destination_region.get(), [])
        self.destination_station.configure(values=stations)
        if stations:
            self.destination_station.current(0)

    def toggle_identity(self):
        self.password_visible = not self.password_visible
        self.id_entry.configure(show="" if self.password_visible else "●")
        self.toggle_id_btn.configure(text="隱藏" if self.password_visible else "顯示")

    def _dates_and_data(self):
        try:
            birthday = date(int(self.birth_year.get()), int(self.birth_month.get()), int(self.birth_day.get()))
            exam = date(int(self.exam_year.get()), int(self.exam_month.get()), int(self.exam_day.get()))
        except ValueError as exc:
            raise ValueError("日期不存在，請重新選擇") from exc
        data = {
            "駕照類型": self.license_type.get(), "姓名": self.name_entry.get().strip(),
            "生日": f"{birthday.year - 1911}{birthday.month:02d}{birthday.day:02d}",
            "電話": self.phone_entry.get().strip(), "電子郵件": self.email_entry.get().strip(),
            "身分證字號": self.id_entry.get().strip().upper(),
            "考試日期": f"{exam.year - 1911}{exam.month:02d}{exam.day:02d}",
            "目的地區": self.destination_region.get(), "目的監理所": self.destination_station.get(),
        }
        errors = validate_form(data, birthday, exam)
        if errors:
            raise ValueError("\n".join(f"• {error}" for error in errors))
        return data

    def _payload(self, data):
        return {"form_data": data, "form_state": {
            "license_type_index": self.license_type.current(), "birth_year": self.birth_year.get(),
            "birth_month": self.birth_month.get(), "birth_day": self.birth_day.get(),
            "exam_year": self.exam_year.get(), "exam_month": self.exam_month.get(), "exam_day": self.exam_day.get(),
            "destination_region_index": self.destination_region.current(), "destination_station_index": self.destination_station.current(),
        }}

    def save_data(self, show_success=True):
        try:
            data = self._dates_and_data()
            save_json(self.data_file, self._payload(data))
        except ValueError as exc:
            messagebox.showwarning("資料有誤", str(exc))
            return False
        self.status_var.set("資料已安全儲存在本機")
        if show_success:
            messagebox.showinfo("儲存成功", "報名資料已儲存。為保護個資，訊息中不顯示完整內容。")
        return True

    def load_data(self):
        try:
            saved = load_json(self.data_file, {})
        except ValueError as exc:
            messagebox.showwarning("載入失敗", str(exc))
            return
        if not saved:
            return
        data, state = saved.get("form_data", {}), saved.get("form_state", {})
        for entry, key in ((self.name_entry, "姓名"), (self.phone_entry, "電話"), (self.email_entry, "電子郵件"), (self.id_entry, "身分證字號")):
            entry.insert(0, data.get(key, ""))
        for widget, key in ((self.license_type, "license_type_index"), (self.destination_region, "destination_region_index")):
            try:
                widget.current(int(state.get(key, 0)))
            except (ValueError, tk.TclError):
                widget.current(0)
        self.update_destination_station()
        for widget, key in ((self.birth_year, "birth_year"), (self.birth_month, "birth_month"), (self.birth_day, "birth_day"), (self.exam_year, "exam_year"), (self.exam_month, "exam_month"), (self.exam_day, "exam_day")):
            if key in state:
                widget.set(state[key])
        try:
            self.destination_station.current(int(state.get("destination_station_index", 0)))
        except (ValueError, tk.TclError):
            self.destination_station.current(0)
        self.status_var.set("已載入上次儲存的資料")

    def clear_form(self):
        for entry in (self.name_entry, self.phone_entry, self.email_entry, self.id_entry):
            entry.delete(0, tk.END)
        today = date.today()
        self.license_type.current(0)
        for widget, value in ((self.birth_year, today.year - 20), (self.birth_month, today.month), (self.birth_day, today.day), (self.exam_year, today.year), (self.exam_month, today.month), (self.exam_day, today.day)):
            widget.set(value)
        self.destination_region.current(0)
        self.update_destination_station()
        self.status_var.set("表單已清除（磁碟中的儲存資料未刪除）")

    def start_registration(self):
        if self.worker and self.worker.is_alive():
            return
        if not self.save_data(show_success=False):
            return
        try:
            data = self._dates_and_data()
        except ValueError:
            return
        self.stop_event.clear()
        self._set_running(True)
        bot = LicenseRegistrationBot(self.stop_event, self.set_status, self.keep_browser_var.get())
        self.worker = threading.Thread(target=self._run_bot, args=(bot, data), daemon=True)
        self.worker.start()

    def _run_bot(self, bot, data):
        try:
            bot.run(data)
        except RegistrationCancelled as exc:
            self.set_status(str(exc))
        except NoAvailableSession as exc:
            self.root.after(0, lambda error=str(exc): messagebox.showinfo("目前無名額", error))
            self.set_status(str(exc))
        except Exception as exc:
            logging.exception("報名流程失敗")
            self.root.after(0, lambda error=str(exc): messagebox.showerror("報名失敗", error))
            self.set_status("報名流程失敗，請查看錯誤訊息")
        finally:
            self.root.after(0, lambda: self._set_running(False))

    def set_status(self, message):
        self.root.after(0, lambda: self.status_var.set(message))

    def stop_registration(self):
        self.stop_event.set()
        self.status_var.set("正在停止…")
        self.stop_btn.configure(state="disabled")

    def _set_running(self, running):
        for button in (self.start_btn, self.save_btn, self.clear_btn):
            button.configure(state="disabled" if running else "normal")
        self.stop_btn.configure(state="normal" if running else "disabled")

    def on_close(self):
        self.stop_event.set()
        self.root.destroy()


def main():
    logging.basicConfig(filename=Path(__file__).resolve().parent / "driver_license.log", level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
    root = tk.Tk()
    LicenseRegistrationForm(root)
    root.mainloop()


if __name__ == "__main__":
    main()
