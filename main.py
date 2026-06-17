# main.py — основной GUI файл, запускается 64-bit Python
# Требования: pip install (ничего дополнительно — tkinter встроен)

import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import subprocess
import json
import os
import threading
from datetime import datetime

# ─── Пути ─────────────────────────────────────────────────────────────────────

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

WORKER_SCRIPT = os.path.join(BASE_DIR, "notes_worker.py")
CONFIG_FILE   = os.path.join(BASE_DIR, "config.json")

# ─── Конфигурация (сохраняется в config.json рядом с .exe) ────────────────────

DEFAULT_CONFIG = {
    "python32_path": "",
    "server_name":   "",
    "user_name":     "",
}

def load_config():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                # Добавляем недостающие ключи из DEFAULT_CONFIG
                for k, v in DEFAULT_CONFIG.items():
                    data.setdefault(k, v)
                return data
    except Exception:
        pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# ─── Автопоиск 32-битного Python ──────────────────────────────────────────────

PYTHON32_SEARCH_PATHS = [
    r"C:\Python313-32\python.exe",
    r"C:\Python312-32\python.exe",
    r"C:\Python311-32\python.exe",
    r"C:\Python310-32\python.exe",
    r"C:\Python39-32\python.exe",
    r"C:\Python38-32\python.exe",
    r"C:\Python313_32\python.exe",
    r"C:\Python312_32\python.exe",
    r"C:\Python311_32\python.exe",
    r"C:\Python310_32\python.exe",
    r"C:\Python\Python313-32\python.exe",
    r"C:\Python\Python312-32\python.exe",
    r"C:\Program Files (x86)\Python313\python.exe",
    r"C:\Program Files (x86)\Python312\python.exe",
    r"C:\Program Files (x86)\Python311\python.exe",
    r"C:\Program Files (x86)\Python310\python.exe",
    r"C:\Program Files (x86)\Python39\python.exe",
    r"C:\Program Files (x86)\Python38\python.exe",
    r"C:\Users\{user}\AppData\Local\Programs\Python\Python313-32\python.exe",
    r"C:\Users\{user}\AppData\Local\Programs\Python\Python312-32\python.exe",
    r"C:\Users\{user}\AppData\Local\Programs\Python\Python311-32\python.exe",
]

def find_python32():
    """Ищет 32-битный Python в стандартных местах."""
    username = os.environ.get("USERNAME", "")
    for path in PYTHON32_SEARCH_PATHS:
        path = path.replace("{user}", username)
        if os.path.exists(path):
            # Проверяем что это действительно 32-bit
            try:
                result = subprocess.run(
                    [path, "-c", "import sys; print(sys.maxsize)"],
                    capture_output=True, text=True, timeout=5
                )
                if result.stdout.strip() == "2147483647":
                    return path
            except Exception:
                continue
    return ""

# ─── Логика вызова воркера ────────────────────────────────────────────────────

def call_worker(python32_path, server_name, user_name, notes_password, expiration_days=365, action="renew"):
    """Запускает notes_worker.py через 32-битный Python и возвращает результат."""
    if not python32_path or not os.path.exists(python32_path):
        return {
            "success": False,
            "message": (
                f"32-битный Python не найден:\n{python32_path}\n\n"
                f"Укажите путь вручную в поле 'Python 32-bit' или нажмите кнопку '...' для выбора."
            ),
            "data": {}
        }

    if not os.path.exists(WORKER_SCRIPT):
        return {
            "success": False,
            "message": (
                f"Файл notes_worker.py не найден:\n{WORKER_SCRIPT}\n\n"
                f"Скопируйте notes_worker.py рядом с программой."
            ),
            "data": {}
        }

    params = json.dumps({
        "action":          action,
        "server_name":     server_name,
        "user_name":       user_name,
        "notes_password":  notes_password,
        "expiration_days": expiration_days,
    }, ensure_ascii=False)

    try:
        proc = subprocess.run(
            [python32_path, WORKER_SCRIPT],
            input=params.encode("utf-8"),
            capture_output=True,
            timeout=60
        )

        def decode(b):
            if not b:
                return ""
            for enc in ("utf-8", "cp1251", "cp866", "latin-1"):
                try:
                    return b.decode(enc).strip()
                except Exception:
                    continue
            return b.decode("utf-8", errors="replace").strip()

        stderr_text = decode(proc.stderr)
        stdout_text = decode(proc.stdout)

        if proc.returncode != 0:
            detail = stderr_text or stdout_text or "нет вывода от воркера"
            return {
                "success": False,
                "message": f"Воркер завершился с ошибкой (код {proc.returncode}):\n{detail}",
                "data": {}
            }

        if not stdout_text:
            detail = stderr_text or "воркер не вернул данные"
            return {
                "success": False,
                "message": f"Воркер запустился но ничего не вернул:\n{detail}",
                "data": {}
            }

        return json.loads(stdout_text)

    except subprocess.TimeoutExpired:
        return {"success": False, "message": "Превышено время ожидания (60 сек).", "data": {}}
    except json.JSONDecodeError as e:
        return {"success": False, "message": f"Некорректный ответ:\n{stdout_text}\n{e}", "data": {}}
    except Exception as e:
        return {"success": False, "message": str(e), "data": {}}

# ─── GUI ──────────────────────────────────────────────────────────────────────

class LotusRenewApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.config_data = load_config()
        self.title("Lotus Notes — Продление сертификата")
        self.resizable(False, False)
        self._center_window(580, 650)
        self._build_ui()
        self._apply_config()

        # Если путь к Python32 не сохранён — ищем автоматически
        if not self.var_python32.get():
            self._auto_find_python32()

    def _center_window(self, w, h):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _apply_config(self):
        self.var_python32.set(self.config_data.get("python32_path", ""))
        self.var_server.set(self.config_data.get("server_name", ""))
        self.var_user.set(self.config_data.get("user_name", ""))

    def _save_config(self):
        self.config_data["python32_path"] = self.var_python32.get().strip()
        self.config_data["server_name"]   = self.var_server.get().strip()
        self.config_data["user_name"]     = self.var_user.get().strip()
        save_config(self.config_data)

    def _auto_find_python32(self):
        self._log("Поиск 32-битного Python...", "warn")
        def search():
            path = find_python32()
            self.after(0, self._on_python32_found, path)
        threading.Thread(target=search, daemon=True).start()

    def _on_python32_found(self, path):
        if path:
            self.var_python32.set(path)
            self._log(f"Python 32-bit найден: {path}", "success")
            self._save_config()
        else:
            self._log("Python 32-bit не найден автоматически — укажите путь вручную.", "error")

    # ── Построение интерфейса ─────────────────────────────────────────────────

    def _build_ui(self):

        # ── Заголовок ─────────────────────────────────────────────────────────
        header = tk.Frame(self, bg="#003366")
        header.pack(fill="x")
        tk.Label(
            header,
            text="Lotus Notes  •  Продление сертификата пользователя",
            bg="#003366", fg="white",
            font=("Segoe UI", 11, "bold"),
            pady=10
        ).pack()

        # ── Настройки Python32 ────────────────────────────────────────────────
        py_frame = tk.LabelFrame(self, text=" Python 32-bit ", font=("Segoe UI", 9, "bold"), padx=10, pady=6)
        py_frame.pack(fill="x", padx=14, pady=(10, 4))

        self.var_python32 = tk.StringVar()
        py_entry = tk.Entry(py_frame, textvariable=self.var_python32, width=44, font=("Segoe UI", 9))
        py_entry.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        tk.Button(
            py_frame, text="...", width=3,
            command=self._browse_python32,
            font=("Segoe UI", 9), relief="flat", bg="#e0e0e0", cursor="hand2"
        ).grid(row=0, column=1)

        tk.Button(
            py_frame, text="Найти автоматически",
            command=self._auto_find_python32,
            font=("Segoe UI", 8), relief="flat", bg="#e8f0fe", cursor="hand2",
            padx=6
        ).grid(row=0, column=2, padx=(4, 0))

        py_frame.columnconfigure(0, weight=1)

        # ── Форма подключения ─────────────────────────────────────────────────
        form = tk.LabelFrame(self, text=" Параметры подключения ", font=("Segoe UI", 9, "bold"), padx=10, pady=8)
        form.pack(fill="x", padx=14, pady=4)

        # Сервер
        tk.Label(form, text="Сервер Domino:", anchor="w", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=4)
        self.var_server = tk.StringVar()
        tk.Entry(form, textvariable=self.var_server, width=38, font=("Segoe UI", 9)).grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=4)
        tk.Label(form, text='например: DominoServer/MyOrg', fg="gray", font=("Segoe UI", 8)).grid(row=1, column=1, sticky="w", padx=(8, 0))

        # Пользователь
        tk.Label(form, text="Пользователь:", anchor="w", font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w", pady=4)
        self.var_user = tk.StringVar()
        tk.Entry(form, textvariable=self.var_user, width=38, font=("Segoe UI", 9)).grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=4)
        tk.Label(form, text='например: Ivan Ivanov/Dept/MyOrg', fg="gray", font=("Segoe UI", 8)).grid(row=3, column=1, sticky="w", padx=(8, 0))

        # Пароль
        tk.Label(form, text="Пароль Notes ID:", anchor="w", font=("Segoe UI", 9)).grid(row=4, column=0, sticky="w", pady=4)
        self.var_password = tk.StringVar()
        self.entry_password = tk.Entry(form, textvariable=self.var_password, width=38, show="•", font=("Segoe UI", 9))
        self.entry_password.grid(row=4, column=1, sticky="ew", padx=(8, 0), pady=4)

        self.show_pass = tk.BooleanVar(value=False)
        tk.Checkbutton(
            form, text="Показать пароль", variable=self.show_pass,
            command=self._toggle_password, font=("Segoe UI", 8), fg="gray"
        ).grid(row=5, column=1, sticky="w", padx=(8, 0))

        form.columnconfigure(1, weight=1)

        # ── Срок продления ────────────────────────────────────────────────────
        exp_frame = tk.LabelFrame(self, text=" Срок продления ", font=("Segoe UI", 9, "bold"), padx=10, pady=8)
        exp_frame.pack(fill="x", padx=14, pady=4)

        self.var_days = tk.IntVar(value=365)
        options = [("6 месяцев", 180), ("1 год", 365), ("2 года", 730), ("3 года", 1095)]

        rb_frame = tk.Frame(exp_frame)
        rb_frame.pack(side="left")
        for label, days in options:
            tk.Radiobutton(rb_frame, text=label, variable=self.var_days, value=days, font=("Segoe UI", 9)).pack(side="left", padx=6)

        tk.Label(exp_frame, text="  или дней:", font=("Segoe UI", 9)).pack(side="left")
        self.var_custom_days = tk.StringVar()
        tk.Entry(exp_frame, textvariable=self.var_custom_days, width=6, font=("Segoe UI", 9)).pack(side="left", padx=4)

        # ── Кнопки ────────────────────────────────────────────────────────────
        btn_frame = tk.Frame(self)
        btn_frame.pack(fill="x", padx=14, pady=(8, 4))

        self.btn_run = tk.Button(
            btn_frame, text="▶  Продлить сертификат",
            command=self._on_run,
            bg="#0066CC", fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2", padx=16, pady=6
        )
        self.btn_run.pack(side="left")

        self.btn_check = tk.Button(
            btn_frame, text="🔍  Проверить сертификат",
            command=self._on_check,
            bg="#28a745", fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2", padx=16, pady=6
        )
        self.btn_check.pack(side="left", padx=(8, 0))

        tk.Button(
            btn_frame, text="Очистить лог",
            command=self._clear_log,
            font=("Segoe UI", 9), relief="flat",
            cursor="hand2", padx=10, pady=6, bg="#e0e0e0"
        ).pack(side="right")

        # ── Панель статуса сертификата ────────────────────────────────────────
        self.status_frame = tk.LabelFrame(self, text=" Статус сертификата ", font=("Segoe UI", 9, "bold"), padx=10, pady=8)
        self.status_frame.pack(fill="x", padx=14, pady=(0, 4))

        # Строки статуса
        rows = [
            ("Пользователь:",  "lbl_fullname"),
            ("Дата выдачи:",   "lbl_issued"),
            ("Истекает:",      "lbl_expdate"),
            ("Статус:",        "lbl_status"),
            ("Осталось дней:", "lbl_days"),
        ]
        for i, (title, attr) in enumerate(rows):
            tk.Label(self.status_frame, text=title, anchor="w",
                     font=("Segoe UI", 9), width=16).grid(row=i, column=0, sticky="w", pady=2)
            lbl = tk.Label(self.status_frame, text="—", anchor="w",
                           font=("Segoe UI", 9, "bold"), fg="gray")
            lbl.grid(row=i, column=1, sticky="w", padx=(8, 0), pady=2)
            setattr(self, attr, lbl)

        self.status_frame.columnconfigure(1, weight=1)

        # ── Прогресс ──────────────────────────────────────────────────────────
        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=14, pady=(0, 4))

        # ── Лог ───────────────────────────────────────────────────────────────
        log_frame = tk.LabelFrame(self, text=" Лог выполнения ", font=("Segoe UI", 9, "bold"), padx=6, pady=6)
        log_frame.pack(fill="both", expand=True, padx=14, pady=(4, 12))

        self.log = scrolledtext.ScrolledText(
            log_frame, height=10, state="disabled",
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white", relief="flat"
        )
        self.log.pack(fill="both", expand=True)

        self.log.tag_config("info",    foreground="#9cdcfe")
        self.log.tag_config("success", foreground="#4ec9b0")
        self.log.tag_config("error",   foreground="#f44747")
        self.log.tag_config("warn",    foreground="#dcdcaa")
        self.log.tag_config("time",    foreground="#569cd6")

        self._log("Готов к работе. Заполните параметры и нажмите «Продлить сертификат».", "info")

    # ── Вспомогательные методы ────────────────────────────────────────────────

    def _browse_python32(self):
        path = filedialog.askopenfilename(
            title="Выберите python.exe (32-bit)",
            filetypes=[("Python", "python.exe"), ("Все файлы", "*.*")],
            initialdir="C:\\"
        )
        if path:
            self.var_python32.set(path)
            self._save_config()
            self._log(f"Python 32-bit установлен: {path}", "success")

    def _toggle_password(self):
        self.entry_password.config(show="" if self.show_pass.get() else "•")

    def _log(self, text, tag="info"):
        self.log.config(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.insert("end", f"[{ts}] ", "time")
        self.log.insert("end", text + "\n", tag)
        self.log.see("end")
        self.log.config(state="disabled")

    def _clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    def _set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        self.btn_run.config(state=state)
        self.btn_check.config(state=state)
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()

    # ── Проверка сертификата ──────────────────────────────────────────────────

    def _on_check(self):
        python32 = self.var_python32.get().strip()
        server   = self.var_server.get().strip()
        user     = self.var_user.get().strip()
        password = self.var_password.get()

        if not python32:
            messagebox.showerror("Ошибка", "Укажите путь к 32-битному Python")
            return
        if not server:
            messagebox.showerror("Ошибка", "Укажите имя сервера Domino")
            return
        if not user:
            messagebox.showerror("Ошибка", "Укажите имя пользователя")
            return

        self._save_config()
        self._set_busy(True)
        self._log("Проверка сертификата...", "warn")
        self._log(f"Пользователь : {user}", "info")
        self._log(f"Сервер       : {server}", "info")

        thread = threading.Thread(
            target=self._run_worker,
            args=(python32, server, user, password, 365, "check"),
            daemon=True
        )
        thread.start()

    def _update_status_panel(self, data):
        """Обновляет панель статуса сертификата."""
        self.lbl_fullname.config(text=data.get("full_name", "—"), fg="#222222")

        self.lbl_issued.config(
            text=data.get("issued_date", "Нет данных"), fg="#555555"
        )

        exp = data.get("expiration_date", "—")
        self.lbl_expdate.config(text=exp, fg="#222222")

        days = data.get("days_left")
        if days is not None:
            self.lbl_days.config(text=f"{days} дней", fg="#222222")
        else:
            self.lbl_days.config(text="—", fg="gray")

        status     = data.get("status", "unknown")
        status_txt = data.get("status_text", "—")

        color_map = {
            "ok":      "#28a745",   # зелёный
            "warning": "#e67e00",   # оранжевый
            "expired": "#dc3545",   # красный
            "unknown": "#888888",   # серый
        }
        self.lbl_status.config(
            text=status_txt,
            fg=color_map.get(status, "#888888")
        )

    # ── Запуск продления ─────────────────────────────────────────────────────

    def _on_run(self):
        python32 = self.var_python32.get().strip()
        server   = self.var_server.get().strip()
        user     = self.var_user.get().strip()
        password = self.var_password.get()

        custom = self.var_custom_days.get().strip()
        if custom:
            if not custom.isdigit() or int(custom) <= 0:
                messagebox.showerror("Ошибка", "Введите корректное количество дней (целое число > 0)")
                return
            days = int(custom)
        else:
            days = self.var_days.get()

        if not python32:
            messagebox.showerror("Ошибка", "Укажите путь к 32-битному Python")
            return
        if not server:
            messagebox.showerror("Ошибка", "Укажите имя сервера Domino")
            return
        if not user:
            messagebox.showerror("Ошибка", "Укажите имя пользователя")
            return

        # Сохраняем настройки
        self._save_config()

        self._set_busy(True)
        self._log("Запуск продления сертификата...", "warn")
        self._log(f"Python 32-bit : {python32}", "info")
        self._log(f"Сервер        : {server}", "info")
        self._log(f"Пользователь  : {user}", "info")
        self._log(f"Срок          : {days} дней", "info")

        thread = threading.Thread(
            target=self._run_worker,
            args=(python32, server, user, password, days, "renew"),
            daemon=True
        )
        thread.start()

    def _run_worker(self, python32, server, user, password, days, action="renew"):
        result = call_worker(python32, server, user, password, days, action)
        self.after(0, self._on_result, result, action)

    def _on_result(self, result, action="renew"):
        self._set_busy(False)

        if result["success"]:
            data = result.get("data", {})
            self._log("─" * 45, "time")

            if action == "check":
                # ── Результат проверки ────────────────────────────────────────
                self._update_status_panel(data)
                status     = data.get("status", "unknown")
                status_txt = data.get("status_text", "")
                exp_date   = data.get("expiration_date", "—")
                days_left  = data.get("days_left")

                tag = "success" if status == "ok" else ("warn" if status == "warning" else "error")
                self._log(f"Пользователь  : {data.get('full_name', '')}", "info")
                self._log(f"Истекает      : {exp_date}", "info")
                if days_left is not None:
                    self._log(f"Осталось      : {days_left} дней", "info")
                self._log(f"Статус        : {status_txt}", tag)

                # Показываем найденные поля для диагностики если дата не найдена
                if status == "unknown":
                    debug = data.get("debug_fields", {})
                    if debug:
                        self._log("── Диагностика: найденные поля ──", "warn")
                        for fname, fval in debug.items():
                            self._log(f"  {fname} = {fval}", "warn")
                    else:
                        self._log("Диагностика: поля с датами не найдены в документе", "error")
                elif "expiration_field" in data:
                    self._log(f"Поле даты     : {data['expiration_field']}", "info")

            else:
                # ── Результат продления ───────────────────────────────────────
                self._log("УСПЕШНО: " + result["message"], "success")
                if "full_name" in data:
                    self._log(f"Пользователь  : {data['full_name']}", "info")
                if "current_expiration" in data:
                    self._log(f"Было          : {data['current_expiration']}", "info")
                if "new_expiration" in data:
                    self._log(f"Станет        : {data['new_expiration']}", "success")
                if "admin_user" in data:
                    self._log(f"Выполнил      : {data['admin_user']}", "info")
                self._log("AdminP обработает запрос при следующем запуске на сервере.", "warn")

                messagebox.showinfo(
                    "Готово",
                    f"Запрос на продление создан!\n\n"
                    f"Пользователь: {data.get('full_name', '')}\n"
                    f"Новая дата истечения: {data.get('new_expiration', 'н/д')}\n\n"
                    f"AdminP обработает запрос при следующем запуске."
                )

            self._log("─" * 45, "time")

        else:
            self._log("─" * 45, "time")
            self._log("ОШИБКА: " + result["message"], "error")
            self._log("─" * 45, "time")
            messagebox.showerror("Ошибка", result["message"])


# ─── Запуск ───────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = LotusRenewApp()
    app.mainloop()
