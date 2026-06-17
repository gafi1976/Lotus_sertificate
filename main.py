# main.py — основной GUI файл, запускается 64-bit Python
# Требования: pip install (ничего дополнительно — tkinter встроен)

import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import subprocess
import json
import os
import threading
from datetime import datetime

# ─── Настройки ────────────────────────────────────────────────────────────────

PYTHON32_PATH = r"C:\Python313-32\python.exe"

# Корректный путь и в обычном режиме и после компиляции PyInstaller
if getattr(sys, "frozen", False):
    # Запущено как .exe — файлы лежат в _internal рядом с .exe
    BASE_DIR = os.path.dirname(sys.executable)
    WORKER_SCRIPT = os.path.join(BASE_DIR, "_internal", "notes_worker.py")
else:
    # Обычный запуск через python main.py
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    WORKER_SCRIPT = os.path.join(BASE_DIR, "notes_worker.py")

# ─── Логика вызова воркера ────────────────────────────────────────────────────

def call_worker(server_name, user_name, notes_password, expiration_days):
    """Запускает notes_worker.py через 32-битный Python и возвращает результат."""
    if not os.path.exists(PYTHON32_PATH):
        return {
            "success": False,
            "message": f"32-битный Python не найден:\n{PYTHON32_PATH}",
            "data": {}
        }

    if not os.path.exists(WORKER_SCRIPT):
        return {
            "success": False,
            "message": f"Файл воркера не найден:\n{WORKER_SCRIPT}",
            "data": {}
        }

    params = json.dumps({
        "server_name":     server_name,
        "user_name":       user_name,
        "notes_password":  notes_password,
        "expiration_days": expiration_days,
    }, ensure_ascii=False)

    try:
        proc = subprocess.run(
            [PYTHON32_PATH, WORKER_SCRIPT],
            input=params,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=60
        )

        if proc.returncode != 0:
            return {
                "success": False,
                "message": f"Воркер завершился с ошибкой (код {proc.returncode}):\n{proc.stderr}",
                "data": {}
            }

        return json.loads(proc.stdout)

    except subprocess.TimeoutExpired:
        return {"success": False, "message": "Превышено время ожидания (60 сек). Проверьте подключение к серверу.", "data": {}}
    except json.JSONDecodeError:
        return {"success": False, "message": f"Некорректный ответ от воркера:\n{proc.stdout}", "data": {}}
    except Exception as e:
        return {"success": False, "message": str(e), "data": {}}

# ─── GUI ──────────────────────────────────────────────────────────────────────

class LotusRenewApp(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title("Lotus Notes — Продление сертификата")
        self.resizable(False, False)
        self._center_window(560, 580)
        self._build_ui()

    def _center_window(self, w, h):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

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

        # ── Форма ─────────────────────────────────────────────────────────────
        form = tk.LabelFrame(self, text=" Параметры подключения ", font=("Segoe UI", 9, "bold"), padx=10, pady=8)
        form.pack(fill="x", padx=14, pady=(12, 4))

        # Сервер
        tk.Label(form, text="Сервер Domino:", anchor="w", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=4)
        self.var_server = tk.StringVar()
        tk.Entry(form, textvariable=self.var_server, width=38, font=("Segoe UI", 9)).grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=4)

        # Подсказка для сервера
        tk.Label(form, text='например: DominoServer/MyOrg', fg="gray", font=("Segoe UI", 8)).grid(row=1, column=1, sticky="w", padx=(8, 0))

        # Пользователь
        tk.Label(form, text="Пользователь:", anchor="w", font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w", pady=4)
        self.var_user = tk.StringVar()
        tk.Entry(form, textvariable=self.var_user, width=38, font=("Segoe UI", 9)).grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=4)

        # Подсказка для пользователя
        tk.Label(form, text='например: Ivan Ivanov/Dept/MyOrg', fg="gray", font=("Segoe UI", 8)).grid(row=3, column=1, sticky="w", padx=(8, 0))

        # Пароль Notes ID
        tk.Label(form, text="Пароль Notes ID:", anchor="w", font=("Segoe UI", 9)).grid(row=4, column=0, sticky="w", pady=4)
        self.var_password = tk.StringVar()
        self.entry_password = tk.Entry(form, textvariable=self.var_password, width=38, show="•", font=("Segoe UI", 9))
        self.entry_password.grid(row=4, column=1, sticky="ew", padx=(8, 0), pady=4)

        # Показать/скрыть пароль
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

        options = [
            ("6 месяцев",  180),
            ("1 год",      365),
            ("2 года",     730),
            ("3 года",    1095),
        ]

        btn_frame = tk.Frame(exp_frame)
        btn_frame.pack(side="left")

        for label, days in options:
            tk.Radiobutton(
                btn_frame, text=label, variable=self.var_days, value=days,
                font=("Segoe UI", 9)
            ).pack(side="left", padx=6)

        # Своё значение
        tk.Label(exp_frame, text="  или дней:", font=("Segoe UI", 9)).pack(side="left")
        self.var_custom_days = tk.StringVar()
        tk.Entry(exp_frame, textvariable=self.var_custom_days, width=6, font=("Segoe UI", 9)).pack(side="left", padx=4)

        # ── Кнопки ────────────────────────────────────────────────────────────
        btn_frame2 = tk.Frame(self)
        btn_frame2.pack(fill="x", padx=14, pady=(8, 4))

        self.btn_run = tk.Button(
            btn_frame2,
            text="▶  Продлить сертификат",
            command=self._on_run,
            bg="#0066CC", fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2",
            padx=16, pady=6
        )
        self.btn_run.pack(side="left")

        tk.Button(
            btn_frame2,
            text="Очистить лог",
            command=self._clear_log,
            font=("Segoe UI", 9),
            relief="flat", cursor="hand2",
            padx=10, pady=6, bg="#e0e0e0"
        ).pack(side="right")

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

        # Теги цветов для лога
        self.log.tag_config("info",    foreground="#9cdcfe")
        self.log.tag_config("success", foreground="#4ec9b0")
        self.log.tag_config("error",   foreground="#f44747")
        self.log.tag_config("warn",    foreground="#dcdcaa")
        self.log.tag_config("time",    foreground="#569cd6")

        self._log("Готов к работе. Заполните параметры и нажмите «Продлить сертификат».", "info")

    # ── Вспомогательные методы ────────────────────────────────────────────────

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
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()

    # ── Обработка нажатия "Продлить" ─────────────────────────────────────────

    def _on_run(self):
        server   = self.var_server.get().strip()
        user     = self.var_user.get().strip()
        password = self.var_password.get()

        # Определяем количество дней
        custom = self.var_custom_days.get().strip()
        if custom:
            if not custom.isdigit() or int(custom) <= 0:
                messagebox.showerror("Ошибка", "Введите корректное количество дней (целое число > 0)")
                return
            days = int(custom)
        else:
            days = self.var_days.get()

        # Валидация
        if not server:
            messagebox.showerror("Ошибка", "Укажите имя сервера Domino")
            return
        if not user:
            messagebox.showerror("Ошибка", "Укажите имя пользователя")
            return

        self._set_busy(True)
        self._log(f"Запуск продления сертификата...", "warn")
        self._log(f"Сервер    : {server}", "info")
        self._log(f"Пользователь: {user}", "info")
        self._log(f"Срок      : {days} дней", "info")

        # Запускаем в отдельном потоке, чтобы GUI не зависал
        thread = threading.Thread(
            target=self._run_worker,
            args=(server, user, password, days),
            daemon=True
        )
        thread.start()

    def _run_worker(self, server, user, password, days):
        """Выполняется в фоновом потоке."""
        result = call_worker(server, user, password, days)
        # Обновляем GUI из главного потока
        self.after(0, self._on_result, result)

    def _on_result(self, result):
        """Вызывается в главном потоке после завершения воркера."""
        self._set_busy(False)

        if result["success"]:
            data = result.get("data", {})
            self._log("─" * 45, "time")
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
            self._log("─" * 45, "time")

            messagebox.showinfo(
                "Готово",
                f"Запрос на продление создан!\n\n"
                f"Пользователь: {data.get('full_name', '')}\n"
                f"Новая дата истечения: {data.get('new_expiration', 'н/д')}\n\n"
                f"AdminP обработает запрос при следующем запуске."
            )
        else:
            self._log("─" * 45, "time")
            self._log("ОШИБКА: " + result["message"], "error")
            self._log("─" * 45, "time")

            messagebox.showerror("Ошибка", result["message"])


# ─── Запуск ───────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = LotusRenewApp()
    app.mainloop()
