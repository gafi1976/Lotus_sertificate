import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import threading
import sys
from io import StringIO
import win32com.client

class RecertifyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ресертификация пользователей Lotus Notes")
        self.root.geometry("650x600")
        self.root.resizable(False, False)

        self.server_var = tk.StringVar(value="YourServer/YourOrg")
        self.user_var = tk.StringVar(value="User Full Name/YourOrg")
        self.session_pwd_var = tk.StringVar()
        self.cert_file_var = tk.StringVar()
        self.cert_pwd_var = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")

        row = 0
        ttk.Label(main_frame, text="Сервер Domino:").grid(row=row, column=0, sticky="w", pady=5)
        ttk.Entry(main_frame, textvariable=self.server_var, width=50).grid(row=row, column=1, columnspan=2, sticky="w", padx=5)
        row += 1

        ttk.Label(main_frame, text="Пользователь (полное имя):").grid(row=row, column=0, sticky="w", pady=5)
        ttk.Entry(main_frame, textvariable=self.user_var, width=50).grid(row=row, column=1, columnspan=2, sticky="w", padx=5)
        row += 1

        ttk.Label(main_frame, text="Пароль вашего ID-файла:").grid(row=row, column=0, sticky="w", pady=5)
        ttk.Entry(main_frame, textvariable=self.session_pwd_var, width=50, show="*").grid(row=row, column=1, columnspan=2, sticky="w", padx=5)
        row += 1

        ttk.Label(main_frame, text="Путь к cert.id:").grid(row=row, column=0, sticky="w", pady=5)
        cert_frame = ttk.Frame(main_frame)
        cert_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)
        ttk.Entry(cert_frame, textvariable=self.cert_file_var, width=40).pack(side=tk.LEFT)
        ttk.Button(cert_frame, text="Обзор...", command=self.browse_cert).pack(side=tk.LEFT, padx=5)
        row += 1

        ttk.Label(main_frame, text="Пароль от cert.id:").grid(row=row, column=0, sticky="w", pady=5)
        ttk.Entry(main_frame, textvariable=self.cert_pwd_var, width=50, show="*").grid(row=row, column=1, columnspan=2, sticky="w", padx=5)
        row += 1

        self.run_btn = ttk.Button(main_frame, text="Запустить ресертификацию", command=self.start_recertify)
        self.run_btn.grid(row=row, column=0, columnspan=3, pady=10)

        row += 1
        ttk.Label(main_frame, text="Лог выполнения:").grid(row=row, column=0, columnspan=3, sticky="w", pady=(10,0))
        row += 1
        self.log_text = scrolledtext.ScrolledText(main_frame, height=18, state='disabled', wrap=tk.WORD)
        self.log_text.grid(row=row, column=0, columnspan=3, sticky="nsew", pady=5)

        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(row, weight=1)

        self.redirect_output()

    def browse_cert(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл cert.id",
            filetypes=[("ID файлы", "*.id"), ("Все файлы", "*.*")]
        )
        if filename:
            self.cert_file_var.set(filename)

    def redirect_output(self):
        self.stdout_redirect = StringIO()
        sys.stdout = self.stdout_redirect
        sys.stderr = self.stdout_redirect
        self.update_log()

    def update_log(self):
        if self.stdout_redirect:
            content = self.stdout_redirect.getvalue()
            if content:
                self.log_text.config(state='normal')
                self.log_text.insert(tk.END, content)
                self.log_text.see(tk.END)
                self.log_text.config(state='disabled')
                self.stdout_redirect.seek(0)
                self.stdout_redirect.truncate(0)
        self.root.after(200, self.update_log)

    def start_recertify(self):
        self.run_btn.config(state='disabled')
        self.log_text.config(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state='disabled')
        thread = threading.Thread(target=self.do_recertify, daemon=True)
        thread.start()

    def do_recertify(self):
        try:
            server = self.server_var.get().strip()
            user = self.user_var.get().strip()
            session_pwd = self.session_pwd_var.get().strip()
            cert_file = self.cert_file_var.get().strip()
            cert_pwd = self.cert_pwd_var.get().strip()

            if not all([server, user, session_pwd, cert_file, cert_pwd]):
                print("ОШИБКА: Все поля должны быть заполнены.")
                return

            print(f"Сервер: {server}")
            print(f"Пользователь: {user}")
            print(f"Файл сертификатора: {cert_file}")

            # Используем win32com напрямую
            self.recertify_with_win32com(server, user, session_pwd, cert_file, cert_pwd)

        except Exception as e:
            print(f"Критическая ошибка: {e}")
            import traceback
            traceback.print_exc()
        finally:
            self.root.after(0, lambda: self.run_btn.config(state='normal'))

    def recertify_with_win32com(self, server, user, session_pwd, cert_file, cert_pwd):
        try:
            # Создаём COM-объект
            session = win32com.client.Dispatch("Lotus.NotesSession")
            # Инициализация с паролем (может потребоваться дополнительный параметр - путь к ID-файлу)
            session.Initialize(session_pwd)
            print("Сессия Notes создана через win32com.")

            admin_proc = session.CreateAdministrationProcess(server)
            admin_proc.CertifierFile = cert_file
            admin_proc.CertifierPassword = cert_pwd

            print("Отправляем запрос на ресертификацию...")
            note_id = admin_proc.RecertifyUser(user)
            if note_id:
                print(f"Запрос успешно отправлен. ID записи: {note_id}")
            else:
                print("Запрос отправлен (возможно, обработан немедленно).")
            print("Готово!")

        except Exception as e:
            print(f"Ошибка при работе с win32com: {e}")
            # Попробуем расшифровать код ошибки
            if "The server threw an exception" in str(e):
                print("Возможно, неверный пароль или недостаточно прав.")
            raise

if __name__ == "__main__":
    root = tk.Tk()
    app = RecertifyApp(root)
    root.mainloop()