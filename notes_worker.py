# notes_worker.py — запускается ТОЛЬКО через 32-битный Python
# C:\Python313-32\python.exe notes_worker.py

import sys
import json
import os
import traceback
from datetime import datetime, timedelta


def send(obj):
    """Гарантированно выводим JSON в stdout и завершаем процесс."""
    try:
        sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
        sys.stdout.flush()
    except Exception:
        # Крайний случай — пишем хоть что-то
        sys.stdout.write('{"success":false,"message":"Критическая ошибка вывода","data":{}}\n')
        sys.stdout.flush()
    sys.exit(0)


def main():
    # ── 1. Читаем параметры из stdin ─────────────────────────────────────────
    try:
        raw = sys.stdin.read()
    except Exception as e:
        send({"success": False, "message": f"Не удалось прочитать stdin: {e}", "data": {}})

    if not raw or not raw.strip():
        send({"success": False, "message": "Воркер не получил параметры (пустой stdin)", "data": {}})

    try:
        params = json.loads(raw)
    except Exception as e:
        send({"success": False, "message": f"Ошибка разбора JSON параметров: {e}\nПолучено: {raw[:200]}", "data": {}})

    server_name     = params.get("server_name", "")
    user_name       = params.get("user_name", "")
    notes_password  = params.get("notes_password", "")
    expiration_days = params.get("expiration_days", 365)

    # ── 2. Импортируем win32com ───────────────────────────────────────────────
    try:
        import win32com.client
    except ImportError as e:
        send({
            "success": False,
            "message": (
                f"Модуль win32com не найден: {e}\n"
                f"Выполните: C:\\Python313-32\\python.exe -m pip install pywin32\n"
                f"Python: {sys.executable}\n"
                f"Версия: {sys.version}"
            ),
            "data": {}
        })

    result = {"success": False, "message": "", "data": {}}

    # ── 3. Основная логика ────────────────────────────────────────────────────
    try:
        # Открываем сессию Notes
        try:
            session = win32com.client.Dispatch("Lotus.NotesSession")
        except Exception as e:
            send({"success": False, "message": f"Не удалось запустить Lotus Notes COM:\n{e}\nУбедитесь что Lotus Notes установлен и зарегистрирован.", "data": {}})

        try:
            session.Initialize(notes_password)
        except Exception as e:
            send({"success": False, "message": f"Ошибка инициализации сессии Notes (неверный пароль?):\n{e}", "data": {}})

        result["data"]["admin_user"] = session.CommonUserName

        # Открываем Domino Directory
        try:
            db = session.GetDatabase(server_name, "names.nsf")
            if not db.IsOpen:
                db.Open()
        except Exception as e:
            send({"success": False, "message": f"Не удалось открыть names.nsf на сервере '{server_name}':\n{e}", "data": {}})

        if not db.IsOpen:
            send({"success": False, "message": f"names.nsf закрыта — нет доступа к серверу '{server_name}'", "data": {}})

        # Ищем пользователя
        view = db.GetView("($Users)")
        if view is None:
            send({"success": False, "message": "Представление ($Users) не найдено в names.nsf", "data": {}})

        doc = view.GetDocumentByKey(user_name, True)
        if doc is None:
            send({"success": False, "message": f"Пользователь '{user_name}' не найден в директории", "data": {}})

        result["data"]["full_name"] = doc.GetItemValue("FullName")[0]

        # Читаем текущую дату истечения
        try:
            cert_exp = doc.GetItemValue("CertExpiration")
            if cert_exp and cert_exp[0]:
                result["data"]["current_expiration"] = str(cert_exp[0])
        except Exception:
            pass

        # Открываем Admin4.nsf
        try:
            admin_db = session.GetDatabase(server_name, "admin4.nsf")
            if not admin_db.IsOpen:
                admin_db.Open()
        except Exception as e:
            send({"success": False, "message": f"Не удалось открыть admin4.nsf:\n{e}", "data": {}})

        if not admin_db.IsOpen:
            send({"success": False, "message": f"admin4.nsf закрыта — нет прав администратора?", "data": {}})

        # Создаём запрос AdminP
        req = admin_db.CreateDocument()
        req.ReplaceItemValue("Form",                "AdminRequest")
        req.ReplaceItemValue("ProxyAction",         "78")  # 78 = Recertify User
        req.ReplaceItemValue("ProxyNameList",       [user_name])
        req.ReplaceItemValue("ProxyServer",         server_name)
        req.ReplaceItemValue("ProxyExpirationDate", expiration_days)
        req.ReplaceItemValue("ProxyCreated",        datetime.now().strftime("%m/%d/%Y"))
        req.Save(True, False)

        new_exp = datetime.now() + timedelta(days=expiration_days)
        result["success"] = True
        result["message"] = "Запрос на продление успешно создан в admin4.nsf"
        result["data"]["new_expiration"] = new_exp.strftime("%d.%m.%Y")

    except Exception as e:
        result["success"] = False
        result["message"] = f"{e}\n\n{traceback.format_exc()}"

    send(result)


if __name__ == "__main__":
    main()
