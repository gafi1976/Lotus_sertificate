# notes_worker.py — запускается ТОЛЬКО через 32-битный Python
# C:\Python313-32\python.exe notes_worker.py

import sys
import json
import os
import traceback
from datetime import datetime, timedelta

# Принудительно utf-8 для stdout/stderr
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")
if hasattr(sys.stdin, "reconfigure"):
    sys.stdin.reconfigure(encoding="utf-8")


def send(obj):
    """Гарантированно выводим JSON в stdout и завершаем процесс."""
    try:
        sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
        sys.stdout.flush()
    except Exception:
        sys.stdout.write('{"success":false,"message":"Критическая ошибка вывода","data":{}}\n')
        sys.stdout.flush()
    sys.exit(0)


def get_session_and_db(win32com, server_name, notes_password):
    """Открывает сессию Notes и базу names.nsf. Возвращает (session, db)."""
    try:
        session = win32com.client.Dispatch("Lotus.NotesSession")
    except Exception as e:
        send({"success": False, "message": f"Не удалось запустить Lotus Notes COM:\n{e}", "data": {}})

    try:
        session.Initialize(notes_password)
    except Exception as e:
        send({"success": False, "message": f"Ошибка инициализации сессии Notes (неверный пароль?):\n{e}", "data": {}})

    try:
        db = session.GetDatabase(server_name, "names.nsf")
        if not db.IsOpen:
            db.Open()
    except Exception as e:
        send({"success": False, "message": f"Не удалось открыть names.nsf на сервере '{server_name}':\n{e}", "data": {}})

    if not db.IsOpen:
        send({"success": False, "message": f"names.nsf закрыта — нет доступа к серверу '{server_name}'", "data": {}})

    return session, db


def find_user_doc(db, user_name):
    """Ищет документ пользователя в names.nsf."""
    view = db.GetView("($Users)")
    if view is None:
        send({"success": False, "message": "Представление ($Users) не найдено в names.nsf", "data": {}})

    doc = view.GetDocumentByKey(user_name, True)
    if doc is None:
        send({"success": False, "message": f"Пользователь '{user_name}' не найден в директории", "data": {}})

    return doc


def days_until(date_str):
    """Вычисляет сколько дней осталось до даты."""
    try:
        # Notes возвращает дату в разных форматах — пробуем несколько
        for fmt in ("%m/%d/%Y", "%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                exp_date = datetime.strptime(date_str.split()[0], fmt)
                delta = (exp_date - datetime.now()).days
                return delta
            except Exception:
                continue
    except Exception:
        pass
    return None


def action_check(win32com, params):
    """Проверяет статус сертификата пользователя."""
    server_name    = params.get("server_name", "")
    user_name      = params.get("user_name", "")
    notes_password = params.get("notes_password", "")

    session, db = get_session_and_db(win32com, server_name, notes_password)
    doc = find_user_doc(db, user_name)

    data = {
        "admin_user": session.CommonUserName,
        "full_name":  doc.GetItemValue("FullName")[0],
    }

    # Дата истечения сертификата
    try:
        cert_exp = doc.GetItemValue("CertExpiration")
        if cert_exp and cert_exp[0]:
            exp_str = str(cert_exp[0])
            data["expiration_date"] = exp_str
            days = days_until(exp_str)
            if days is not None:
                data["days_left"] = days
                if days < 0:
                    data["status"] = "expired"
                    data["status_text"] = f"ПРОСРОЧЕН ({abs(days)} дней назад)"
                elif days <= 30:
                    data["status"] = "warning"
                    data["status_text"] = f"Истекает через {days} дней — скоро!"
                elif days <= 90:
                    data["status"] = "warning"
                    data["status_text"] = f"Истекает через {days} дней"
                else:
                    data["status"] = "ok"
                    data["status_text"] = f"Действителен ещё {days} дней"
            else:
                data["status"] = "ok"
                data["status_text"] = "Действителен"
        else:
            data["expiration_date"] = "Не указана"
            data["status"] = "unknown"
            data["status_text"] = "Дата не определена"
    except Exception as e:
        data["expiration_date"] = "Ошибка чтения"
        data["status"] = "unknown"
        data["status_text"] = str(e)

    # Дата выдачи сертификата
    try:
        cert_issued = doc.GetItemValue("CertIssued")
        if cert_issued and cert_issued[0]:
            data["issued_date"] = str(cert_issued[0])
    except Exception:
        pass

    send({"success": True, "message": "Данные сертификата получены", "data": data})


def action_renew(win32com, params):
    """Создаёт запрос AdminP на продление сертификата."""
    server_name     = params.get("server_name", "")
    user_name       = params.get("user_name", "")
    notes_password  = params.get("notes_password", "")
    expiration_days = params.get("expiration_days", 365)

    session, db = get_session_and_db(win32com, server_name, notes_password)
    doc = find_user_doc(db, user_name)

    result_data = {
        "admin_user": session.CommonUserName,
        "full_name":  doc.GetItemValue("FullName")[0],
    }

    # Текущая дата истечения
    try:
        cert_exp = doc.GetItemValue("CertExpiration")
        if cert_exp and cert_exp[0]:
            result_data["current_expiration"] = str(cert_exp[0])
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
        send({"success": False, "message": "admin4.nsf закрыта — нет прав администратора?", "data": {}})

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
    result_data["new_expiration"] = new_exp.strftime("%d.%m.%Y")

    send({"success": True, "message": "Запрос на продление успешно создан в admin4.nsf", "data": result_data})


def main():
    # ── Читаем параметры из stdin ─────────────────────────────────────────────
    try:
        raw = sys.stdin.read()
    except Exception as e:
        send({"success": False, "message": f"Не удалось прочитать stdin: {e}", "data": {}})

    if not raw or not raw.strip():
        send({"success": False, "message": "Воркер не получил параметры (пустой stdin)", "data": {}})

    try:
        params = json.loads(raw)
    except Exception as e:
        send({"success": False, "message": f"Ошибка разбора JSON: {e}\nПолучено: {raw[:200]}", "data": {}})

    # ── Импортируем win32com ──────────────────────────────────────────────────
    try:
        import win32com.client
        import win32com
    except ImportError as e:
        send({
            "success": False,
            "message": (
                f"Модуль win32com не найден: {e}\n"
                f"Выполните: C:\\Python313-32\\python.exe -m pip install pywin32\n"
                f"Python: {sys.executable}"
            ),
            "data": {}
        })

    # ── Выбираем действие ─────────────────────────────────────────────────────
    action = params.get("action", "renew")

    try:
        if action == "check":
            action_check(win32com, params)
        else:
            action_renew(win32com, params)
    except Exception as e:
        send({"success": False, "message": f"{e}\n\n{traceback.format_exc()}", "data": {}})


if __name__ == "__main__":
    main()
