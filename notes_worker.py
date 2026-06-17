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


def parse_cert_date_from_bytes(raw_bytes):
    """
    Парсит дату истечения из бинарных данных сертификата Lotus Notes.
    Notes хранит дату как 6 байт в формате: YY YY MM DD HH MM
    (первые 2 байта = год в big-endian, затем месяц, день, часы, минуты)
    """
    try:
        if len(raw_bytes) < 6:
            return None

        # Ищем валидную дату перебирая смещения в бинарных данных
        # Notes дата: 2 байта год (big-endian) + 1 байт месяц + 1 байт день
        for offset in range(0, len(raw_bytes) - 5):
            try:
                year  = int.from_bytes(raw_bytes[offset:offset+2], 'big')
                month = raw_bytes[offset+2]
                day   = raw_bytes[offset+3]

                # Проверяем что это похоже на дату (год 2020-2040, месяц 1-12, день 1-31)
                if 2020 <= year <= 2040 and 1 <= month <= 12 and 1 <= day <= 31:
                    return datetime(year, month, day)
            except Exception:
                continue
    except Exception:
        pass
    return None


def get_cert_expiration_from_doc(doc):
    """
    Читает дату истечения сертификата из документа пользователя.
    """
    # ── Метод 1: стандартные текстовые поля ──────────────────────────────────
    TEXT_FIELDS = [
        "CertExpiration", "CertificateExpiration", "Expiration",
        "CertExp", "HTTPPasswordExpires", "PasswordExpiration",
        "CIntDate", "CertDate", "CertValidUntil", "ValidUntil",
    ]
    for field in TEXT_FIELDS:
        try:
            val = doc.GetItemValue(field)
            if val and val[0] and str(val[0]).strip():
                v = str(val[0]).strip()
                if any(c.isdigit() for c in v):
                    return v, field
        except Exception:
            continue

    # ── Метод 2: читаем все DATETIME поля (тип 7) ─────────────────────────────
    try:
        items = doc.Items
        for item in items:
            try:
                if item.Type == 7:  # DATETIME
                    val = doc.GetItemValue(item.Name)
                    if val and val[0]:
                        return str(val[0]), item.Name
            except Exception:
                continue
    except Exception:
        pass

    # ── Метод 3: парсим бинарные данные поля Certificate ─────────────────────
    for field in ["Certificate", "Certificates", "UserCertificate"]:
        try:
            item = doc.GetFirstItem(field)
            if item is None:
                continue
            val = doc.GetItemValue(field)
            if val and val[0]:
                raw_str = str(val[0]).replace(" ", "").replace("\n", "")
                try:
                    raw_bytes = bytes.fromhex(raw_str)
                    date = parse_cert_date_from_bytes(raw_bytes)
                    if date:
                        return date.strftime("%m/%d/%Y"), f"{field}[binary]"
                except Exception:
                    pass
        except Exception:
            continue

    return None, None
    """Вычисляет сколько дней осталось до даты. Поддерживает все форматы Notes."""
    try:
        date_str = str(date_val).strip()

        # ── Пробуем строковые форматы ─────────────────────────────────────────
        string_formats = [
            "%m/%d/%Y",      # 12/31/2027
            "%d.%m.%Y",      # 31.12.2027
            "%Y-%m-%d",      # 2027-12-31
            "%d/%m/%Y",      # 31/12/2027
            "%Y%m%d",        # 20271231
            "%m/%d/%Y %H:%M:%S",   # 12/31/2027 00:00:00
            "%d.%m.%Y %H:%M:%S",   # 31.12.2027 00:00:00
            "%Y-%m-%dT%H:%M:%S",   # 2027-12-31T00:00:00
        ]
        for fmt in string_formats:
            try:
                exp_date = datetime.strptime(date_str.split()[0] if " " in date_str else date_str, fmt)
                return (exp_date - datetime.now()).days
            except Exception:
                continue

        # ── Пробуем числовой формат ───────────────────────────────────────────
        # Убираем всё нечисловое кроме точки
        clean = ''.join(c for c in date_str if c.isdigit() or c == '.')
        if clean:
            num = float(clean)

            # Notes хранит даты как количество секунд с 01.01.1900
            # Обычно это большое число > 3_000_000_000 (это уже после 1995 года в Notes)
            if num > 3_000_000_000:
                # Notes datetime: секунды с 01.01.1900 00:00:00
                notes_epoch = datetime(1899, 12, 30)
                exp_date = notes_epoch + __import__('datetime').timedelta(seconds=num)
                return (exp_date - datetime.now()).days

            elif num > 1_000_000_000:
                # Unix timestamp: секунды с 01.01.1970
                exp_date = datetime.fromtimestamp(num)
                return (exp_date - datetime.now()).days

            elif num > 40_000:
                # OLE/Excel дата: дни с 30.12.1899
                notes_epoch = datetime(1899, 12, 30)
                exp_date = notes_epoch + __import__('datetime').timedelta(days=int(num))
                return (exp_date - datetime.now()).days

    except Exception:
        pass
    return None


def format_date(date_val):
    """Форматирует дату Notes в читаемый вид дд.мм.гггг."""
    try:
        date_str = str(date_val).strip()

        string_formats = [
            "%m/%d/%Y", "%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y",
            "%m/%d/%Y %H:%M:%S", "%d.%m.%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S",
        ]
        for fmt in string_formats:
            try:
                d = datetime.strptime(date_str.split()[0] if " " in date_str else date_str, fmt)
                return d.strftime("%d.%m.%Y")
            except Exception:
                continue

        clean = ''.join(c for c in date_str if c.isdigit() or c == '.')
        if clean:
            num = float(clean)
            notes_epoch = datetime(1899, 12, 30)
            if num > 3_000_000_000:
                d = notes_epoch + __import__('datetime').timedelta(seconds=num)
                return d.strftime("%d.%m.%Y")
            elif num > 1_000_000_000:
                d = datetime.fromtimestamp(num)
                return d.strftime("%d.%m.%Y")
            elif num > 40_000:
                d = notes_epoch + __import__('datetime').timedelta(days=int(num))
                return d.strftime("%d.%m.%Y")
    except Exception:
        pass
    return str(date_val)


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

    # ── Ищем дату истечения через новый универсальный метод ──────────────────
    exp_str, exp_field_found = get_cert_expiration_from_doc(doc)

    # ── Если не нашли — сканируем ВСЕ поля документа для диагностики ─────────
    diag_fields = {}
    if not exp_str:
        try:
            items = doc.Items
            for item in items:
                try:
                    val = doc.GetItemValue(item.Name)
                    if val and val[0] and str(val[0]).strip():
                        diag_fields[item.Name] = f"[тип={item.Type}] {str(val[0])[:60]}"
                except Exception:
                    pass
        except Exception:
            pass
        data["debug_fields"] = diag_fields

    if exp_str:
        data["expiration_date"] = format_date(exp_str)  # красивый формат дд.мм.гггг
        data["expiration_field"] = exp_field_found
        days = days_until(exp_str)
        if days is not None:
            data["days_left"] = days
            if days < 0:
                data["status"] = "expired"
                data["status_text"] = f"ПРОСРОЧЕН ({abs(days)} дней назад)"
            elif days <= 30:
                data["status"] = "warning"
                data["status_text"] = f"Истекает через {days} дней — срочно!"
            elif days <= 90:
                data["status"] = "warning"
                data["status_text"] = f"Истекает через {days} дней"
            else:
                data["status"] = "ok"
                data["status_text"] = f"Действителен ещё {days} дней"
        else:
            data["status"] = "ok"
            data["status_text"] = f"Действителен (дата: {format_date(exp_str)})"
    else:
        data["expiration_date"] = "Не найдена"
        data["status"] = "unknown"
        data["status_text"] = (
            "Дата истечения не найдена в документе.\n"
            f"Найденные поля: {list(diag_fields.keys()) if diag_fields else 'нет'}"
        )

    # Ищем дату выдачи
    ISSUED_FIELDS = ["CertIssued", "CertificateIssued", "Issued", "CertDate"]
    for field in ISSUED_FIELDS:
        try:
            val = doc.GetItemValue(field)
            if val and val[0] and str(val[0]).strip():
                data["issued_date"] = format_date(str(val[0]))
                break
        except Exception:
            continue

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
