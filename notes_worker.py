# notes_worker.py — запускается ТОЛЬКО через 32-битный Python
# C:\Python313-32\python.exe notes_worker.py

import sys
import json
import traceback
from datetime import datetime, timedelta

# Принудительно utf-8
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")
if hasattr(sys.stdin, "reconfigure"):
    sys.stdin.reconfigure(encoding="utf-8")


def send(obj):
    sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
    sys.stdout.flush()
    sys.exit(0)


# ─── Парсинг дат ──────────────────────────────────────────────────────────────

def format_date(date_str):
    """Форматирует строку даты в дд.мм.гггг."""
    s = str(date_str).strip()
    fmts = [
        "%Y-%m-%d %H:%M:%S%z", "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",   "%m/%d/%Y",
        "%d.%m.%Y %H:%M:%S",   "%d.%m.%Y",
        "%Y-%m-%dT%H:%M:%S",   "%Y-%m-%d",
    ]
    for fmt in fmts:
        try:
            d = datetime.strptime(s[:19], fmt[:len(s[:19])])
            return d.strftime("%d.%m.%Y")
        except Exception:
            continue
    # Если не распарсили — вернуть как есть
    return s[:10]


def days_until(date_str):
    """Вычисляет сколько дней до даты."""
    s = str(date_str).strip()
    fmts = [
        "%Y-%m-%d %H:%M:%S%z", "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",   "%m/%d/%Y",
        "%d.%m.%Y %H:%M:%S",   "%d.%m.%Y",
        "%Y-%m-%dT%H:%M:%S",   "%Y-%m-%d",
    ]
    for fmt in fmts:
        try:
            d = datetime.strptime(s[:19], fmt[:len(s[:19])])
            return (d - datetime.now()).days
        except Exception:
            continue
    return None


def parse_cert_bytes(raw_bytes):
    """
    Извлекает дату из бинарных данных сертификата Notes.
    Формат Notes: ... 2 байта год (big-endian) + 1 байт месяц + 1 байт день ...
    """
    for offset in range(0, max(0, len(raw_bytes) - 4)):
        try:
            year  = int.from_bytes(raw_bytes[offset:offset+2], 'big')
            month = raw_bytes[offset + 2]
            day   = raw_bytes[offset + 3]
            if 2000 <= year <= 2050 and 1 <= month <= 12 and 1 <= day <= 31:
                return datetime(year, month, day)
        except Exception:
            continue
    return None


# ─── Сессия и поиск пользователя ─────────────────────────────────────────────

def get_session_and_db(win32com, server_name, notes_password):
    try:
        session = win32com.client.Dispatch("Lotus.NotesSession")
    except Exception as e:
        send({"success": False, "message": f"Не удалось запустить Lotus Notes COM:\n{e}", "data": {}})

    try:
        session.Initialize(notes_password)
    except Exception as e:
        send({"success": False, "message": f"Ошибка инициализации (неверный пароль?):\n{e}", "data": {}})

    try:
        db = session.GetDatabase(server_name, "names.nsf")
        if not db.IsOpen:
            db.Open()
    except Exception as e:
        send({"success": False, "message": f"Не удалось открыть names.nsf:\n{e}", "data": {}})

    if not db.IsOpen:
        send({"success": False, "message": f"names.nsf закрыта — нет доступа к серверу '{server_name}'", "data": {}})

    return session, db


def find_user_doc(db, user_name):
    view = db.GetView("($Users)")
    if view is None:
        send({"success": False, "message": "Представление ($Users) не найдено", "data": {}})
    doc = view.GetDocumentByKey(user_name, True)
    if doc is None:
        send({"success": False, "message": f"Пользователь '{user_name}' не найден", "data": {}})
    return doc


# ─── Получение даты сертификата ───────────────────────────────────────────────

def get_expiration(doc, id_file_path=""):
    """
    Дата истечения хранится в ID файле пользователя (TECH10.ID),
    а не в документе names.nsf.
    Читаем путь к ID файлу из документа и парсим его бинарные данные.
    """

    # ── Метод 1: ищем ID файл по имени пользователя автоматически ───────────
    import os

    # Извлекаем короткое имя из CN=tech10/O=Guli → tech10
    try:
        short_name = doc.GetItemValue("ShortName")
        short_name = str(short_name[0]).strip() if short_name and short_name[0] else ""
    except Exception:
        short_name = ""

    if not short_name:
        # Парсим из FullName: CN=tech10/O=Guli → tech10
        try:
            full = doc.GetItemValue("FullName")[0]
            parts = str(full).replace("CN=", "").split("/")
            short_name = parts[0].strip()
        except Exception:
            short_name = ""

    # Стандартные папки Notes Data
    data_dirs = [
        r"C:\Program Files (x86)\IBM\Notes\Data",
        r"C:\Program Files\IBM\Notes\Data",
        r"C:\Program Files (x86)\HCL\Notes\Data",
        r"C:\Program Files\HCL\Notes\Data",
        r"C:\Lotus\Notes\Data",
        r"C:\IBM\Notes\Data",
        os.path.expandvars(r"%APPDATA%\IBM\Notes\Data"),
        os.path.expandvars(r"%LOCALAPPDATA%\IBM\Notes\Data"),
        os.path.expandvars(r"%APPDATA%\HCL\Notes\Data"),
    ]

    auto_found = None
    if short_name:
        for data_dir in data_dirs:
            for name_variant in [short_name, short_name.upper(), short_name.lower()]:
                candidate = os.path.join(data_dir, f"{name_variant}.ID")
                if os.path.exists(candidate):
                    auto_found = candidate
                    break
            if auto_found:
                break

    # Читаем ID файл (приоритет: вручную указанный → автонайденный)
    id_path = id_file_path or auto_found
    if id_path and os.path.exists(id_path):
        d = parse_id_file(id_path)
        if d:
            return d.strftime("%Y-%m-%d"), f"IDFile:{id_path}"

    # ── Метод 2: ClntDate — только как справочная инфо, НЕ как дата истечения──
    # (не возвращаем её как дату истечения — это дата последнего входа!)

    return None, None


def parse_id_file(id_file_path):
    """
    Парсит ID файл Lotus Notes и извлекает дату истечения сертификата.
    ID файл — бинарный формат, дата хранится в структуре сертификата.
    Ищем год в диапазоне 2020-2060.
    """
    try:
        with open(id_file_path, "rb") as f:
            data = f.read()

        # Notes хранит дату как 2 байта год (big-endian) + 1 байт месяц + 1 байт день
        # Ищем самую позднюю валидную дату (это и есть дата истечения)
        found_dates = []
        for offset in range(0, len(data) - 4):
            try:
                year  = int.from_bytes(data[offset:offset+2], 'big')
                month = data[offset + 2]
                day   = data[offset + 3]
                if 2020 <= year <= 2060 and 1 <= month <= 12 and 1 <= day <= 31:
                    try:
                        d = datetime(year, month, day)
                        found_dates.append(d)
                    except Exception:
                        continue
            except Exception:
                continue

        if found_dates:
            # Берём самую позднюю дату — это дата истечения сертификата
            return max(found_dates)

    except Exception:
        pass
    return None


# ─── Действия ─────────────────────────────────────────────────────────────────

def action_check(win32com, params):
    server_name    = params.get("server_name", "")
    user_name      = params.get("user_name", "")
    notes_password = params.get("notes_password", "")
    id_file_path   = params.get("id_file_path", "")

    session, db = get_session_and_db(win32com, server_name, notes_password)
    doc = find_user_doc(db, user_name)

    data = {
        "admin_user": session.CommonUserName,
        "full_name":  doc.GetItemValue("FullName")[0],
    }

    exp_str, exp_field = get_expiration(doc, id_file_path)

    if exp_str:
        data["expiration_date"]  = format_date(exp_str)
        data["expiration_field"] = exp_field
        # Показываем какой ID файл был использован
        if "IDFile:" in (exp_field or ""):
            data["id_file_used"] = exp_field.replace("IDFile:", "")
        days = days_until(exp_str)
        if days is not None:
            data["days_left"] = days
            if days < 0:
                data["status"]      = "expired"
                data["status_text"] = f"ПРОСРОЧЕН ({abs(days)} дней назад)"
            elif days <= 30:
                data["status"]      = "warning"
                data["status_text"] = f"Истекает через {days} дней — срочно!"
            elif days <= 90:
                data["status"]      = "warning"
                data["status_text"] = f"Истекает через {days} дней"
            else:
                data["status"]      = "ok"
                data["status_text"] = f"Действителен ещё {days} дней"
        else:
            data["status"]      = "ok"
            data["status_text"] = f"Действителен"
    else:
        # Дата не найдена — показываем подсказку
        data["expiration_date"] = "Не найдена"
        data["status"]          = "unknown"

        # Показываем ClntDate как справочную информацию
        info_lines = []
        for field, label in [("ClntDate", "Последний вход"), ("$Revisions", "Изменён"), ("HTTPPasswordChangeDate", "Смена пароля")]:
            try:
                val = doc.GetItemValue(field)
                if val and val[0] and str(val[0]).strip():
                    info_lines.append(f"{label}: {format_date(str(val[0]))}")
            except Exception:
                pass

        data["status_text"] = (
            "Укажите путь к ID файлу пользователя\n"
            "в поле 'ID файл (.ID)' и нажмите '...'\n"
            f"Обычно: C:\\Program Files (x86)\\IBM\\Notes\\Data\\{user_name.split('/')[0].split('=')[-1]}.ID"
        )
        # Сообщаем какой файл искали автоматически
        if auto_found is None and short_name:
            data["status_text"] += f"\n\nФайл {short_name}.ID не найден автоматически"

    # Дата последнего входа клиента
    try:
        val = doc.GetItemValue("ClntDate")
        if val and val[0]:
            data["last_login"] = format_date(str(val[0]))
    except Exception:
        pass

    send({"success": True, "message": "Данные получены", "data": data})


def action_renew(win32com, params):
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

    # Текущая дата
    exp_str, _ = get_expiration(doc)
    if exp_str:
        result_data["current_expiration"] = format_date(exp_str)

    # Открываем Admin4.nsf
    try:
        admin_db = session.GetDatabase(server_name, "admin4.nsf")
        if not admin_db.IsOpen:
            admin_db.Open()
    except Exception as e:
        send({"success": False, "message": f"Не удалось открыть admin4.nsf:\n{e}", "data": {}})

    if not admin_db.IsOpen:
        send({"success": False, "message": "admin4.nsf закрыта — нет прав администратора?", "data": {}})

    req = admin_db.CreateDocument()
    req.ReplaceItemValue("Form",                "AdminRequest")
    req.ReplaceItemValue("ProxyAction",         "78")
    req.ReplaceItemValue("ProxyNameList",       [user_name])
    req.ReplaceItemValue("ProxyServer",         server_name)
    req.ReplaceItemValue("ProxyExpirationDate", expiration_days)
    req.ReplaceItemValue("ProxyCreated",        datetime.now().strftime("%m/%d/%Y"))
    req.Save(True, False)

    new_exp = datetime.now() + timedelta(days=expiration_days)
    result_data["new_expiration"] = new_exp.strftime("%d.%m.%Y")

    send({"success": True, "message": "Запрос на продление успешно создан в admin4.nsf", "data": result_data})


# ─── Точка входа ──────────────────────────────────────────────────────────────

def main():
    try:
        raw = sys.stdin.read()
    except Exception as e:
        send({"success": False, "message": f"Ошибка чтения stdin: {e}", "data": {}})

    if not raw or not raw.strip():
        send({"success": False, "message": "Пустой stdin", "data": {}})

    try:
        params = json.loads(raw)
    except Exception as e:
        send({"success": False, "message": f"Ошибка разбора JSON: {e}", "data": {}})

    try:
        import win32com.client
        import win32com
    except ImportError as e:
        send({"success": False, "message": f"win32com не найден: {e}\nУстановите: C:\\Python313-32\\python.exe -m pip install pywin32", "data": {}})

    action = params.get("action", "renew")
    try:
        if action == "check":
            action_check(win32com, params)
        else:
            action_renew(win32com, params)
    except Exception as e:
        send({"success": False, "message": f"{e}\n{traceback.format_exc()}", "data": {}})


if __name__ == "__main__":
    main()
