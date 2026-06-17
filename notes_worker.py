# notes_worker.py — запускается ТОЛЬКО через 32-битный Python
# C:\Python313-32\python.exe notes_worker.py

import sys
import json
import win32com.client
from datetime import datetime, timedelta


def renew_certificate(server_name, user_name, notes_password, expiration_days):
    result = {"success": False, "message": "", "data": {}}

    try:
        # Открываем сессию Notes
        session = win32com.client.Dispatch("Lotus.NotesSession")
        session.Initialize(notes_password)
        result["data"]["admin_user"] = session.CommonUserName

        # Открываем Domino Directory
        db = session.GetDatabase(server_name, "names.nsf")
        if not db.IsOpen:
            db.Open()

        if not db.IsOpen:
            result["message"] = f"Не удалось открыть names.nsf на сервере '{server_name}'"
            return result

        # Ищем пользователя
        view = db.GetView("($Users)")
        if view is None:
            result["message"] = "Представление ($Users) не найдено в names.nsf"
            return result

        doc = view.GetDocumentByKey(user_name, True)

        if doc is None:
            result["message"] = f"Пользователь '{user_name}' не найден в директории"
            return result

        result["data"]["full_name"] = doc.GetItemValue("FullName")[0]

        # Читаем текущую дату истечения сертификата
        try:
            cert_exp = doc.GetItemValue("CertExpiration")
            if cert_exp and cert_exp[0]:
                result["data"]["current_expiration"] = str(cert_exp[0])
        except Exception:
            pass

        # Открываем Admin4.nsf и создаём запрос AdminP
        admin_db = session.GetDatabase(server_name, "admin4.nsf")
        if not admin_db.IsOpen:
            admin_db.Open()

        if not admin_db.IsOpen:
            result["message"] = f"Не удалось открыть admin4.nsf на сервере '{server_name}'"
            return result

        req = admin_db.CreateDocument()
        req.ReplaceItemValue("Form",                 "AdminRequest")
        req.ReplaceItemValue("ProxyAction",          "78")   # 78 = Recertify User
        req.ReplaceItemValue("ProxyNameList",        [user_name])
        req.ReplaceItemValue("ProxyServer",          server_name)
        req.ReplaceItemValue("ProxyExpirationDate",  expiration_days)
        req.ReplaceItemValue("ProxyCreated",         datetime.now().strftime("%m/%d/%Y"))
        req.Save(True, False)

        new_exp = datetime.now() + timedelta(days=expiration_days)
        result["success"] = True
        result["message"] = "Запрос на продление успешно создан в admin4.nsf"
        result["data"]["new_expiration"] = new_exp.strftime("%d.%m.%Y")

    except Exception as e:
        result["message"] = str(e)

    return result


if __name__ == "__main__":
    try:
        params = json.loads(sys.stdin.read())
        output = renew_certificate(
            server_name     = params["server_name"],
            user_name       = params["user_name"],
            notes_password  = params.get("notes_password", ""),
            expiration_days = params.get("expiration_days", 365),
        )
    except Exception as e:
        output = {"success": False, "message": f"Ошибка воркера: {e}", "data": {}}

    print(json.dumps(output, ensure_ascii=False))
