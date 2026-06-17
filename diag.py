# diag.py — диагностика полей документа пользователя
# Запускать через 32-битный Python:
# C:\Python313-32\python.exe diag.py

import sys
import win32com.client

# ─── НАСТРОЙТЕ ЭТИ ПАРАМЕТРЫ ─────────────────────────────────────────────────
SERVER_NAME    = "Guli_mail/Guli"   # ← имя вашего сервера
USER_NAME      = "tech10/Guli"      # ← имя пользователя
NOTES_PASSWORD = ""                  # ← пароль если есть
# ──────────────────────────────────────────────────────────────────────────────

print("=" * 60)
print("ДИАГНОСТИКА ПОЛЕЙ ПОЛЬЗОВАТЕЛЯ LOTUS NOTES")
print("=" * 60)

try:
    session = win32com.client.Dispatch("Lotus.NotesSession")
    session.Initialize(NOTES_PASSWORD)
    print(f"Сессия открыта: {session.CommonUserName}\n")

    db = session.GetDatabase(SERVER_NAME, "names.nsf")
    if not db.IsOpen:
        db.Open()
    print(f"names.nsf открыта: {db.IsOpen}\n")

    view = db.GetView("($Users)")
    doc  = view.GetDocumentByKey(USER_NAME, True)

    if doc is None:
        print(f"ОШИБКА: Пользователь '{USER_NAME}' не найден!")
        sys.exit(1)

    print(f"Пользователь найден: {doc.GetItemValue('FullName')[0]}\n")
    print("-" * 60)
    print(f"{'ИМЯ ПОЛЯ':<35} {'ТИП':>4}  {'ЗНАЧЕНИЕ'}")
    print("-" * 60)

    # Выводим ВСЕ поля документа
    items = doc.Items
    for item in items:
        try:
            name  = item.Name
            itype = item.Type
            val   = doc.GetItemValue(name)
            if val and val[0]:
                val_str = str(val[0])[:60]
            else:
                val_str = "(пусто)"
            print(f"{name:<35} {itype:>4}  {val_str}")
        except Exception as e:
            print(f"{item.Name:<35}  ERR  {e}")

    print("-" * 60)
    print("\nТипы полей Notes:")
    print("  1 = TEXT        3 = NUMBER    6 = RICHTEXT")
    print("  7 = DATETIME    8 = NAMES    16 = USERDATA (бинарные)")

except Exception as e:
    print(f"ОШИБКА: {e}")

input("\nНажмите Enter для выхода...")
