# read_nsf.py — чтение содержимого NSF файла
# C:\Python313-32\python.exe read_nsf.py

import win32com.client
import json

# ─── НАСТРОЙКИ ────────────────────────────────────────────────────────────────
SERVER_NAME    = ""         # "" = локальный, или "Guli_mail/Guli"
NSF_FILE       = "admin4.nsf"   # имя файла или полный путь
NOTES_PASSWORD = ""
MAX_DOCS       = 50         # сколько документов показать
# ──────────────────────────────────────────────────────────────────────────────

session = win32com.client.Dispatch("Lotus.NotesSession")
session.Initialize(NOTES_PASSWORD)
print(f"Сессия: {session.CommonUserName}\n")

db = session.GetDatabase(SERVER_NAME, NSF_FILE)
if not db.IsOpen:
    db.Open()

print(f"База: {db.FilePath}")
print(f"Заголовок: {db.Title}")
print(f"Документов: {db.AllDocuments.Count}\n")

# Показываем все представления
print("=== ПРЕДСТАВЛЕНИЯ (Views) ===")
views = db.Views
for v in views:
    try:
        print(f"  {v.Name}")
    except Exception:
        pass

print("\n=== ДОКУМЕНТЫ ===")
view = db.GetView("($All)")
if view is None:
    view = db.AllDocuments

doc = view.GetFirstDocument() if hasattr(view, 'GetFirstDocument') else None
count = 0

while doc and count < MAX_DOCS:
    print(f"\n--- Документ {count+1} ---")
    try:
        items = doc.Items
        for item in items:
            try:
                val = doc.GetItemValue(item.Name)
                if val and val[0]:
                    val_str = str(val[0])[:80]
                    print(f"  {item.Name:<30} = {val_str}")
            except Exception:
                pass
    except Exception as e:
        print(f"  Ошибка: {e}")
    
    count += 1
    try:
        doc = view.GetNextDocument(doc)
    except Exception:
        break

print(f"\nПоказано {count} документов")
input("Нажмите Enter для выхода...")
