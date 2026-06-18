# diag.py — диагностика ID файла Lotus Notes
# Запускать через 32-битный Python:
# C:\Python313-32\python.exe diag.py

import sys
import os
from datetime import datetime

# ─── НАСТРОЙТЕ ЭТИ ПАРАМЕТРЫ ─────────────────────────────────────────────────
ID_FILE_PATH = r"C:\Program Files (x86)\IBM\Notes\Data\TECH10.ID"
# ─────────────────────────────────────────────────────────────────────────────

def try_notes_com(id_file_path):
    """Пробуем получить дату через Notes COM API напрямую."""
    print("\n=== Метод 1: Notes COM API ===")
    try:
        import win32com.client
        session = win32com.client.Dispatch("Lotus.NotesSession")
        session.Initialize("")

        # Открываем ID файл через NotesDatabase
        try:
            # Метод через GetIDInfo
            id_info = session.GetEnvironmentString("$ID_FILE", True)
            print(f"Текущий ID файл из среды: {id_info}")
        except Exception as e:
            print(f"GetEnvironmentString: {e}")

        # Пробуем через NotesName
        try:
            name = session.CreateName(session.UserName)
            print(f"Имя пользователя: {name.Abbreviated}")
            print(f"Каноническое:     {name.Canonical}")
        except Exception as e:
            print(f"CreateName: {e}")

        # Пробуем получить дату через адресную книгу
        try:
            books = session.AddressBooks
            for book in books:
                print(f"Адресная книга: Server='{book.Server}' File='{book.FilePath}'")
        except Exception as e:
            print(f"AddressBooks: {e}")

    except Exception as e:
        print(f"COM ошибка: {e}")


def read_id_file_raw(id_file_path):
    """Читаем ID файл как бинарные данные и ищем дату."""
    print(f"\n=== Метод 2: Бинарный анализ файла ===")
    print(f"Файл: {id_file_path}")

    if not os.path.exists(id_file_path):
        print(f"ФАЙЛ НЕ НАЙДЕН: {id_file_path}")
        return

    with open(id_file_path, "rb") as f:
        data = f.read()

    print(f"Размер файла: {len(data)} байт")
    print(f"\nПервые 64 байта (HEX):")
    print(" ".join(f"{b:02X}" for b in data[:64]))

    print(f"\n--- Поиск дат в диапазоне 2000-2050 ---")
    found = []
    for offset in range(len(data) - 4):
        # Формат 1: 2 байта год big-endian + месяц + день
        year  = int.from_bytes(data[offset:offset+2], 'big')
        month = data[offset+2]
        day   = data[offset+3]
        if 2000 <= year <= 2050 and 1 <= month <= 12 and 1 <= day <= 31:
            try:
                d = datetime(year, month, day)
                found.append((offset, d, "big-endian"))
            except Exception:
                pass

        # Формат 2: 2 байта год little-endian + месяц + день
        year2 = int.from_bytes(data[offset:offset+2], 'little')
        if 2000 <= year2 <= 2050 and 1 <= month <= 12 and 1 <= day <= 31:
            try:
                d = datetime(year2, month, day)
                found.append((offset, d, "little-endian"))
            except Exception:
                pass

    if found:
        print(f"Найдено {len(found)} совпадений:")
        # Показываем уникальные даты
        unique = {}
        for offset, d, fmt in found:
            key = d.strftime("%Y-%m-%d")
            if key not in unique:
                unique[key] = (offset, d, fmt)
        for key, (offset, d, fmt) in sorted(unique.items()):
            print(f"  offset=0x{offset:04X} ({offset:5d})  дата={d.strftime('%d.%m.%Y')}  формат={fmt}")
        print(f"\nСамая поздняя дата (вероятно дата истечения):")
        latest = max(unique.values(), key=lambda x: x[1])
        print(f"  >>> {latest[1].strftime('%d.%m.%Y')} <<<")
    else:
        print("Даты не найдены методом поиска year+month+day")

    print(f"\n--- Поиск строк с годом 20xx ---")
    # Ищем текстовое представление дат
    text = data.decode("latin-1", errors="replace")
    import re
    matches = re.findall(r'(20\d\d[-/\.]\d\d[-/\.]\d\d|\d\d[-/\.]\d\d[-/\.]20\d\d)', text)
    if matches:
        print(f"Найдены текстовые даты: {set(matches)}")
    else:
        print("Текстовых дат не найдено")

    print(f"\n--- Все байты вокруг возможных дат ---")
    # Ищем байты 0x07 0xE9 (2025 в big-endian) или похожее
    target_years = [2037, 2036, 2035, 2030, 2027, 2026, 2025]
    for year in target_years:
        be = year.to_bytes(2, 'big')
        le = year.to_bytes(2, 'little')
        for enc, label in [(be, "BE"), (le, "LE")]:
            pos = 0
            while True:
                idx = data.find(enc, pos)
                if idx == -1:
                    break
                context = data[max(0,idx-4):idx+8]
                print(f"  {year} ({label}) offset=0x{idx:04X}: {' '.join(f'{b:02X}' for b in context)}")
                pos = idx + 1


# ─── Запуск ───────────────────────────────────────────────────────────────────
print("=" * 60)
print("ДИАГНОСТИКА ID ФАЙЛА LOTUS NOTES")
print("=" * 60)

try_notes_com(ID_FILE_PATH)
read_id_file_raw(ID_FILE_PATH)

input("\nНажмите Enter для выхода...")
