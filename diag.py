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

    print(f"\n--- Детальный анализ области вокруг найденных дат ---")
    # Выводим 32 байта вокруг каждой найденной позиции
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
                # Показываем 32 байта вокруг
                start = max(0, idx - 16)
                end   = min(len(data), idx + 16)
                chunk = data[start:end]
                hex_str = " ".join(f"{b:02X}" for b in chunk)
                print(f"\n  {year} ({label}) найден на offset=0x{idx:04X} ({idx}):")
                print(f"  HEX  : {hex_str}")

                # Пробуем разные интерпретации даты вокруг этого смещения
                print(f"  Интерпретации:")
                for delta in range(-8, 8):
                    o = idx + delta
                    if o < 0 or o + 6 > len(data):
                        continue
                    # little-endian: год(2) + месяц(1) + день(1)
                    try:
                        y = int.from_bytes(data[o:o+2], 'little')
                        m = data[o+2]
                        d = data[o+3]
                        if 2000 <= y <= 2050 and 1 <= m <= 12 and 1 <= d <= 31:
                            print(f"    LE offset={o}: {d:02d}.{m:02d}.{y} ({' '.join(f'{b:02X}' for b in data[o:o+4])})")
                    except Exception:
                        pass
                    # big-endian: год(2) + месяц(1) + день(1)
                    try:
                        y = int.from_bytes(data[o:o+2], 'big')
                        m = data[o+2]
                        d = data[o+3]
                        if 2000 <= y <= 2050 and 1 <= m <= 12 and 1 <= d <= 31:
                            print(f"    BE offset={o}: {d:02d}.{m:02d}.{y} ({' '.join(f'{b:02X}' for b in data[o:o+4])})")
                    except Exception:
                        pass
                    # Notes TIMEDATE формат: 8 байт
                    try:
                        import struct
                        # Notes TIMEDATE = 2 x uint32, второй = дата
                        td = struct.unpack_from('<II', data, o)
                        # Notes дата: биты 0-5=день, 6-9=месяц, 10-19=год от 1900
                        date_part = td[1]
                        nn_day   = (date_part >> 0)  & 0x3F
                        nn_month = (date_part >> 6)  & 0x0F
                        nn_year  = (date_part >> 10) & 0x3FF
                        if nn_year > 100 and 1 <= nn_month <= 12 and 1 <= nn_day <= 31:
                            real_year = 1900 + nn_year
                            if 2000 <= real_year <= 2060:
                                print(f"    NOTES-TD offset={o}: {nn_day:02d}.{nn_month:02d}.{real_year}")
                    except Exception:
                        pass
                pos = idx + 1

    print(f"\n--- Полный HEX дамп от 0x07C0 до 0x0820 ---")
    start = 0x07C0
    end   = min(len(data), 0x0820)
    chunk = data[start:end]
    for i in range(0, len(chunk), 16):
        row = chunk[i:i+16]
        hex_part = " ".join(f"{b:02X}" for b in row)
        asc_part = "".join(chr(b) if 32 <= b < 127 else "." for b in row)
        print(f"  {start+i:04X}: {hex_part:<48}  {asc_part}")


# ─── Запуск ───────────────────────────────────────────────────────────────────
print("=" * 60)
print("ДИАГНОСТИКА ID ФАЙЛА LOTUS NOTES")
print("=" * 60)

try_notes_com(ID_FILE_PATH)
read_id_file_raw(ID_FILE_PATH)

input("\nНажмите Enter для выхода...")
