# diag2.py — глубокий анализ Notes TIMEDATE формата в ID файле
# C:\Python313-32\python.exe diag2.py

import struct
from datetime import datetime, timedelta

ID_FILE_PATH = r"C:\Program Files (x86)\IBM\Notes\Data\TECH10.ID"

with open(ID_FILE_PATH, "rb") as f:
    data = f.read()

print(f"Файл: {ID_FILE_PATH}")
print(f"Размер: {len(data)} байт\n")

# ─── Notes TIMEDATE формат ───────────────────────────────────────────────────
# Notes хранит дату как 8 байт: Innards[0] (время) + Innards[1] (дата)
# Innards[1] кодирует: Julian Day Number
# Julian Day для 12.11.2037 = ?

def notes_timedate_to_datetime(innards0, innards1):
    """
    Конвертирует Notes TIMEDATE (два uint32) в datetime.
    Notes epoch: January 1, 4713 BC (Julian calendar)
    Но практически: Notes считает дни с 1 января 1 года
    """
    try:
        # Метод 1: Notes Julian day
        # innards[1] >> 1 = Julian day number
        jd = innards1 >> 1
        # Julian day 2415021 = 1900-01-01
        # Julian day 2440588 = 1970-01-01
        if jd > 2415000:
            days_from_1900 = jd - 2415021
            d = datetime(1900, 1, 1) + timedelta(days=days_from_1900)
            if 1990 <= d.year <= 2060:
                return d
    except Exception:
        pass

    try:
        # Метод 2: просто дни с какой-то эпохи
        # Пробуем разные эпохи
        for epoch_year in [1900, 1899, 1904, 1970, 1601]:
            epoch = datetime(epoch_year, 1, 1)
            for divisor in [1, 100, 1000]:
                try:
                    days = innards1 // divisor
                    d = epoch + timedelta(days=days)
                    if 2000 <= d.year <= 2060:
                        return d
                except Exception:
                    pass
    except Exception:
        pass
    return None


print("=== Поиск Notes TIMEDATE структур (8 байт) ===")
print("Ищем все 8-байтные блоки где результат даёт год 2020-2060\n")

found_dates = []
for offset in range(0, len(data) - 8, 1):
    try:
        i0, i1 = struct.unpack_from('<II', data, offset)
        d = notes_timedate_to_datetime(i0, i1)
        if d and 2020 <= d.year <= 2060:
            found_dates.append((offset, i0, i1, d))
    except Exception:
        pass

if found_dates:
    # Убираем дубликаты по дате
    unique = {}
    for offset, i0, i1, d in found_dates:
        key = d.strftime("%Y-%m-%d")
        if key not in unique:
            unique[key] = (offset, i0, i1, d)

    print(f"Найдено {len(unique)} уникальных дат:")
    for key, (offset, i0, i1, d) in sorted(unique.items()):
        raw = data[offset:offset+8]
        print(f"  offset=0x{offset:04X}  дата={d.strftime('%d.%m.%Y')}  "
              f"raw={' '.join(f'{b:02X}' for b in raw)}")
else:
    print("Даты не найдены через TIMEDATE метод")

# ─── Прямой поиск 2037 ───────────────────────────────────────────────────────
print("\n=== Прямой поиск года 2037 (0x07ED) во всех форматах ===")
y2037_be = (2037).to_bytes(2, 'big')   # 07 ED
y2037_le = (2037).to_bytes(2, 'little') # ED 07

for enc, label in [(y2037_be, "BE 07ED"), (y2037_le, "LE ED07")]:
    pos = data.find(enc)
    if pos >= 0:
        ctx = data[max(0,pos-8):pos+16]
        print(f"  Найден {label} на offset=0x{pos:04X}:")
        print(f"  {' '.join(f'{b:02X}' for b in ctx)}")
    else:
        print(f"  {label} — НЕ НАЙДЕН в файле")

# ─── Полный HEX дамп с анализом каждого TIMEDATE ─────────────────────────────
print("\n=== Все 8-байтные блоки с юлианскими датами 1990-2060 ===")
for offset in range(0, len(data) - 8, 4):  # кратно 4
    try:
        i0, i1 = struct.unpack_from('<II', data, offset)
        # Julian day number
        jd = i1 >> 1
        if 2440000 <= jd <= 2480000:  # примерно 1968-2078
            days = jd - 2440588  # от 1970-01-01
            d = datetime(1970, 1, 1) + timedelta(days=days)
            if 1990 <= d.year <= 2060:
                raw = data[offset:offset+8]
                print(f"  offset=0x{offset:04X}  JD={jd}  дата={d.strftime('%d.%m.%Y')}  "
                      f"raw={' '.join(f'{b:02X}' for b in raw)}")
    except Exception:
        pass

input("\nНажмите Enter для выхода...")
