# diag3.py — точный поиск даты истечения в ID файле
# C:\Python313-32\python.exe diag3.py

import struct
from datetime import datetime, timedelta
from collections import Counter

ID_FILE_PATH = r"C:\Program Files (x86)\IBM\Notes\Data\TECH10.ID"

epoch = datetime(1900, 1, 1)
now   = datetime.now()

with open(ID_FILE_PATH, "rb") as f:
    data = f.read()

print(f"Файл: {ID_FILE_PATH}")
print(f"Размер: {len(data)} байт\n")

# Вычисляем i1 для известных дат
targets = {
    "12.11.2037": (datetime(2037,11,12) - epoch).days,
    "26.10.2037": (datetime(2037,10,26) - epoch).days,
    "08.03.2037": (datetime(2037, 3, 8) - epoch).days,
}

print("=== Ищем точные i1 значения для известных дат ===")
for date_str, i1_val in targets.items():
    # Ищем i1 как little-endian uint32
    needle = struct.pack('<I', i1_val)
    count = 0
    positions = []
    pos = 0
    while True:
        idx = data.find(needle, pos)
        if idx == -1:
            break
        # Проверяем что это реально i1 (второй uint32 в 8-байтном блоке)
        # Значит должен быть на смещении +4 от начала блока
        positions.append(idx)
        count += 1
        pos = idx + 1
    print(f"  {date_str} (i1={i1_val}, hex={i1_val:#06x}): найден {count} раз на offset: {[hex(p) for p in positions[:5]]}")

print()

# Собираем все будущие даты в диапазоне +10..+15 лет
i1_min = (now - epoch).days
i1_max = (datetime(now.year + 15, 12, 31) - epoch).days

print(f"=== Все даты в диапазоне +10..+15 лет ===")
candidates = []
for offset in range(0, len(data) - 8, 1):
    try:
        i0, i1 = struct.unpack_from('<II', data, offset)
        if i1_min < i1 <= i1_max:
            d = (epoch + timedelta(days=i1)).date()
            if d.year >= now.year + 10:
                candidates.append((d, offset, i1))
    except Exception:
        pass

# Уникальные даты с количеством
counts = Counter(d for d, _, _ in candidates)
print(f"Найдено уникальных дат: {len(counts)}")
for d, cnt in sorted(counts.items()):
    offsets = [hex(off) for dd, off, _ in candidates if dd == d][:3]
    print(f"  {d.strftime('%d.%m.%Y')}  встречается {cnt:3d} раз  offsets={offsets}")

print(f"\nМаксимальная дата: {max(counts.keys()).strftime('%d.%m.%Y') if counts else 'нет'}")
print(f"Минимальная дата: {min(counts.keys()).strftime('%d.%m.%Y') if counts else 'нет'}")

input("\nНажмите Enter для выхода...")
