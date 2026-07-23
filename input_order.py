#!/usr/bin/env python3
"""Создание карточки заказа. Сканирует Входящие/, спрашивает данные и статус."""

import os, json, datetime, shutil
from pathlib import Path

ROOT = Path(__file__).parent
INBOX = ROOT / "Входящие"
ORDERS = ROOT / "Заказы"

# Какие испытания где проводятся
OUTSOURCED = {
    "ФА": "Физико-химический фазовый анализ",
    "Мех.": "Механические испытания",
    "SEM": "Электронная микроскопия SEM/EDS",
    "SSC": "СКРН (SSC) — NACE TM0177",
    "HIC": "HIC — NACE TM0284",
    "Дл.пр.": "Длительная прочность — ГОСТ 10145",
}

IN_HOUSE = {
    "ВИК": "Визуальный и измерительный контроль",
    "УЗТ": "Ультразвуковая толщинометрия",
    "Хим.": "Химический анализ",
    "Осадок": "Анализ технологического осадка",
    "Оптика": "Металлография (оптическая микроскопия)",
    "Мкротв.": "Микротвёрдость",
    "Твёрд.": "Измерение твёрдости (макро)",
}

def ask(question, default=""):
    prompt = f"{question} " if not default else f"{question} [{default}]: "
    answer = input(prompt).strip()
    return answer if answer else default

def ask_yesno(question):
    answer = input(f"{question} (д/н): ").strip().lower()
    return answer in ("д", "да", "y", "yes")

def select_files():
    """Показать файлы из Входящие/ и дать выбрать."""
    if not INBOX.exists():
        return []
    files = sorted([f for f in INBOX.iterdir() if f.is_file()])
    if not files:
        return []
    print("\nФайлы во Входящие/:")
    for i, f in enumerate(files, 1):
        print(f"  {i:2d}. {f.name}")
    chosen = input("Какие к заказу? (номера через пробел / Enter — все): ").strip()
    if not chosen:
        return files
    selected = []
    for num in chosen.split():
        try:
            idx = int(num) - 1
            if 0 <= idx < len(files):
                selected.append(files[idx])
        except ValueError:
            pass
    return selected

def ask_test_status(test_type, test_name, outsourced=False):
    """Спросить статус одного испытания."""
    where = "сторонняя" if outsourced else "своя"
    print(f"\n  [{test_type}] {test_name} ({where})")
    if outsourced:
        status = ask("  Статус (отдали/результат есть/не нужно)", "не нужно")
        if status == "отдали":
            file = ask("  Файл результата (или Enter если ждём)")
            return {"вид": test_type, "статус": "отдали", "файл": file}
        elif status == "результат есть":
            file = ask("  Файл результата")
            return {"вид": test_type, "статус": "есть", "файл": file}
        else:
            return {"вид": test_type, "статус": "не нужно", "файл": ""}
    else:
        if ask_yesno("  Проведено?"):
            file = ask("  Файл протокола (или Enter)")
            return {"вид": test_type, "статус": "есть", "файл": file}
        else:
            return {"вид": test_type, "статус": "не проведено", "файл": ""}

def main():
    INBOX.mkdir(exist_ok=True)
    ORDERS.mkdir(exist_ok=True)

    print("\n=== НОВЫЙ ЗАКАЗ ===\n")

    # Вариант А: файлы из Входящие/
    files = select_files()

    order = {
        "номер": ask("Номер заказа (напр. П-675)"),
        "заказчик": ask("Заказчик"),
        "объект": ask("Объект исследования"),
        "материал": ask("Марка материала"),
        "дата": datetime.date.today().isoformat(),
        "испытания": [],
        "файлы": [f.name for f in files],
    }

    order_name = f"{order['номер']}_{order['заказчик']}"
    order_dir = ORDERS / order_name

    # Копируем файлы
    if files:
        os.makedirs(order_dir, exist_ok=True)
        print(f"\nКопирую {len(files)} файлов...")
        for f in files:
            dst = order_dir / f.name
            if not dst.exists():
                shutil.copy2(f, dst)
                print(f"  + {f.name}")
            else:
                print(f"  • {f.name} (уже есть)")

    # Статус испытаний: сторонние
    print("\n--- Сторонние организации ---")
    for code, desc in OUTSOURCED.items():
        order["испытания"].append(ask_test_status(code, desc, outsourced=True))

    # Статус испытаний: свои
    print("\n--- Свои испытания ---")
    for code, desc in IN_HOUSE.items():
        order["испытания"].append(ask_test_status(code, desc, outsourced=False))

    # Создаём структуру
    os.makedirs(order_dir, exist_ok=True)
    for sub in ("Протоколы", "Паспорт", "Отчёт"):
        (order_dir / sub).mkdir(exist_ok=True)

    # Сохраняем JSON
    json_path = order_dir / f"{order['номер']}_данные.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(order, f, ensure_ascii=False, indent=2)

    print(f"\n{'='*50}")
    print(f"Заказ:      {order_name}")
    print(f"Папка:      {order_dir}")
    print(f"Данные:     {json_path}")
    print(f"Испытаний:  {len(order['испытания'])}")
    print(f"  Свои:     {sum(1 for t in order['испытания'] if t['вид'] in IN_HOUSE)}")
    print(f"  Сторонние:{sum(1 for t in order['испытания'] if t['вид'] in OUTSOURCED)}")
    print(f"Файлов:     {len(files)}")
    print(f"{'='*50}")

if __name__ == "__main__":
    main()
