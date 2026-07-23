#!/usr/bin/env python3
"""Заполнение шаблона протокола данными П-674 (электролит М) — lxml."""

import zipfile, os
from lxml import etree

DIR = r"D:\ИИ\Git Projetcs\Nushensia\WorkFlow\Паспорта"
TEMPLATE = os.path.join(DIR, "Шаблоны", "Шаблон_протокола_фазового_анализа.docx")
OUTPUT = os.path.join(DIR, "Шаблоны", "Протокол_П-674_М.docx")

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = '{%s}' % NS_W

with zipfile.ZipFile(TEMPLATE, 'r') as z:
    data = {n: z.read(n) for n in z.namelist()}

doc = etree.fromstring(data['word/document.xml'])

# Replacements
repl = {
    '[номер]': 'П-674', '[индекс]': 'КА', '[год]': '2026',
    '[дд]': '02', '[месяц]': 'июля', '[гггг]': '2026',
    '[заказчик]': 'ТОО «ПетроКазахстан Ойл Продактс»',
    '[наименование_продукции]': 'Cr-Mo сталь, вырезка из трубы змеевика радиантной части печи П-101',
    '[характеристика_образца]': 'П-674, электролит М',
    '[вид_испытаний]': 'Физико-химический фазовый анализ',
    '[нормативные_документы]': 'Методика анализа фазового состава... св-во №01.00225/206-03-2011',
    '[оборудование]': 'Многофункциональный рентгеновский дифрактометр Rigaku Ultima IV',
    '[количество_образцов_даты]': '1 образец, дата испытаний 02.07.2026 г.',
    '[условия_электрохимического_травления]': 'Электролит HCl/CH₃OH, 1:20, -10 °C, 2 А/дм², 2 ч',
    '[режимы_рентгенографирования]': 'Cu Kα₁, Брэгга-Брентано, детектор D/teX Ultra, 40 кВ / 40 мА',
    '[описание_фазового_состава]': 'Фазовый состав: M₃O₄, M₂₃C₆, M₆C(η), M₇C₃, M₂C, M₂O₃, MO',
    '[обозначение]': 'П-674',
    '[список_фаз]': 'оксиды M₃O₄, карбиды M₂₃C₆, η-M₆C, M₇C₃, M₂C, окислы M₂O₃, MO',
    '[количество_фаз]': 'В образце обнаружено 7 фаз.',
    '[размер_ОКР]': '—', '[микроискажения]': '—',
    '[текст_примечания]': 'Обнаружение оксидов (Fe₃O₄, Fe₂O₃, FeO) — аэробное окисление.',
    '{{Должность}}': 'Заместитель технического директора по развитию и науке',
    '{{ФИО}}': 'Добротворская А.Н.',
}

count = 0
for r in doc.iter(W + 'r'):
    for t in r.iter(W + 't'):
        if t.text:
            orig = t.text
            for key, val in repl.items():
                if key in orig:
                    t.text = orig.replace(key, val)
                    count += 1
                    break
print(f"Заменено: {count}")

# Phase table
phase_data = [
    ('1', 'Magnetite', 'Fe\u2082.\u2089\u2081\u2080 O\u2084', 'Cubic', '227 : Fd-3m', '8.387,8.387,8.387,90,90,90', '00-086-1338'),
    ('2', 'Chromium Carbide', 'Cr\u2082\u2083 C\u2086', 'Cubic', '225 : Fm-3m', '10.650,10.650,10.650,90,90,90', '00-071-0552'),
    ('3', 'Nickel Molyb. Carbide', 'Ni\u2083 Mo\u2083 C', 'Cubic', '227 : Fd-3m', '11.050,11.050,11.050,90,90,90', '00-089-4883'),
    ('4', 'Hematite', 'Fe\u2082 O\u2083', 'Trigonal', '167 : R-3c', '5.038,5.038,13.756,90,90,120', '00-085-0599'),
    ('5', 'Molybdenum Carbide', 'Mo\u2082 C', 'Hexagonal', '194 : P6\u2083/mmc', '3.012,3.012,4.735,90,90,120', '00-035-0787'),
    ('6', 'Wuestite, syn', 'Fe O', 'Cubic', '225 : Fm-3m', '4.303,4.303,4.303,90,90,90', '00-075-1550'),
    ('7', 'heptachromium tricarbide', 'Cr\u2087 C\u2083', 'Orthorhombic', '51 : Pmcm', '7.015,12.153,4.532,90,90,90', '00-036-1482'),
]
col_widths = [226, 765, 538, 702, 515, 1785, 470]

for tbl in doc.iter(W + 'tbl'):
    grid = tbl.find(W + 'tblGrid')
    if grid is not None and len(grid.findall(W + 'gridCol')) == 7:
        rows = tbl.findall(W + 'tr')
        if len(rows) > 1:
            tbl.remove(rows[1])
        for pd in phase_data:
            tr = etree.SubElement(tbl, W + 'tr')
            for ci, val in enumerate(pd):
                tc = etree.SubElement(tr, W + 'tc')
                tcPr = etree.SubElement(tc, W + 'tcPr')
                tcW = etree.SubElement(tcPr, W + 'tcW')
                tcW.set(W + 'w', str(col_widths[ci])); tcW.set(W + 'type', 'pct')
                tcB = etree.SubElement(tcPr, W + 'tcBorders')
                for s in ['top','left','bottom','right']:
                    b = etree.SubElement(tcB, W + s)
                    b.set(W + 'val', 'single'); b.set(W + 'sz', '4'); b.set(W + 'space', '0'); b.set(W + 'color', 'auto')
                p = etree.SubElement(tc, W + 'p')
                pPr = etree.SubElement(p, W + 'pPr')
                jc = etree.SubElement(pPr, W + 'jc'); jc.set(W + 'val', 'center')
                sp = etree.SubElement(pPr, W + 'spacing'); sp.set(W + 'line', '240'); sp.set(W + 'lineRule', 'auto')
                r = etree.SubElement(p, W + 'r')
                rPr = etree.SubElement(r, W + 'rPr')
                rf = etree.SubElement(rPr, W + 'rFonts'); rf.set(W + 'ascii','Times New Roman'); rf.set(W + 'hAnsi','Times New Roman')
                sz = etree.SubElement(rPr, W + 'sz'); sz.set(W + 'val', '18')
                t = etree.SubElement(r, W + 't'); t.set('{http://www.w3.org/XML/1998/namespace}space','preserve'); t.text = val
        print(f"Таблица: {len(phase_data)} строк")
        break

data['word/document.xml'] = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)

with zipfile.ZipFile(OUTPUT, 'w') as z:
    for n, c in data.items():
        z.writestr(n, c)
print(f"\n✅ {OUTPUT}")
print(f"   {os.path.getsize(OUTPUT)} байт")
