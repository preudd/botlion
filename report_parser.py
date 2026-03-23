# -*- coding: utf-8 -*-
"""
Парсер Excel файлов с экспортом чеков для формирования вечернего отчёта.
Ключевые столбцы: I (8) - Тип операции, O (14) - Список позиций, P (15) - Итог, Q (16) - Способ оплаты
"""
import pandas as pd
import re
from datetime import datetime
from collections import defaultdict


def parse_items(text):
    """Парсит содержимое столбца O - список позиций в формате [Name, qty, price, total, ...]"""
    if pd.isna(text) or not str(text).strip():
        return []
    items = []
    parts = re.split(r'\],?\s*\[', str(text))
    for p in parts:
        p = p.strip().strip('[]')
        if not p:
            continue
        fields = [x.strip() for x in p.split(',')]
        if len(fields) >= 4:
            name = fields[0]
            try:
                qty = int(float(fields[1]))
                price = float(fields[2])
                total = float(fields[3])
                items.append({'name': name, 'qty': qty, 'price': price, 'total': total})
            except (ValueError, TypeError):
                pass
    return items


def _line_sum(item):
    """Сумма по строке: количество × цена (3-е поле), как в чеке."""
    q = item['qty']
    p = item['price']
    return q * p


def _parse_date(val):
    if pd.isna(val):
        return None
    if isinstance(val, datetime):
        return val
    s = str(val).strip()
    for fmt in ('%d.%m.%y', '%d.%m.%Y', '%Y-%m-%d', '%d/%m/%y'):
        try:
            return datetime.strptime(s[:10], fmt)
        except ValueError:
            continue
    return None


def _is_weekend(date_val):
    dt = _parse_date(date_val)
    if not dt:
        return False
    return dt.weekday() >= 5


def _combo_unit_price(total_item, qty, price):
    """Цена за единицу комбо для разбивки 1150/1250 и 1400/1500."""
    u = total_item / qty if qty else 0
    if 1100 <= u <= 1300 or 1350 <= u <= 1600:
        return u
    if price is not None:
        pr = float(price)
        if 1100 <= pr <= 1300 or 1350 <= pr <= 1600:
            return pr
    return u


def _combo_ag_vr_split(unit_price, qty, is_weekend_row):
    """
    Билет+Аквагрим / Билет+VR: 1150 будни → 800+350, 1250 выходные → 900+350.
    Если цена в диапазоне — по цене; иначе по дате чека.
    """
    if 1100 <= unit_price <= 1200:
        return 800 * qty, 350 * qty
    if 1200 < unit_price <= 1300:
        return 900 * qty, 350 * qty
    if is_weekend_row:
        return 900 * qty, 350 * qty
    return 800 * qty, 350 * qty


def _combo_all_split(unit_price, qty, is_weekend_row):
    """Все включено: 1400 будни → 700+350+350, 1500 выходные → 800+350+350."""
    if 1350 <= unit_price <= 1450:
        return 700 * qty, 350 * qty, 350 * qty
    if 1450 < unit_price <= 1600:
        return 800 * qty, 350 * qty, 350 * qty
    if is_weekend_row:
        return 800 * qty, 350 * qty, 350 * qty
    return 700 * qty, 350 * qty, 350 * qty


def parse_excel_report(file_path):
    """Парсит Excel файл и формирует данные для отчёта."""
    df = pd.read_excel(file_path)

    COL_DATE = 5
    COL_OPERATION = 8
    COL_ITEMS = 14
    COL_TOTAL = 15
    COL_PAYMENT = 16

    revenue_total = 0
    terminal_total = 0
    cash_total = 0
    receipt_count = 0

    category_qty = defaultdict(int)
    category_sum = defaultdict(float)

    combo_ag_qty = 0
    combo_ag_entry = 0
    combo_ag_ag = 0
    combo_vr_qty = 0
    combo_vr_entry = 0
    combo_vr_vr = 0
    combo_all_qty = 0
    combo_all_entry = 0
    combo_all_ag = 0
    combo_all_vr = 0

    return_total = 0
    advance_dr = 0
    rent = 0
    advance_graduation = 0

    # Позиции, не попавшие в основные категории (наименование как в чеке)
    prochee_detail = defaultdict(lambda: {'qty': 0, 'sum': 0.0})

    for idx, row in df.iterrows():
        op_type = str(row.iloc[COL_OPERATION]).strip() if pd.notna(row.iloc[COL_OPERATION]) else ''
        total = float(row.iloc[COL_TOTAL]) if pd.notna(row.iloc[COL_TOTAL]) else 0
        payment = str(row.iloc[COL_PAYMENT]).strip() if pd.notna(row.iloc[COL_PAYMENT]) else ''

        if op_type == 'Возврат прихода':
            return_total += total
            continue

        if op_type != 'Приход':
            continue

        receipt_count += 1
        revenue_total += total

        if 'безналич' in payment.lower() or 'карт' in payment.lower():
            terminal_total += total
        elif 'налич' in payment.lower():
            cash_total += total

        row_date = row.iloc[COL_DATE]
        weekend = _is_weekend(row_date)
        items = parse_items(row.iloc[COL_ITEMS])

        for item in items:
            name = str(item['name']).strip()
            nl = name.lower()
            qty = item['qty']
            price = item['price']
            total_item = item['total']
            ls = _line_sum(item)

            # --- Комбо (раньше всего) ---
            if 'тройное комбо' in nl:
                combo_all_qty += qty
                up = _combo_unit_price(total_item, qty, price)
                e, a, v = _combo_all_split(up, qty, weekend)
                combo_all_entry += e
                combo_all_ag += a
                combo_all_vr += v
                category_qty['Комбо все включено'] += qty
                category_sum['Комбо все включено'] += ls
                continue

            if 'комбо на др 3 часа' in nl:
                category_qty['Комбо на ДР 3 часа'] += qty
                category_sum['Комбо на ДР 3 часа'] += ls
                continue

            if 'комбо(билет+аквагрим)' in nl or 'комбо(билет+аква)' in nl:
                combo_ag_qty += qty
                up = _combo_unit_price(total_item, qty, price)
                ent, ag = _combo_ag_vr_split(up, qty, weekend)
                combo_ag_entry += ent
                combo_ag_ag += ag
                category_qty['Комбо(Билет+Аквагрим)'] += qty
                category_sum['Комбо(Билет+Аквагрим)'] += ls
                continue

            if 'комбо(билет+viar)' in nl or 'комбо(билет+vr)' in nl:
                combo_vr_qty += qty
                up = _combo_unit_price(total_item, qty, price)
                ent, vr = _combo_ag_vr_split(up, qty, weekend)
                combo_vr_entry += ent
                combo_vr_vr += vr
                category_qty['Комбо(Билет+VR)'] += qty
                category_sum['Комбо(Билет+VR)'] += ls
                continue

            # --- Вход безлимит: только Билет(БЕЗЛИМИТ), любая цена ---
            if 'билет(безлимит)' in nl or 'вход безлимит' in nl:
                category_qty['Вход безлимит'] += qty
                category_sum['Вход безлимит'] += ls
                continue

            # --- Билет 1 час ---
            if 'билет(1час)' in nl or 'билет 1час' in nl:
                category_qty['Билет 1час'] += qty
                category_sum['Билет 1час'] += ls
                continue

            # --- Акции: только 500 ₽ ---
            if abs(float(price) - 500) < 0.01:
                if 'счастливые' in nl and 'час' in nl:
                    category_qty['Акция счастливые часы'] += qty
                    category_sum['Акция счастливые часы'] += ls
                    continue
                if 'последний' in nl and 'час' in nl:
                    category_qty['Акция последний час'] += qty
                    category_sum['Акция последний час'] += ls
                    continue

            # --- Аквагрим: только отдельная позиция, не комбо ---
            if nl == 'аквагрим' or (nl.startswith('аквагрим') and 'комбо' not in nl):
                category_sum['Аквагрим'] += ls
                continue

            # --- Viar: только отдельная позиция ---
            if nl == 'viar' or (nl.startswith('viar') and 'комбо' not in nl):
                category_sum['Виар'] += ls
                continue

            # --- Шары: только «Шар» (и опечатка Щар) ---
            if nl in ('шар', 'щар'):
                category_sum['Шары'] += ls
                continue

            # --- Сопровождающий: только явное название ---
            if 'сопровождающ' in nl:
                category_sum['Сопровождающий'] += ls
                continue

            # --- Прочие из маппинга ---
            if 'аванс др' in nl or 'аванс день рождения' in nl:
                advance_dr += ls
                continue
            if 'аренда комнаты' in nl:
                rent += ls
                continue
            if 'аванс выпускной' in nl:
                advance_graduation += ls
                continue

            category_qty['Прочее'] += qty
            category_sum['Прочее'] += ls
            prochee_detail[name]['qty'] += qty
            prochee_detail[name]['sum'] += ls

    return {
        'revenue': revenue_total,
        'terminal': terminal_total,
        'cash': cash_total,
        'receipt_count': receipt_count,
        'category_qty': dict(category_qty),
        'category_sum': dict(category_sum),
        'return_total': return_total,
        'combo_ag': (combo_ag_qty, combo_ag_entry, combo_ag_ag),
        'combo_vr': (combo_vr_qty, combo_vr_entry, combo_vr_vr),
        'combo_all': (combo_all_qty, combo_all_entry, combo_all_ag, combo_all_vr),
        'advance_dr': advance_dr,
        'rent': rent,
        'advance_graduation': advance_graduation,
        'prochee_detail': {k: dict(v) for k, v in prochee_detail.items()},
    }


def _format_prochee_block(data, rub_fn):
    """Текст блока «Прочее» с разбивкой по наименованиям."""
    detail = data.get('prochee_detail') or {}
    if not detail:
        return "Прочее: позиций нет."
    lines = ["Прочее (не вошло в основные категории):"]
    total = 0.0
    for name in sorted(detail.keys(), key=str.lower):
        d = detail[name]
        q = int(d['qty'])
        s = float(d['sum'])
        total += s
        lines.append(f"• {name} — {q} шт {rub_fn(s)} руб")
    lines.append(f"Итого прочее: {rub_fn(total)} руб")
    return "\n".join(lines)


def format_report(data, report_date=None):
    """Форматирует отчёт в текстовый вид по шаблону."""
    if report_date is None:
        report_date = datetime.now().strftime('%d.%m.%Y')

    def rub(val):
        return f'{val:,.0f}'.replace(',', ' ') if isinstance(val, (int, float)) else '0'

    cq = data['category_qty']
    cs = data['category_sum']
    ca = data['combo_ag']
    cv = data['combo_vr']
    call = data['combo_all']
    prochee_block = _format_prochee_block(data, rub)

    report = f"""{report_date}

За сегодня:

Выручка -  {rub(data['revenue'])} руб.
Терминал -  {rub(data['terminal'])} руб.
Наличка приход - {rub(data['cash'])} руб.

Количество чеков - {data['receipt_count']} шт. 

Вход безлимит - {cq.get('Вход безлимит', 0)} шт {rub(cs.get('Вход безлимит', 0))} руб
Билет 1час – {cq.get('Билет 1час', 0)} шт {rub(cs.get('Билет 1час', 0))} руб
Акция счастливые часы - {cq.get('Акция счастливые часы', 0)} шт {rub(cs.get('Акция счастливые часы', 0))} руб 
Акция последний час - {cq.get('Акция последний час', 0)} шт {rub(cs.get('Акция последний час', 0))} руб

Аквагрим - {rub(cs.get('Аквагрим', 0))} руб
Виар - {rub(cs.get('Виар', 0))} руб
Шары - {rub(cs.get('Шары', 0))} руб
Сопровождающий - {rub(cs.get('Сопровождающий', 0))} руб

{prochee_block}

Комбо(Билет+Аквагрим): {int(ca[0])}шт 
Вход – {rub(ca[1])}
Аквагрим – {rub(ca[2])}

Комбо(Билет+VR): {int(cv[0])} шт 
Вход – {rub(cv[1])}
VR – {rub(cv[2])}

Комбо все включено : {int(call[0])} шт
Вход - {rub(call[1])}
Аквагрим - {rub(call[2])}
VR - {rub(call[3])}

Возврат - {rub(data['return_total'])}

Аванс ДР - {rub(data['advance_dr'])}
Аренда комнаты - {rub(data['rent'])}
Аванс выпускной - {rub(data['advance_graduation'])}
Комбо на ДР 3 часа - {cq.get('Комбо на ДР 3 часа', 0)} шт {rub(cs.get('Комбо на ДР 3 часа', 0))}

Касса: 
Инкассация: 
Остаток в кассе: """

    return report.strip()
