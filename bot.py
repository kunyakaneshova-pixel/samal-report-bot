import os
import io
import json
import tempfile
import traceback
from datetime import datetime
from collections import defaultdict

import requests
from flask import Flask, request
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import ColorScaleRule

app = Flask(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN", "").strip()
TELEGRAM_API = f"https://api.telegram.org/bot{BOT_TOKEN}"
TELEGRAM_FILE_API = f"https://api.telegram.org/file/bot{BOT_TOKEN}"

with open("config.json", "r", encoding="utf-8") as f:
    CONFIG = json.load(f)

CATEGORIES = CONFIG["categories"]
PLANS = CONFIG["plans"]
STORE_MAPPING = CONFIG["store_mapping"]
GROUP_MAPPING = CONFIG["group_mapping"]
DAYS_IN_MONTH = int(CONFIG.get("days_in_month", 30))

LAST_CITY_SUMMARY = None
LAST_STORE_SUMMARY = None
LAST_CATEGORY_SUMMARY = None
LAST_PRODUCT_SUMMARY = None
LAST_PRODUCT_CITY_SUMMARY = None
LAST_AVG_CHECK_SUMMARY = None
LAST_REPORT_META = None
LAST_REPORT_SNAPSHOT = None
PREVIOUS_REPORT_SNAPSHOT = None

TITLE_FILL = "1F4E78"
HEADER_FILL = "D9EAF7"
STORE_FILL = "E2F0D9"


def tg(method, **kwargs):
    r = requests.post(f"{TELEGRAM_API}/{method}", timeout=60, **kwargs)
    r.raise_for_status()
    data = r.json()
    if not data.get("ok"):
        raise RuntimeError(data)
    return data["result"]


MAIN_KEYBOARD = {
    "keyboard": [
        [{"text": "📊 Сделать отчет"}, {"text": "🏙 Сводка по городам"}],
        [{"text": "🏆 Лучшие магазины"}, {"text": "⚠️ Отстающие магазины"}],
        [{"text": "📈 Лучшие категории"}, {"text": "📉 Отстающие категории"}],
        [{"text": "🚀 Рост за день"}, {"text": "🏙 Рост по городам"}],
        [{"text": "🔥 Топ продукции"}, {"text": "🐢 Слабые позиции"}],
        [{"text": "🏙 Топ продукции по городам"}],
        [{"text": "🧾 Средний чек"}, {"text": "↔️ Сравнить со вчера"}],
        [{"text": "📋 Правила объединения"}, {"text": "🎯 Планы"}],
        [{"text": "ℹ️ Помощь"}]
    ],
    "resize_keyboard": True,
    "is_persistent": True
}

def send_message(chat_id, text, keyboard=False):
    payload = {"chat_id": chat_id, "text": text}
    if keyboard:
        payload["reply_markup"] = MAIN_KEYBOARD
    return tg("sendMessage", json=payload)


def send_document(chat_id, path, caption=""):
    with open(path, "rb") as f:
        files = {"document": (os.path.basename(path), f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        data = {"chat_id": str(chat_id), "caption": caption}
        return tg("sendDocument", data=data, files=files)


def parse_iiko(path):
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]

    period_text = str(ws["A3"].value or "")
    m = __import__("re").search(r"по\s+(\d{2})\.(\d{2})\.(\d{4})", period_text)
    if not m:
        raise ValueError("Не смогла определить конечную дату периода в строке A3.")
    day_end = int(m.group(1))
    month = int(m.group(2))
    year = int(m.group(3))

    blocks=[]
    enterprise=None
    subgroup=None
    items=[]

    for row in ws.iter_rows(min_row=4, values_only=True):
        a,b,c,d = row[:4]

        if a and str(a).strip() != "Итого":
            if items and enterprise and subgroup:
                blocks.append((enterprise, subgroup, items))
            enterprise = str(a).strip()
            subgroup = str(b).strip() if b else None
            items=[]

        elif enterprise and b and c is not None:
            if items and subgroup:
                blocks.append((enterprise, subgroup, items))
            subgroup = str(b).strip()
            items=[]

        elif enterprise and b and c is None:
            if items and subgroup:
                blocks.append((enterprise, subgroup, items))
                items=[]
            subgroup=None
            continue

        if enterprise and c is not None and d is not None:
            try:
                items.append((str(c).strip(), float(d)))
            except Exception:
                pass

    if items and enterprise and subgroup:
        blocks.append((enterprise, subgroup, items))

    return blocks, day_end, month, year



def _city_from_enterprise(enterprise):
    s = (enterprise or "").lower()
    if "astana" in s:
        return "Астана"
    if "uralsk" in s:
        return "Уральск"
    if "kyzyl" in s or "qyzyl" in s or "кызыл" in s:
        return "Кызылорда"
    if "kostan" in s or "костан" in s:
        return "Костанай"
    if "aktobe" in s or "актоб" in s or s.startswith("z-aktobe"):
        return "Ақтөбе"
    return "Прочие"


def parse_product_report(path):
    wb = load_workbook(path, data_only=True, read_only=False)
    ws = wb[wb.sheetnames[0]]

    header_row = None
    enterprise_col = None
    dish_col = None
    qty_col = None
    revenue_col = None

    for r in range(1, min(ws.max_row, 50) + 1):
        labels = {}
        for c in range(1, min(ws.max_column, 30) + 1):
            value = ws.cell(r, c).value
            if value is None:
                continue
            labels[c] = str(value).strip().lower()

        for c, label in labels.items():
            if "торговое предприятие" in label:
                enterprise_col = c
            elif label == "блюдо":
                dish_col = c
            elif "количество блюд" in label:
                qty_col = c
            elif "сумма со скидкой" in label:
                revenue_col = c

        if enterprise_col and dish_col and qty_col and revenue_col:
            header_row = r
            break

        enterprise_col = dish_col = qty_col = revenue_col = None

    if not header_row:
        return None, None

    products = defaultdict(lambda: {"qty": 0.0, "revenue": 0.0})
    city_products = defaultdict(lambda: defaultdict(lambda: {"qty": 0.0, "revenue": 0.0}))

    current_enterprise = None

    for r in range(header_row + 1, ws.max_row + 1):
        enterprise = ws.cell(r, enterprise_col).value
        if enterprise:
            current_enterprise = str(enterprise).strip()

        dish = ws.cell(r, dish_col).value
        qty = ws.cell(r, qty_col).value
        revenue = ws.cell(r, revenue_col).value

        if dish is None or qty is None or revenue is None:
            continue

        name = str(dish).strip()
        if not name:
            continue

        try:
            qty = float(qty)
            revenue = float(revenue)
        except (TypeError, ValueError):
            continue

        products[name]["qty"] += qty
        products[name]["revenue"] += revenue

        city = _city_from_enterprise(current_enterprise)
        city_products[city][name]["qty"] += qty
        city_products[city][name]["revenue"] += revenue

    overall = dict(products) if products else None
    by_city = {city: dict(items) for city, items in city_products.items()} if city_products else None
    return overall, by_city


def parse_average_check_report(path):
    """
    Поддерживает два формата iiko:

    1) Специальный отчет по среднему чеку:
       Торговое предприятие | Сумма со скидкой | Чеков | Средняя сумма заказа

    2) Отчет по наименованиям/категориям:
       средний чек считается по итоговой строке предприятия:
       Сумма со скидкой / Чеков
    """
    wb = load_workbook(path, data_only=True, read_only=False)
    ws = wb[wb.sheetnames[0]]

    header_row = None
    enterprise_col = None
    revenue_col = None
    checks_col = None
    avg_col = None

    for r in range(1, min(ws.max_row, 50) + 1):
        ent = rev = chk = avg = None

        for c in range(1, min(ws.max_column, 30) + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            label = str(v).strip().lower()

            if "торговое предприятие" in label:
                ent = c
            elif "сумма со скидкой" in label:
                rev = c
            elif label == "чеков" or ("чек" in label and "кол" in label):
                chk = c
            elif "средняя сумма заказа" in label or ("средн" in label and "чек" in label):
                avg = c

        if ent and rev and chk:
            header_row = r
            enterprise_col = ent
            revenue_col = rev
            checks_col = chk
            avg_col = avg
            break

    if not header_row:
        return None

    result = {}

    # Формат 1: отдельный отчет среднего чека — одна строка = один магазин
    if avg_col:
        for r in range(header_row + 1, ws.max_row + 1):
            enterprise = ws.cell(r, enterprise_col).value
            revenue = ws.cell(r, revenue_col).value
            checks = ws.cell(r, checks_col).value
            avg_check = ws.cell(r, avg_col).value

            if not enterprise:
                continue

            enterprise = str(enterprise).strip()
            if enterprise.lower() in ("итого", "всего") or enterprise.lower().endswith(" всего"):
                continue

            try:
                revenue = float(revenue)
                checks = float(checks)
                avg_check = float(avg_check)
            except (TypeError, ValueError):
                continue

            if checks <= 0:
                continue

            store = _display_store_from_enterprise(enterprise)

            result[store] = {
                "avg_check": avg_check,
                "checks": checks,
                "revenue": revenue,
                "raw_store": enterprise,
            }

        return result if result else None

    # Формат 2: отчет по наименованиям — берём итоговую строку магазина
    for r in range(header_row + 1, ws.max_row + 1):
        enterprise = ws.cell(r, enterprise_col).value
        if not enterprise:
            continue

        enterprise = str(enterprise).strip()

        if not enterprise.lower().endswith(" всего"):
            continue

        group_value = ws.cell(r, 2).value
        dish_value = ws.cell(r, 3).value
        if group_value is not None or dish_value is not None:
            continue

        revenue = ws.cell(r, revenue_col).value
        checks = ws.cell(r, checks_col).value

        try:
            revenue = float(revenue)
            checks = float(checks)
        except (TypeError, ValueError):
            continue

        if checks <= 0:
            continue

        raw_store = enterprise[:-6].strip()
        store = _display_store_from_enterprise(raw_store)

        result[store] = {
            "avg_check": revenue / checks,
            "checks": checks,
            "revenue": revenue,
            "raw_store": raw_store,
        }

    return result if result else None


def analyze(path):
    blocks, day_end, month, year = parse_iiko(path)
    fact_by_store = {store: {c: 0.0 for c in CATEGORIES} for store in PLANS}
    unknown_stores=[]
    unknown_groups=[]

    for enterprise, subgroup, items in blocks:
        key = f"{enterprise}|||{subgroup}"
        store = STORE_MAPPING.get(key)
        if not store:
            unknown_stores.append(f"{enterprise} / {subgroup}")
            continue

        for raw_group, amount in items:
            cat = GROUP_MAPPING.get(raw_group)
            if not cat:
                unknown_groups.append(raw_group)
                continue
            fact_by_store[store][cat] += amount

    if unknown_stores:
        raise ValueError("Неизвестные магазины в новой выгрузке:\n" + "\n".join(sorted(set(unknown_stores))))
    if unknown_groups:
        raise ValueError("Новые группы товаров, которых нет в правилах:\n" + "\n".join(sorted(set(unknown_groups))))

    store_facts = {s: sum(v.values()) for s,v in fact_by_store.items()}
    total_fact = sum(store_facts.values())
    return fact_by_store, store_facts, total_fact, day_end, month, year


def style_title(ws, cell_range, text):
    ws.merge_cells(cell_range)
    c = ws[cell_range.split(":")[0]]
    c.value = text
    c.fill = PatternFill("solid", fgColor=TITLE_FILL)
    c.font = Font(bold=True, color="FFFFFF", size=14)
    c.alignment = Alignment(horizontal="center", vertical="center")


def style_header(row_cells):
    for c in row_cells:
        c.fill = PatternFill("solid", fgColor=HEADER_FILL)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def money(x):
    return float(x or 0)


def build_report(input_path, output_path):
    fact_by_store, store_facts, total_fact, days_fact, month, year = analyze(input_path)
    month_plan = sum(v["plan"] for v in PLANS.values())
    forecast_total = total_fact / days_fact * DAYS_IN_MONTH if days_fact else 0

    wb = Workbook()
    wb.remove(wb.active)

    ws_sum = wb.create_sheet("Сводка")
    ws_cat = wb.create_sheet("Категории")
    ws_store = wb.create_sheet("Магазины")
    ws_detail = wb.create_sheet("По магазинам и категориям")
    ws_all = wb.create_sheet("Все магазины")
    ws_rules = wb.create_sheet("Объединения")

    # Summary
    style_title(ws_sum, "A1:H1", f"Продажи SamalCakes — отчет за 1–{days_fact:02d}.{month:02d}.{year}")
    ws_sum.append([])
    ws_sum.append(["Показатель","Значение"])
    style_header(ws_sum[3])
    rows = [
        ("План месяца", month_plan),
        (f"Факт 1–{days_fact} сентября", total_fact),
        ("Дней факта", days_fact),
        ("Выполнение", total_fact/month_plan if month_plan else 0),
        ("Прогноз месяца", forecast_total),
        ("Прогноз, %", forecast_total/month_plan if month_plan else 0),
        ("Отклонение", forecast_total-month_plan),
        ("Дней осталось", DAYS_IN_MONTH-days_fact),
    ]
    for row in rows:
        ws_sum.append(row)
    for r in [4,5,8,10]:
        ws_sum[f"B{r}"].number_format = '#,##0'
    for r in [7,9]:
        ws_sum[f"B{r}"].number_format = '0.0%'

    # City summary
    city_agg = defaultdict(lambda: {"plan":0.0,"fact":0.0})
    for store in PLANS:
        city=store.split("/")[0]
        city_agg[city]["plan"] += PLANS[store]["plan"]
        city_agg[city]["fact"] += store_facts[store]

    style_title(ws_sum, "A14:H14", "Сводка по городам")
    headers=["Город","План, тг","Факт, тг","Выполнение","Прогноз, тг","Прогноз, %","Отклонение, тг","Доля факта"]
    for col, h in enumerate(headers,1):
        ws_sum.cell(15,col,h)
    style_header(ws_sum[15])

    rr=16
    for city, d in city_agg.items():
        p=d["plan"]; f=d["fact"]; fc=f/days_fact*DAYS_IN_MONTH if days_fact else 0
        vals=[city,p,f,f/p if p else 0,fc,fc/p if p else 0,fc-p,f/total_fact if total_fact else 0]
        for c,v in enumerate(vals,1):
            ws_sum.cell(rr,c,v)
        rr += 1

    # Categories
    style_title(ws_cat, "A1:H1", "Сводка по категориям")
    for c,h in enumerate(headers := ["Категория","План, тг","Факт, тг","Выполнение","Прогноз, тг","Прогноз, %","Отклонение, тг","Доля факта"],1):
        ws_cat.cell(3,c,h)
    style_header(ws_cat[3])
    cat_plan=defaultdict(float); cat_fact=defaultdict(float)
    for store,p in PLANS.items():
        for cat in CATEGORIES:
            cat_plan[cat] += p["categories"].get(cat,0)
            cat_fact[cat] += fact_by_store[store].get(cat,0)

    rr=4
    for cat in CATEGORIES:
        p=cat_plan[cat]; f=cat_fact[cat]; fc=f/days_fact*DAYS_IN_MONTH if days_fact else 0
        vals=[cat,p,f,f/p if p else 0,fc,fc/p if p else 0,fc-p,f/total_fact if total_fact else 0]
        for c,v in enumerate(vals,1):
            ws_cat.cell(rr,c,v)
        rr += 1

    # Stores
    style_title(ws_store, "A1:H1", "Прогноз по магазинам")
    store_headers=["Магазин","Город","План, тг","Факт, тг","Выполнение","Прогноз, тг","Прогноз, %","Отклонение, тг"]
    for c,h in enumerate(store_headers,1):
        ws_store.cell(3,c,h)
    style_header(ws_store[3])

    rr=4
    for store,pdata in PLANS.items():
        city=store.split("/")[0]; p=pdata["plan"]; f=store_facts[store]; fc=f/days_fact*DAYS_IN_MONTH if days_fact else 0
        vals=[store,city,p,f,f/p if p else 0,fc,fc/p if p else 0,fc-p]
        for c,v in enumerate(vals,1):
            ws_store.cell(rr,c,v)
        rr += 1

    # All stores with numbering
    style_title(ws_all, "A1:I1", "Все магазины — контрольный список")
    all_headers=["№","Город","Магазин","План, тг","Факт, тг","Выполнение","Прогноз, тг","Прогноз, %","Отклонение, тг"]
    for c,h in enumerate(all_headers,1):
        ws_all.cell(3,c,h)
    style_header(ws_all[3])
    rr=4
    city_counts=defaultdict(int)
    for n,(store,pdata) in enumerate(PLANS.items(),1):
        city=store.split("/")[0]; city_counts[city]+=1
        p=pdata["plan"]; f=store_facts[store]; fc=f/days_fact*DAYS_IN_MONTH if days_fact else 0
        vals=[n,city,store,p,f,f/p if p else 0,fc,fc/p if p else 0,fc-p]
        for c,v in enumerate(vals,1):
            ws_all.cell(rr,c,v)
        rr += 1
    ws_all["K1"]="Всего магазинов"; ws_all["L1"]=len(PLANS)
    kr=2
    for city,count in city_counts.items():
        ws_all.cell(kr,11,city); ws_all.cell(kr,12,count); kr+=1

    # Detail
    style_title(ws_detail, "A1:H1", "По каждому магазину — план, факт и категории")
    ws_detail["A2"]="Период факта"; ws_detail["B2"]=f"01.{month:02d}.{year}–{days_fact:02d}.{month:02d}.{year}"
    ws_detail["A3"]="Дней факта"; ws_detail["B3"]=days_fact
    ws_detail["A4"]="Дней в месяце"; ws_detail["B4"]=DAYS_IN_MONTH
    detail_headers=["Магазин / категория","План","Факт","Выполнение","Прогноз","Прогноз, %","Отставание","Доля в магазине, %"]
    for c,h in enumerate(detail_headers,1):
        ws_detail.cell(6,c,h)
    style_header(ws_detail[6])
    rr=7
    for store,pdata in PLANS.items():
        p=pdata["plan"]; f=store_facts[store]; fc=f/days_fact*DAYS_IN_MONTH if days_fact else 0
        vals=[store,p,f,f/p if p else 0,fc,fc/p if p else 0,fc-p,1]
        for c,v in enumerate(vals,1):
            ws_detail.cell(rr,c,v)
            ws_detail.cell(rr,c).fill=PatternFill("solid",fgColor=STORE_FILL)
            ws_detail.cell(rr,c).font=Font(bold=True)
        rr += 1
        for cat in CATEGORIES:
            cp=pdata["categories"].get(cat,0); cf=fact_by_store[store].get(cat,0); cfc=cf/days_fact*DAYS_IN_MONTH if days_fact else 0
            vals=[cat,cp,cf,cf/cp if cp else 0,cfc,cfc/cp if cp else 0,cfc-cp,cf/f if f else 0]
            for c,v in enumerate(vals,1):
                ws_detail.cell(rr,c,v)
            rr += 1
        rr += 1

    # Rules
    style_title(ws_rules, "A1:C1", "Правила объединения категорий")
    for c,h in enumerate(["Исходная группа iiko","Категория отчета","Примечание"],1):
        ws_rules.cell(3,c,h)
    style_header(ws_rules[3])
    rr=4
    for raw,cat in GROUP_MAPPING.items():
        ws_rules.cell(rr,1,raw); ws_rules.cell(rr,2,cat); rr+=1

    # Formatting
    for ws in [ws_sum, ws_cat, ws_store, ws_all, ws_detail]:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value,(int,float)):
                    if cell.column in [4,6,8] and ws.title in ["Категории","Сводка","По магазинам и категориям"]:
                        pass

    # Sheet-specific formats
    for ws, start_row, end_row, money_cols, pct_cols in [
        (ws_sum,16,15+len(city_agg),[2,3,5,7],[4,6,8]),
        (ws_cat,4,3+len(CATEGORIES),[2,3,5,7],[4,6,8]),
        (ws_store,4,3+len(PLANS),[3,4,6,8],[5,7]),
        (ws_all,4,3+len(PLANS),[4,5,7,9],[6,8]),
    ]:
        for r in range(start_row,end_row+1):
            for c in money_cols: ws.cell(r,c).number_format='#,##0'
            for c in pct_cols: ws.cell(r,c).number_format='0.0%'

    for r in range(7, ws_detail.max_row+1):
        for c in [2,3,5,7]: ws_detail.cell(r,c).number_format='#,##0'
        for c in [4,6,8]: ws_detail.cell(r,c).number_format='0.0%'

    # Widths and freeze panes
    for ws in [ws_sum,ws_cat,ws_store,ws_all,ws_detail,ws_rules]:
        ws.freeze_panes = "A4" if ws.title in ["Магазины","Все магазины"] else ("A7" if ws.title=="По магазинам и категориям" else None)
        for col in range(1, ws.max_column+1):
            letter=__import__("openpyxl").utils.get_column_letter(col)
            max_len=0
            for cell in ws[letter]:
                if cell.value is not None:
                    max_len=max(max_len,len(str(cell.value)))
            ws.column_dimensions[letter].width=min(max(max_len+2,12),36)

    # Conditional scales
    for ws, rng in [
        (ws_sum, f"F16:F{15+len(city_agg)}"),
        (ws_cat, f"F4:F{3+len(CATEGORIES)}"),
        (ws_store, f"G4:G{3+len(PLANS)}"),
        (ws_all, f"H4:H{3+len(PLANS)}"),
        (ws_detail, f"F7:F{ws_detail.max_row}"),
    ]:
        ws.conditional_formatting.add(rng, ColorScaleRule(
            start_type='min', start_color='F4CCCC',
            mid_type='percentile', mid_value=50, mid_color='FFF2CC',
            end_type='max', end_color='D9EAD3'
        ))

    wb.save(output_path)

    city_summary = {}
    for city, d in city_agg.items():
        p = d["plan"]
        f = d["fact"]
        fc = f / days_fact * DAYS_IN_MONTH if days_fact else 0
        city_summary[city] = {
            "plan": p,
            "fact": f,
            "forecast": fc,
            "forecast_pct": (fc / p if p else 0),
        }

    store_summary = {}
    for store, pdata in PLANS.items():
        p = pdata["plan"]
        f = store_facts[store]
        fc = f / days_fact * DAYS_IN_MONTH if days_fact else 0
        store_summary[store] = {
            "plan": p,
            "fact": f,
            "forecast": fc,
            "forecast_pct": (fc / p if p else 0),
            "deviation": fc - p,
        }

    category_summary = {}
    for cat in CATEGORIES:
        p = cat_plan[cat]
        f = cat_fact[cat]
        fc = f / days_fact * DAYS_IN_MONTH if days_fact else 0
        category_summary[cat] = {
            "plan": p,
            "fact": f,
            "forecast": fc,
            "forecast_pct": (fc / p if p else 0),
            "deviation": fc - p,
        }

    return total_fact, month_plan, forecast_total, days_fact, city_summary, store_summary, category_summary


@app.get("/")
def home():
    return "Samal Report Bot работает ✅"


@app.get("/setup")
def setup_webhook():
    if not BOT_TOKEN:
        return "BOT_TOKEN не задан", 500
    external = os.environ.get("RENDER_EXTERNAL_URL","").rstrip("/")
    if not external:
        return "RENDER_EXTERNAL_URL пока недоступен", 500
    url = external + "/telegram"
    result = tg("setWebhook", json={"url": url, "drop_pending_updates": True})
    return f"Webhook установлен: {url} -> {result}"


@app.post("/telegram")
def telegram_webhook():
    global LAST_CITY_SUMMARY, LAST_STORE_SUMMARY, LAST_CATEGORY_SUMMARY, LAST_PRODUCT_SUMMARY, LAST_PRODUCT_CITY_SUMMARY, LAST_AVG_CHECK_SUMMARY, LAST_REPORT_META, LAST_REPORT_SNAPSHOT, PREVIOUS_REPORT_SNAPSHOT
    update = request.get_json(silent=True) or {}
    try:
        msg = update.get("message") or {}
        chat = msg.get("chat") or {}
        chat_id = chat.get("id")
        if not chat_id:
            return "ok"

        text = msg.get("text","").strip()

        if text == "/start":
            send_message(
                chat_id,
                "Привет! 👋\n"
                "Я твой помощник по отчетам SamalCakes.\n\n"
                "Нажми «📊 Сделать отчет» или просто отправь свежий Excel-файл (.xlsx) из iiko.",
                keyboard=True
            )
            return "ok"

        if text == "📊 Сделать отчет":
            send_message(
                chat_id,
                "Отправь сюда свежую выгрузку .xlsx из iiko.\n"
                "Я автоматически сделаю полный отчет: сводка, города, все магазины, категории и прогноз.",
                keyboard=True
            )
            return "ok"

        if text == "🏙 Сводка по городам":
            if not LAST_CITY_SUMMARY:
                send_message(
                    chat_id,
                    "Пока нет свежей сводки. Сначала отправь новый Excel-файл из iiko.",
                    keyboard=True
                )
                return "ok"

            lines = ["🏙 Сводка по городам", ""]
            if LAST_REPORT_META:
                lines.append(f"Период: 1–{LAST_REPORT_META['days']} число")
                lines.append("")

            for city, d in LAST_CITY_SUMMARY.items():
                fact = d["fact"]
                forecast = d["forecast"]
                pct = d["forecast_pct"]
                fact_s = f"{fact:,.0f}".replace(",", " ")
                forecast_s = f"{forecast:,.0f}".replace(",", " ")
                lines.append(
                    f"{city}:\n"
                    f"  Факт: {fact_s} тг\n"
                    f"  Прогноз: {forecast_s} тг ({pct:.1%})"
                )

            send_message(chat_id, "\n".join(lines), keyboard=True)
            return "ok"

        if text == "🏆 Лучшие магазины":
            if not LAST_STORE_SUMMARY:
                send_message(
                    chat_id,
                    "Пока нет свежих данных. Сначала отправь новый Excel-файл из iiko.",
                    keyboard=True
                )
                return "ok"

            ranked = sorted(
                LAST_STORE_SUMMARY.items(),
                key=lambda x: x[1]["forecast_pct"],
                reverse=True
            )[:10]

            lines = ["🏆 Топ-10 лучших магазинов по прогнозу выполнения плана", ""]
            for i, (store, d) in enumerate(ranked, 1):
                pct = d["forecast_pct"]
                forecast_s = f"{d['forecast']:,.0f}".replace(",", " ")
                lines.append(f"{i}. {store} — {pct:.1%} | прогноз {forecast_s} тг")

            send_message(chat_id, "\n".join(lines), keyboard=True)
            return "ok"

        if text == "⚠️ Отстающие магазины":
            if not LAST_STORE_SUMMARY:
                send_message(
                    chat_id,
                    "Пока нет свежих данных. Сначала отправь новый Excel-файл из iiko.",
                    keyboard=True
                )
                return "ok"

            ranked = sorted(
                LAST_STORE_SUMMARY.items(),
                key=lambda x: x[1]["forecast_pct"]
            )[:10]

            lines = ["⚠️ Топ-10 отстающих магазинов по прогнозу выполнения плана", ""]
            for i, (store, d) in enumerate(ranked, 1):
                pct = d["forecast_pct"]
                deviation_s = f"{d['deviation']:,.0f}".replace(",", " ")
                lines.append(f"{i}. {store} — {pct:.1%} | отклонение {deviation_s} тг")

            send_message(chat_id, "\n".join(lines), keyboard=True)
            return "ok"

        if text == "📈 Лучшие категории":
            if not LAST_CATEGORY_SUMMARY:
                send_message(
                    chat_id,
                    "Пока нет свежих данных. Сначала отправь новый Excel-файл из iiko.",
                    keyboard=True
                )
                return "ok"

            ranked = sorted(
                LAST_CATEGORY_SUMMARY.items(),
                key=lambda x: x[1]["forecast_pct"],
                reverse=True
            )

            lines = ["📈 Лучшие категории по прогнозу выполнения плана", ""]
            for i, (cat, d) in enumerate(ranked, 1):
                pct = d["forecast_pct"]
                forecast_s = f"{d['forecast']:,.0f}".replace(",", " ")
                lines.append(f"{i}. {cat} — {pct:.1%} | прогноз {forecast_s} тг")

            send_message(chat_id, "\n".join(lines), keyboard=True)
            return "ok"

        if text == "📉 Отстающие категории":
            if not LAST_CATEGORY_SUMMARY:
                send_message(
                    chat_id,
                    "Пока нет свежих данных. Сначала отправь новый Excel-файл из iiko.",
                    keyboard=True
                )
                return "ok"

            ranked = sorted(
                LAST_CATEGORY_SUMMARY.items(),
                key=lambda x: x[1]["forecast_pct"]
            )

            lines = ["📉 Отстающие категории по прогнозу выполнения плана", ""]
            for i, (cat, d) in enumerate(ranked, 1):
                pct = d["forecast_pct"]
                deviation_s = f"{d['deviation']:,.0f}".replace(",", " ")
                lines.append(f"{i}. {cat} — {pct:.1%} | отклонение {deviation_s} тг")

            send_message(chat_id, "\n".join(lines), keyboard=True)
            return "ok"

        if text == "🚀 Рост за день":
            if not LAST_REPORT_SNAPSHOT or not PREVIOUS_REPORT_SNAPSHOT:
                send_message(
                    chat_id,
                    "Для роста за день мне нужны два последовательных отчета: вчерашний и сегодняшний.",
                    keyboard=True
                )
                return "ok"

            cur = LAST_REPORT_SNAPSHOT
            prev = PREVIOUS_REPORT_SNAPSHOT

            if cur["days"] <= prev["days"]:
                send_message(
                    chat_id,
                    "Сначала отправь более свежий отчет за следующий день.",
                    keyboard=True
                )
                return "ok"

            store_changes = []
            for store, cur_fact in cur["store_facts"].items():
                prev_fact = prev["store_facts"].get(store, 0)
                store_changes.append((store, cur_fact - prev_fact))
            store_changes.sort(key=lambda x: x[1], reverse=True)

            cat_changes = []
            for cat, cur_fact in cur["category_facts"].items():
                prev_fact = prev["category_facts"].get(cat, 0)
                cat_changes.append((cat, cur_fact - prev_fact))
            cat_changes.sort(key=lambda x: x[1], reverse=True)

            total_add = cur["total_fact"] - prev["total_fact"]
            total_add_s = f"{total_add:,.0f}".replace(",", " ")

            lines = [
                f"🚀 Рост за день: {prev['days']} → {cur['days']} сентября",
                "",
                f"Продажи за день: {total_add_s} тг",
                "",
                "🏆 Топ-10 магазинов за день:"
            ]
            for i, (store, delta) in enumerate(store_changes[:10], 1):
                delta_s = f"{delta:,.0f}".replace(",", " ")
                share = delta / total_add if total_add else 0
                lines.append(f"{i}. {store} — {delta_s} тг ({share:.1%} дня)")

            lines += ["", "📈 Топ категорий за день:"]
            for i, (cat, delta) in enumerate(cat_changes, 1):
                delta_s = f"{delta:,.0f}".replace(",", " ")
                share = delta / total_add if total_add else 0
                lines.append(f"{i}. {cat} — {delta_s} тг ({share:.1%} дня)")

            send_message(chat_id, "\n".join(lines), keyboard=True)
            return "ok"

        if text == "🏙 Рост по городам":
            if not LAST_REPORT_SNAPSHOT or not PREVIOUS_REPORT_SNAPSHOT:
                send_message(
                    chat_id,
                    "Для роста по городам мне нужны два последовательных отчета: вчерашний и сегодняшний.",
                    keyboard=True
                )
                return "ok"

            cur = LAST_REPORT_SNAPSHOT
            prev = PREVIOUS_REPORT_SNAPSHOT

            if cur["days"] <= prev["days"]:
                send_message(
                    chat_id,
                    "Сначала отправь более свежий отчет за следующий день.",
                    keyboard=True
                )
                return "ok"

            cur_cities = cur.get("city_facts", {})
            prev_cities = prev.get("city_facts", {})

            if not cur_cities:
                send_message(
                    chat_id,
                    "В сохраненном отчете пока нет разбивки по городам. Отправь два новых последовательных отчета после этого обновления.",
                    keyboard=True
                )
                return "ok"

            city_changes = []
            for city, cur_fact in cur_cities.items():
                prev_fact = prev_cities.get(city, 0)
                city_changes.append((city, cur_fact - prev_fact))
            city_changes.sort(key=lambda x: x[1], reverse=True)

            total_add = cur["total_fact"] - prev["total_fact"]

            lines = [
                f"🏙 Рост по городам: {prev['days']} → {cur['days']} сентября",
                ""
            ]
            for i, (city, delta) in enumerate(city_changes, 1):
                delta_s = f"{delta:,.0f}".replace(",", " ")
                share = delta / total_add if total_add else 0
                lines.append(f"{i}. {city} — +{delta_s} тг | {share:.1%} продаж дня")

            send_message(chat_id, "\n".join(lines), keyboard=True)
            return "ok"

        if text == "↔️ Сравнить со вчера":
            if not LAST_REPORT_SNAPSHOT or not PREVIOUS_REPORT_SNAPSHOT:
                send_message(
                    chat_id,
                    "Для сравнения мне нужны два последовательных отчета.\n"
                    "Сначала отправь вчерашний файл, потом сегодняшний. После этого нажми эту кнопку.",
                    keyboard=True
                )
                return "ok"

            cur = LAST_REPORT_SNAPSHOT
            prev = PREVIOUS_REPORT_SNAPSHOT

            if cur["days"] <= prev["days"]:
                send_message(
                    chat_id,
                    "Не вижу более нового отчета. Отправь файл за следующий день и попробуй снова.",
                    keyboard=True
                )
                return "ok"

            total_add = cur["total_fact"] - prev["total_fact"]
            pct_add = (total_add / prev["total_fact"]) if prev["total_fact"] else 0

            store_changes = []
            for store, cur_fact in cur["store_facts"].items():
                prev_fact = prev["store_facts"].get(store, 0)
                store_changes.append((store, cur_fact - prev_fact))
            store_changes.sort(key=lambda x: x[1], reverse=True)

            cat_changes = []
            for cat, cur_fact in cur["category_facts"].items():
                prev_fact = prev["category_facts"].get(cat, 0)
                cat_changes.append((cat, cur_fact - prev_fact))
            cat_changes.sort(key=lambda x: x[1], reverse=True)

            total_add_s = f"{total_add:,.0f}".replace(",", " ")

            lines = [
                f"↔️ Сравнение: 1–{prev['days']} → 1–{cur['days']}",
                "",
                f"Продажи за новый день: +{total_add_s} тг",
                f"Рост накопительного факта: {pct_add:.1%}",
                "",
                "🏆 Больше всего добавили магазины:"
            ]

            for i, (store, delta) in enumerate(store_changes[:5], 1):
                delta_s = f"{delta:,.0f}".replace(",", " ")
                lines.append(f"{i}. {store}: +{delta_s} тг")

            lines.extend(["", "📈 Больше всего добавили категории:"])
            for i, (cat, delta) in enumerate(cat_changes[:5], 1):
                delta_s = f"{delta:,.0f}".replace(",", " ")
                lines.append(f"{i}. {cat}: +{delta_s} тг")

            cur_cities = cur.get("city_facts", {})
            prev_cities = prev.get("city_facts", {})
            if cur_cities:
                city_changes = []
                for city, cur_fact in cur_cities.items():
                    city_changes.append((city, cur_fact - prev_cities.get(city, 0)))
                city_changes.sort(key=lambda x: x[1], reverse=True)

                lines.extend(["", "🏙 По городам за новый день:"])
                for i, (city, delta) in enumerate(city_changes, 1):
                    delta_s = f"{delta:,.0f}".replace(",", " ")
                    lines.append(f"{i}. {city}: +{delta_s} тг")

            send_message(chat_id, "\n".join(lines), keyboard=True)
            return "ok"

        if text == "🔥 Топ продукции":
            if not LAST_PRODUCT_SUMMARY:
                send_message(
                    chat_id,
                    "Пока нет данных по продукции. Отправь файл iiko «по наименованиям.xlsx».",
                    keyboard=True
                )
                return "ok"

            ranked = sorted(
                LAST_PRODUCT_SUMMARY.items(),
                key=lambda x: x[1]["revenue"],
                reverse=True
            )[:15]

            lines = ["🔥 Топ-15 продукции по выручке", ""]
            for i, (name, d) in enumerate(ranked, 1):
                rev = f"{d['revenue']:,.0f}".replace(",", " ")
                qty = f"{d['qty']:,.0f}".replace(",", " ")
                lines.append(f"{i}. {name} — {rev} тг | {qty} шт.")

            send_message(chat_id, "\n".join(lines), keyboard=True)
            return "ok"

        if text == "🐢 Слабые позиции":
            if not LAST_PRODUCT_SUMMARY:
                send_message(
                    chat_id,
                    "Пока нет данных по продукции. Отправь файл iiko «по наименованиям.xlsx».",
                    keyboard=True
                )
                return "ok"

            ranked = sorted(
                [(name, d) for name, d in LAST_PRODUCT_SUMMARY.items() if d["qty"] > 0],
                key=lambda x: (x[1]["qty"], x[1]["revenue"])
            )[:15]

            lines = ["🐢 15 самых слабых позиций по количеству продаж", ""]
            for i, (name, d) in enumerate(ranked, 1):
                rev = f"{d['revenue']:,.0f}".replace(",", " ")
                qty = f"{d['qty']:,.0f}".replace(",", " ")
                lines.append(f"{i}. {name} — {qty} шт. | {rev} тг")

            send_message(chat_id, "\n".join(lines), keyboard=True)
            return "ok"

        if text == "🏙 Топ продукции по городам":
            if not LAST_PRODUCT_CITY_SUMMARY:
                send_message(
                    chat_id,
                    "Пока нет данных по продукции по городам. Сначала отправь файл iiko «по наименованиям.xlsx».",
                    keyboard=True
                )
                return "ok"

            city_order = ["Ақтөбе", "Уральск", "Астана", "Кызылорда", "Костанай"]
            lines = ["🏙 Топ продукции по городам", ""]

            for city in city_order:
                products = LAST_PRODUCT_CITY_SUMMARY.get(city, {})
                if not products:
                    continue

                ranked = sorted(
                    products.items(),
                    key=lambda x: x[1]["revenue"],
                    reverse=True
                )[:10]

                lines.append(f"📍 {city}")
                for i, (name, d) in enumerate(ranked, 1):
                    rev = f"{d['revenue']:,.0f}".replace(",", " ")
                    qty = f"{d['qty']:,.0f}".replace(",", " ")
                    lines.append(f"{i}. {name} — {rev} тг | {qty} шт.")
                lines.append("")

            send_message(chat_id, "\n".join(lines), keyboard=True)
            return "ok"

        if text == "🧾 Средний чек":
            if not LAST_AVG_CHECK_SUMMARY:
                send_message(
                    chat_id,
                    "Для среднего чека мне нужна выгрузка iiko, где есть «Средний чек» или одновременно «Количество чеков» и «Выручка». Отправь такой .xlsx — я запомню данные.",
                    keyboard=True
                )
                return "ok"

            ranked = sorted(
                LAST_AVG_CHECK_SUMMARY.items(),
                key=lambda x: x[1]["avg_check"],
                reverse=True
            )

            lines = ["🧾 Средний чек по точкам", "", "🏆 Самый высокий:"]
            for i, (store, d) in enumerate(ranked[:10], 1):
                avg_s = f"{d['avg_check']:,.0f}".replace(",", " ")
                lines.append(f"{i}. {store} — {avg_s} тг")

            lines += ["", "⚠️ Самый низкий:"]
            for i, (store, d) in enumerate(list(reversed(ranked[-10:])), 1):
                avg_s = f"{d['avg_check']:,.0f}".replace(",", " ")
                lines.append(f"{i}. {store} — {avg_s} тг")

            send_message(chat_id, "\n".join(lines), keyboard=True)
            return "ok"

        if text == "📋 Правила объединения":
            rules = (
                "Правила объединения категорий:\n\n"
                "• Печенье → Чайные наборы, пирожное\n"
                "• Заказные торты → Заказные десерты Samal Premium\n"
                "• Булочки + Самса + Хлебобулочные → Хлебобулочные изделия\n"
                "• Шоколад → Десерты\n"
                "• Горячие напитки + Напитки Актобе + Упаковка + Доставка + Хоз товары + Инвентарь/посуда → Общая номенклатура"
            )
            send_message(chat_id, rules, keyboard=True)
            return "ok"

        if text == "🎯 Планы":
            total = sum(v["plan"] for v in PLANS.values())
            city_plans = defaultdict(float)
            for store, data in PLANS.items():
                city_plans[store.split("/")[0]] += data["plan"]

            lines = [f"Общий план месяца: {total:,.0f} тг".replace(",", " "), ""]
            for city, value in city_plans.items():
                lines.append(f"{city}: {value:,.0f} тг".replace(",", " "))

            send_message(chat_id, "\n".join(lines), keyboard=True)
            return "ok"

        if text == "ℹ️ Помощь":
            send_message(
                chat_id,
                "Как пользоваться ботом:\n\n"
                "1. Нажми «📊 Сделать отчет».\n"
                "2. Отправь свежий файл .xlsx из iiko.\n"
                "3. Подожди немного.\n"
                "4. Я верну готовый Excel-отчет.\n"
                "5. Потом можешь смотреть города, магазины, категории и сравнение со вчера прямо в Telegram.\n\n"
                "Можно и без кнопки — просто отправить файл.",
                keyboard=True
            )
            return "ok"

        doc = msg.get("document")
        if not doc:
            send_message(chat_id, "Отправь, пожалуйста, Excel-файл .xlsx из iiko.")
            return "ok"

        name = doc.get("file_name","")
        if not name.lower().endswith(".xlsx"):
            send_message(chat_id, "Мне нужен файл Excel с расширением .xlsx.")
            return "ok"

        send_message(chat_id, "Файл получила ✅ Считаю отчет...")

        f_info = tg("getFile", json={"file_id": doc["file_id"]})
        file_path = f_info["file_path"]
        content = requests.get(f"{TELEGRAM_FILE_API}/{file_path}", timeout=60).content

        with tempfile.TemporaryDirectory() as td:
            inp = os.path.join(td, "iiko.xlsx")
            out = os.path.join(td, "Samal_Report.xlsx")
            with open(inp,"wb") as f:
                f.write(content)

            product_summary, product_city_summary = parse_product_report(inp)
            avg_check_summary = parse_average_check_report(inp)

            if product_summary:
                LAST_PRODUCT_SUMMARY = product_summary
                LAST_PRODUCT_CITY_SUMMARY = product_city_summary
                if avg_check_summary:
                    LAST_AVG_CHECK_SUMMARY = avg_check_summary

                message = (
                    "Файл по наименованиям обработан ✅\n"
                    "Доступны: «🔥 Топ продукции», «🐢 Слабые позиции», "
                    "«🏙 Топ продукции по городам»"
                )
                if avg_check_summary:
                    message += " и «🧾 Средний чек»."
                else:
                    message += "."

                send_message(chat_id, message, keyboard=True)
                return "ok"

            if avg_check_summary:
                LAST_AVG_CHECK_SUMMARY = avg_check_summary
                send_message(
                    chat_id,
                    f"Отчет по среднему чеку обработан ✅\n"
                    f"Найдено магазинов: {len(avg_check_summary)}\n\n"
                    "Теперь нажми «🧾 Средний чек».",
                    keyboard=True
                )
                return "ok"

            total_fact, plan, forecast, days, city_summary, store_summary, category_summary = build_report(inp, out)
            pct = forecast/plan if plan else 0
            LAST_CITY_SUMMARY = city_summary
            LAST_STORE_SUMMARY = store_summary
            LAST_CATEGORY_SUMMARY = category_summary
            LAST_REPORT_META = {"days": days}

            city_facts_snapshot = {}
            for city, d in city_summary.items():
                city_facts_snapshot[city] = d["fact"]

            current_snapshot = {
                "days": days,
                "total_fact": total_fact,
                "store_facts": {k: v["fact"] for k, v in store_summary.items()},
                "category_facts": {k: v["fact"] for k, v in category_summary.items()},
                "city_facts": city_facts_snapshot,
            }

            if LAST_REPORT_SNAPSHOT is None:
                LAST_REPORT_SNAPSHOT = current_snapshot
            elif days > LAST_REPORT_SNAPSHOT.get("days", 0):
                PREVIOUS_REPORT_SNAPSHOT = LAST_REPORT_SNAPSHOT
                LAST_REPORT_SNAPSHOT = current_snapshot
            elif days == LAST_REPORT_SNAPSHOT.get("days", 0):
                LAST_REPORT_SNAPSHOT = current_snapshot
            else:
                PREVIOUS_REPORT_SNAPSHOT = current_snapshot

            caption = (
                f"Готово ✅\n"
                f"Факт за 1–{days}: {total_fact:,.0f} тг\n"
                f"Прогноз: {forecast:,.0f} тг ({pct:.1%})"
            ).replace(",", " ")

            send_document(chat_id, out, caption=caption)

    except Exception as e:
        traceback.print_exc()
        try:
            chat_id = ((update.get("message") or {}).get("chat") or {}).get("id")
            if chat_id:
                send_message(chat_id, "Не получилось обработать файл.\n\n" + str(e)[:3000])
        except Exception:
            pass

    return "ok"
