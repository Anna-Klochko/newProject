"""
BI-Дашборд агентства недвижимости «Надо Брать»
Версия БЕЗ зависимости от Excel — все данные захардкожены.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# ─── НАСТРОЙКИ СТРАНИЦЫ ────────────────────────────────────────────────────────
st.set_page_config(
    page_title="BI | Надо Брать",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── ЦВЕТОВАЯ ПАЛИТРА ─────────────────────────────────────────────────────────
COLORS = {
    "primary":      "#2D3748",
    "accent":       "#C9A84C",
    "accent_light": "#F0D98C",
    "accent_dark":  "#A07830",
    "bg":           "#F4F6FA",
    "card_bg":      "#FFFFFF",
    "success":      "#38A169",
    "warning":      "#D69E2E",
    "danger":       "#C53030",
    "muted":        "#718096",
    "border":       "#E2D9C8",
    "instagram":    "#C13584",
    "yandex":       "#D4A017",
    "avito":        "#3A7D44",
    "vk":           "#4A76A8",
    "cian":         "#2563EB",
    "telegram":     "#0EA5E9",
}

MONTHS_ORDER = ["январь", "февраль", "март", "апрель"]
MONTHS_RU    = {"январь": "Январь", "февраль": "Февраль", "март": "Март", "апрель": "Апрель"}

PLOT_LAYOUT = dict(
    plot_bgcolor="#FFFFFF",
    paper_bgcolor="#FFFFFF",
    font=dict(family="Inter", size=11, color=COLORS["primary"]),
    hoverlabel=dict(
        bgcolor="white",
        bordercolor=COLORS["border"],
        font=dict(family="Inter", size=12, color=COLORS["primary"]),
    ),
)

# ═══════════════════════════════════════════════════════════════════════════════
# ДАННЫЕ — захардкоженные из Excel
# ═══════════════════════════════════════════════════════════════════════════════

# ── Резервные фонды ────────────────────────────────────────────────────────────
RESERVES = {
    "Резерв рассрочек": 2_083_921.75,
    "Налоговый резерв":  5_844_886.51,
    "Резерв ЧП (5%)":   2_998_162.52,
}

# ── CPL по месяцам (ручные значения) ──────────────────────────────────────────
CPL_BY_MONTH = {"январь": 1467, "февраль": 1918, "март": 1951, "апрель": 2156}

# ── Месячные данные план/факт ──────────────────────────────────────────────────
MONTHLY_DATA = {
    "январь": [
        {"Раздел": "Выручка",                  "Статья": "Выручка",                  "Тип": "Доход",   "План": 13053000.00, "Факт р/с": 5631120.42,  "Уровень": "summary"},
        {"Раздел": "Маркетинг и СММ",           "Статья": "Маркетинг и СММ",           "Тип": "Расход",  "План": 1135500.00,  "Факт р/с": 405807.89,   "Уровень": "summary"},
        {"Раздел": "ЗП",                        "Статья": "ЗП",                        "Тип": "Расход",  "План": 5289610.00,  "Факт р/с": 2366889.59,  "Уровень": "summary"},
        {"Раздел": "Услуги",                    "Статья": "Услуги",                    "Тип": "Расход",  "План": 70000.00,    "Факт р/с": 70247.00,    "Уровень": "summary"},
        {"Раздел": "Подписки",                  "Статья": "Подписки",                  "Тип": "Расход",  "План": 375259.00,   "Факт р/с": 118573.00,   "Уровень": "summary"},
        {"Раздел": "Мотивация",                 "Статья": "Мотивация",                 "Тип": "Расход",  "План": 60000.00,    "Факт р/с": 27000.00,    "Уровень": "summary"},
        {"Раздел": "Административные расходы",  "Статья": "Административные расходы",  "Тип": "Расход",  "План": 2429479.83,  "Факт р/с": 562693.70,   "Уровень": "summary"},
        {"Раздел": "Офис",                      "Статья": "Офис",                      "Тип": "Расход",  "План": 454000.00,   "Факт р/с": 414333.96,   "Уровень": "summary"},
        {"Раздел": "Командировки",              "Статья": "Командировки",              "Тип": "Расход",  "План": 235000.00,   "Факт р/с": 0.00,        "Уровень": "summary"},
        {"Раздел": "Фонды",                     "Статья": "Фонды",                     "Тип": "Расход",  "План": 276060.00,   "Факт р/с": 0.00,        "Уровень": "summary"},
    ],
    "февраль": [
        {"Раздел": "Выручка",                  "Статья": "Выручка",                  "Тип": "Доход",   "План": 13068000.00, "Факт р/с": 19127666.03, "Уровень": "summary"},
        {"Раздел": "Маркетинг и СММ",           "Статья": "Маркетинг и СММ",           "Тип": "Расход",  "План": 973500.00,   "Факт р/с": 1257655.51,  "Уровень": "summary"},
        {"Раздел": "ЗП",                        "Статья": "ЗП",                        "Тип": "Расход",  "План": 5295160.00,  "Факт р/с": 9125340.05,  "Уровень": "summary"},
        {"Раздел": "Услуги",                    "Статья": "Услуги",                    "Тип": "Расход",  "План": 70000.00,    "Факт р/с": 133885.00,   "Уровень": "summary"},
        {"Раздел": "Подписки",                  "Статья": "Подписки",                  "Тип": "Расход",  "План": 375259.00,   "Факт р/с": 117300.00,   "Уровень": "summary"},
        {"Раздел": "Мотивация",                 "Статья": "Мотивация",                 "Тип": "Расход",  "План": 60000.00,    "Факт р/с": 84155.76,    "Уровень": "summary"},
        {"Раздел": "Административные расходы",  "Статья": "Административные расходы",  "Тип": "Расход",  "План": 2432418.33,  "Факт р/с": 2566431.54,  "Уровень": "summary"},
        {"Раздел": "Офисы",                     "Статья": "Офисы",                     "Тип": "Расход",  "План": 454000.00,   "Факт р/с": 726817.22,   "Уровень": "summary"},
        {"Раздел": "Командировки",              "Статья": "Командировки",              "Тип": "Расход",  "План": 235000.00,   "Факт р/с": 170900.00,   "Уровень": "summary"},
        {"Раздел": "Фонды",                     "Статья": "Фонды",                     "Тип": "Расход",  "План": 276360.00,   "Факт р/с": 403357.34,   "Уровень": "summary"},
    ],
    "март": [
        {"Раздел": "Выручка",                  "Статья": "Выручка",                  "Тип": "Доход",   "План": 13068000.00, "Факт р/с": 7957384.35,  "Уровень": "summary"},
        {"Раздел": "Маркетинг и СММ",           "Статья": "Маркетинг и СММ",           "Тип": "Расход",  "План": 1135500.00,  "Факт р/с": 704434.44,   "Уровень": "summary"},
        {"Раздел": "ЗП",                        "Статья": "ЗП",                        "Тип": "Расход",  "План": 5295160.00,  "Факт р/с": 3214944.58,  "Уровень": "summary"},
        {"Раздел": "Услуги",                    "Статья": "Услуги",                    "Тип": "Расход",  "План": 70000.00,    "Факт р/с": 76990.00,    "Уровень": "summary"},
        {"Раздел": "Подписки",                  "Статья": "Подписки",                  "Тип": "Расход",  "План": 375259.00,   "Факт р/с": 301705.00,   "Уровень": "summary"},
        {"Раздел": "Мотивация",                 "Статья": "Мотивация",                 "Тип": "Расход",  "План": 60000.00,    "Факт р/с": 55621.00,    "Уровень": "summary"},
        {"Раздел": "Административные расходы",  "Статья": "Административные расходы",  "Тип": "Расход",  "План": 2488168.33,  "Факт р/с": 1034480.81,  "Уровень": "summary"},
        {"Раздел": "Офисы",                     "Статья": "Офисы",                     "Тип": "Расход",  "План": 629000.00,   "Факт р/с": 1493827.24,  "Уровень": "summary"},
        {"Раздел": "Командировки",              "Статья": "Командировки",              "Тип": "Расход",  "План": 235000.00,   "Факт р/с": 493227.50,   "Уровень": "summary"},
        {"Раздел": "Фонды",                     "Статья": "Фонды",                     "Тип": "Расход",  "План": 276360.00,   "Факт р/с": 158147.69,   "Уровень": "summary"},
    ],
    "апрель": [
        {"Раздел": "Выручка",                  "Статья": "Выручка",                  "Тип": "Доход",   "План": 13083000.00, "Факт р/с": 8150381.44,  "Уровень": "summary"},
        {"Раздел": "Маркетинг и СММ",           "Статья": "Маркетинг и СММ",           "Тип": "Расход",  "План": 1220000.00,  "Факт р/с": 1192812.12,  "Уровень": "summary"},
        {"Раздел": "ЗП",                        "Статья": "ЗП",                        "Тип": "Расход",  "План": 5325710.00,  "Факт р/с": 4570684.83,  "Уровень": "summary"},
        {"Раздел": "Услуги",                    "Статья": "Услуги",                    "Тип": "Расход",  "План": 80000.00,    "Факт р/с": 72269.00,    "Уровень": "summary"},
        {"Раздел": "Подписки",                  "Статья": "Подписки",                  "Тип": "Расход",  "План": 271999.00,   "Факт р/с": 182174.00,   "Уровень": "summary"},
        {"Раздел": "Мотивация",                 "Статья": "Мотивация",                 "Тип": "Расход",  "План": 65000.00,    "Факт р/с": 40400.00,    "Уровень": "summary"},
        {"Раздел": "Административные расходы",  "Статья": "Административные расходы",  "Тип": "Расход",  "План": 2515205.99,  "Факт р/с": 983496.59,   "Уровень": "summary"},
        {"Раздел": "Офисы",                     "Статья": "Офисы",                     "Тип": "Расход",  "План": 859800.00,   "Факт р/с": 169388.92,   "Уровень": "summary"},
        {"Раздел": "Командировки",              "Статья": "Командировки",              "Тип": "Расход",  "План": 230000.00,   "Факт р/с": 93577.50,    "Уровень": "summary"},
        {"Раздел": "Фонды",                     "Статья": "Фонды",                     "Тип": "Расход",  "План": 276660.00,   "Факт р/с": 0.00,        "Уровень": "summary"},
    ],
}

# ── Сделки (реестр) ────────────────────────────────────────────────────────────
DEALS = [
    {"ID":"AN-001","Менеджер":"Лусинэ Хохикян",    "Тип":"Элитная",       "Комиссия":4575288.50,"Факт":4004548.50,"Статус":"Выплачено","Месяц":"январь"},
    {"ID":"AN-002","Менеджер":"Ксения Ким",         "Тип":"Элитная",       "Комиссия":3777501.85,"Факт":3777501.85,"Статус":"Выплачено","Месяц":"февраль"},
    {"ID":"AN-003","Менеджер":"Виктор Жданков",     "Тип":"Элитная",       "Комиссия":3427499.98,"Факт":3427454.08,"Статус":"Выплачено","Месяц":"январь"},
    {"ID":"AN-004","Менеджер":"Маргарита Панова",   "Тип":"Первичка Мск",  "Комиссия":2803904.09,"Факт":2803904.09,"Статус":"Выплачено","Месяц":"февраль"},
    {"ID":"AN-005","Менеджер":"Роман Малиновский",  "Тип":"Первичка Мск",  "Комиссия":2656082.65,"Факт":1906871.23,"Статус":"Ожидается","Месяц":"март"},
    {"ID":"AN-006","Менеджер":"Екатерина Ташкова",  "Тип":"Элитная",       "Комиссия":2561557.50,"Факт":2561557.50,"Статус":"Выплачено","Месяц":"январь"},
    {"ID":"AN-007","Менеджер":"Ярослав Любимов",    "Тип":"Первичка Мск",  "Комиссия":2449849.26,"Факт":2449849.26,"Статус":"Выплачено","Месяц":"февраль"},
    {"ID":"AN-008","Менеджер":"Ангелина Воронцова", "Тип":"Первичка СПб",  "Комиссия":2297666.43,"Факт":1920637.22,"Статус":"Ожидается","Месяц":"март"},
    {"ID":"AN-009","Менеджер":"Снежана Липницкая",  "Тип":"Первичка СПб",  "Комиссия":2287274.76,"Факт":1128226.24,"Статус":"Ожидается","Месяц":"апрель"},
    {"ID":"AN-010","Менеджер":"Элеонора Шерман",    "Тип":"Первичка Мск",  "Комиссия":2180260.76,"Факт":2180260.76,"Статус":"Выплачено","Месяц":"февраль"},
    {"ID":"AN-011","Менеджер":"Кристина Белова",    "Тип":"Вторичка СПб",  "Комиссия":1802660.21,"Факт":1106854.97,"Статус":"Ожидается","Месяц":"март"},
    {"ID":"AN-012","Менеджер":"Георгий Церетели",   "Тип":"Первичка СПб",  "Комиссия":1666585.87,"Факт":1666585.87,"Статус":"Выплачено","Месяц":"январь"},
    {"ID":"AN-013","Менеджер":"Роман Ткаченко",     "Тип":"Первичка Крд",  "Комиссия":1606484.75,"Факт":0.00,      "Статус":"Ожидается","Месяц":"апрель"},
    {"ID":"AN-014","Менеджер":"Илья Резник",        "Тип":"Первичка СПб",  "Комиссия":1578023.14,"Факт":1208193.91,"Статус":"Ожидается","Месяц":"март"},
    {"ID":"AN-015","Менеджер":"Тимур Багратион",    "Тип":"Первичка СПб",  "Комиссия":1496156.41,"Факт":1496156.40,"Статус":"Выплачено","Месяц":"февраль"},
    {"ID":"AN-016","Менеджер":"Марк Левицкий",      "Тип":"Первичка Крд",  "Комиссия":1476400.62,"Факт":1476400.62,"Статус":"Выплачено","Месяц":"январь"},
    {"ID":"AN-017","Менеджер":"Валерия Орлова",     "Тип":"Первичка СПб",  "Комиссия":1338746.56,"Факт":920052.21, "Статус":"Ожидается","Месяц":"апрель"},
    {"ID":"AN-018","Менеджер":"Денис Ковальчук",    "Тип":"Первичка Крд",  "Комиссия":1329247.39,"Факт":1329247.39,"Статус":"Выплачено","Месяц":"март"},
    {"ID":"AN-019","Менеджер":"Владислав Власов",   "Тип":"Вторичка СПб",  "Комиссия":1309953.60,"Факт":1309953.60,"Статус":"Выплачено","Месяц":"февраль"},
    {"ID":"AN-020","Менеджер":"Диана Громова",      "Тип":"Первичка СПб",  "Комиссия":1295607.04,"Факт":828775.03, "Статус":"Ожидается","Месяц":"апрель"},
    {"ID":"AN-021","Менеджер":"Марина Шумкова",     "Тип":"Первичка Крд",  "Комиссия":1286788.16,"Факт":724429.76, "Статус":"Ожидается","Месяц":"март"},
    {"ID":"AN-022","Менеджер":"Алина Берг",         "Тип":"Первичка СПб",  "Комиссия":1257635.19,"Факт":912645.97, "Статус":"Ожидается","Месяц":"апрель"},
    {"ID":"AN-023","Менеджер":"Артур Вершинин",     "Тип":"Первичка СПб",  "Комиссия":1095237.06,"Факт":1095237.05,"Статус":"Выплачено","Месяц":"январь"},
    {"ID":"AN-024","Менеджер":"Арсен Галустян",     "Тип":"Первичка Крд",  "Комиссия":962143.84, "Факт":0.00,      "Статус":"Ожидается","Месяц":"апрель"},
    {"ID":"AN-025","Менеджер":"Оксана Сергиенко",   "Тип":"Первичка Крд",  "Комиссия":936273.04, "Факт":0.00,      "Статус":"Ожидается","Месяц":"апрель"},
    {"ID":"AN-026","Менеджер":"Карина Самойлова",   "Тип":"Первичка Крд",  "Комиссия":631208.74, "Факт":631208.73, "Статус":"Выплачено","Месяц":"март"},
    {"ID":"AN-027","Менеджер":"Алена Коваленко",    "Тип":"Первичка Крд",  "Комиссия":371541.91, "Факт":0.00,      "Статус":"Ожидается","Месяц":"апрель"},
    {"ID":"AN-028","Менеджер":"Дмитрий Кравченко",  "Тип":"Первичка Крд",  "Комиссия":355358.42, "Факт":0.00,      "Статус":"Ожидается","Месяц":"апрель"},
]

# ── Лиды (агрегированные счётчики по каналам и месяцам) ───────────────────────
LEADS_BY_CHANNEL_MONTH = {
    "Instagram":     {"январь":  80, "февраль": 260, "март": 120, "апрель": 150},
    "Яндекс.Директ": {"январь":  22, "февраль":  62, "март":  34, "апрель":  68},
    "Авито":         {"январь":  18, "февраль":  57, "март":  29, "апрель":  33},
    "VK Таргет":     {"январь":  12, "февраль":  52, "март":  29, "апрель":  38},
    "Циан":          {"январь":  11, "февраль":  47, "март":  19, "апрель":  35},
    "Telegram":      {"январь":   7, "февраль":  42, "март":  14, "апрель":  28},
}

CPL_BY_CHANNEL_MONTH = {
    "Instagram":     {"январь": 1100.0, "февраль": 1400.0, "март": 1450.0, "апрель": 1700.0},
    "Яндекс.Директ": {"январь": 1870.1, "февраль": 2408.1, "март": 2458.9, "апрель": 2494.8},
    "Авито":         {"январь": 1923.1, "февраль": 2455.9, "март": 2460.4, "апрель": 2500.6},
    "VK Таргет":     {"январь": 1879.1, "февраль": 2436.2, "март": 2447.8, "апрель": 2497.7},
    "Циан":          {"январь": 1876.1, "февраль": 2456.5, "март": 2444.2, "апрель": 2499.2},
    "Telegram":      {"январь": 1882.1, "февраль": 2423.8, "март": 2452.0, "апрель": 2496.2},
}

# ═══════════════════════════════════════════════════════════════════════════════
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ═══════════════════════════════════════════════════════════════════════════════
def fmt_rub(value: float, short: bool = False) -> str:
    if short:
        if abs(value) >= 1_000_000:
            return f"{value/1_000_000:.2f} млн ₽"
        if abs(value) >= 1_000:
            return f"{value/1_000:.0f} тыс ₽"
        return f"{value:,.0f} ₽"
    return f"{value:,.0f} ₽".replace(",", " ")

def metric_html(label, value, delta="", color_class=""):
    delta_html = ""
    if delta:
        css = "delta-pos" if (delta.startswith("+") or "▲" in delta) else "delta-neg"
        delta_html = f'<div class="metric-delta {css}">{delta}</div>'
    return f"""
    <div class="metric-card {color_class}">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        {delta_html}
    </div>"""

def get_monthly_df(months):
    rows = []
    for m in months:
        for r in MONTHLY_DATA.get(m, []):
            rows.append({**r, "Месяц": m})
    return pd.DataFrame(rows)

def get_deals_df(months):
    rows = [d for d in DEALS if d["Месяц"] in months]
    return pd.DataFrame(rows) if rows else pd.DataFrame(DEALS)

def get_city(typ):
    if "СПб" in typ:   return "Санкт-Петербург"
    if "Мск" in typ:   return "Москва"
    if "Крд" in typ:   return "Краснодар"
    if "Элит" in typ:  return "Элитная недвижимость"
    return "Прочее"

# ═══════════════════════════════════════════════════════════════════════════════
# CSS
# ═══════════════════════════════════════════════════════════════════════════════
def inject_css():
    st.markdown('<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">', unsafe_allow_html=True)
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
    .stApp {{ background-color: #F4F6FA; }}
    .block-container {{ padding-top: 1.2rem !important; max-width: 1400px !important; }}

    .page-header {{
        background: linear-gradient(135deg,#FFFFFF 0%,#FDF6E3 100%);
        border:1px solid {COLORS['border']}; border-left:5px solid {COLORS['accent']};
        padding:22px 30px; border-radius:14px; margin-bottom:22px;
        box-shadow:0 2px 12px rgba(201,168,76,0.12);
    }}
    .page-header h1 {{ margin:0; font-size:1.75rem; font-weight:700; color:{COLORS['primary']}; letter-spacing:-0.02em; }}
    .page-header p  {{ margin:6px 0 0; color:{COLORS['muted']}; font-size:0.9rem; }}
    .gold-line {{ display:inline-block; width:40px; height:3px; background:{COLORS['accent']}; border-radius:2px; margin-bottom:8px; }}

    .metric-card {{
        background:{COLORS['card_bg']}; border-radius:12px; padding:18px 22px 16px;
        box-shadow:0 2px 14px rgba(0,0,0,0.07),0 1px 3px rgba(0,0,0,0.04);
        border:1px solid #EDE8DE; border-top:4px solid {COLORS['accent']};
        margin-bottom:10px; transition:box-shadow 0.2s,transform 0.15s;
    }}
    .metric-card:hover {{ box-shadow:0 6px 24px rgba(201,168,76,0.20); transform:translateY(-1px); }}
    .metric-card.green  {{ border-top-color:{COLORS['success']}; }}
    .metric-card.red    {{ border-top-color:{COLORS['danger']}; }}
    .metric-card.purple {{ border-top-color:#7C3AED; }}
    .metric-card.insta  {{ border-top-color:{COLORS['instagram']}; }}
    .metric-label {{ font-size:0.70rem; color:{COLORS['muted']}; font-weight:600; text-transform:uppercase; letter-spacing:0.08em; margin-bottom:4px; }}
    .metric-value {{ font-size:2.6rem; font-weight:700; color:{COLORS['primary']}; line-height:1.1; margin-top:2px; letter-spacing:-0.02em; }}
    .metric-delta {{ font-size:0.82rem; margin-top:6px; color:{COLORS['muted']}; }}
    .delta-pos {{ color:{COLORS['success']}; font-weight:600; }}
    .delta-neg {{ color:{COLORS['danger']}; font-weight:600; }}

    section[data-testid="stSidebar"] {{ background:#1C1C1E !important; }}
    section[data-testid="stSidebar"] * {{ color:#F5F5F0 !important; }}
    section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] > div {{
        background-color:#2A2A2A !important; border-color:rgba(201,168,76,0.4) !important;
    }}
    section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] span {{
        color:#F0EDE6 !important; font-weight:600 !important; font-size:0.92rem !important;
    }}
    section[data-testid="stSidebar"] .stSelectbox svg {{ fill:#C9A84C !important; }}

    .reserve-box {{
        background:linear-gradient(145deg,#2A2A2A,#333);
        border:1px solid {COLORS['accent']}; border-radius:12px; padding:16px 18px; margin-top:12px;
    }}
    .reserve-box h4 {{ margin:0 0 14px; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.1em; color:{COLORS['accent']} !important; font-weight:700; }}
    .reserve-item {{ display:flex; justify-content:space-between; align-items:center; font-size:0.82rem; margin-bottom:9px; padding-bottom:9px; border-bottom:1px solid rgba(201,168,76,0.15); }}
    .reserve-item:last-child {{ border-bottom:none; margin-bottom:0; padding-bottom:0; }}
    .reserve-item .r-name {{ color:#CCCCCC !important; font-size:0.8rem; }}
    .reserve-item .r-val  {{ font-weight:700; color:{COLORS['accent']} !important; font-size:0.88rem; }}
    .reserve-total {{ margin-top:12px; padding-top:10px; border-top:1px solid {COLORS['accent']}; display:flex; justify-content:space-between; font-size:0.85rem; }}
    .reserve-total .r-name {{ color:#FFFFFF !important; font-weight:600; }}
    .reserve-total .r-val  {{ color:{COLORS['accent_light']} !important; font-weight:800; font-size:0.95rem; }}

    .section-title {{ font-size:0.95rem; font-weight:700; color:{COLORS['primary']}; margin:20px 0 10px; padding-bottom:7px; border-bottom:2px solid #E8E2D4; display:flex; align-items:center; gap:8px; letter-spacing:-0.01em; }}

    .stTabs [data-baseweb="tab-list"] {{ gap:6px; background:#FFFFFF; padding:8px 14px; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,0.06); border:1px solid {COLORS['border']}; }}
    .stTabs [data-baseweb="tab"] {{ border-radius:8px; font-weight:600; font-size:0.88rem; color:{COLORS['muted']} !important; padding:8px 16px !important; }}
    .stTabs [aria-selected="true"] {{ background:linear-gradient(135deg,{COLORS['accent']} 0%,{COLORS['accent_dark']} 100%) !important; color:white !important; box-shadow:0 2px 8px rgba(201,168,76,0.35); }}

    .sales-kpi-card {{ background:#FFFFFF; border-radius:10px; padding:14px 20px 16px; flex:1; min-width:0; box-shadow:0 1px 8px rgba(0,0,0,0.06); border:1px solid #EDE8DE; }}
    .sales-kpi-label {{ font-size:0.78rem; color:#718096; font-weight:400; margin-bottom:6px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
    .sales-kpi-value {{ font-size:1.85rem; font-weight:700; line-height:1.1; letter-spacing:-0.02em; }}
    .sales-kpi-value.blue {{ color:#2B6CB0; }}
    .sales-kpi-value.teal {{ color:#2C7A7B; }}

    .js-plotly-plot {{ border-radius:12px; }}
    [data-testid="stPlotlyChart"] {{ background:#FFFFFF; border-radius:12px; padding:8px; box-shadow:0 1px 8px rgba(0,0,0,0.06); border:1px solid #EDE8DE; margin-bottom:6px; }}

    @media (max-width:768px) {{
        .block-container {{ padding:0.6rem 0.6rem 2rem !important; max-width:100% !important; }}
        .page-header {{ padding:14px 16px; margin-bottom:12px; border-radius:10px; }}
        .page-header h1 {{ font-size:1.15rem !important; }}
        .metric-value {{ font-size:1.7rem !important; }}
        .metric-card {{ padding:14px 14px 12px; }}
        .stTabs [data-baseweb="tab"] {{ padding:6px 10px !important; font-size:0.75rem !important; }}
        .stTabs [data-baseweb="tab-list"] {{ padding:6px 8px; gap:4px; }}
        .section-title {{ font-size:0.85rem; margin:14px 0 8px; }}
        .sales-kpi-row {{ flex-wrap:wrap !important; gap:8px !important; }}
        .sales-kpi-card {{ padding:12px 14px; min-width:calc(50% - 8px) !important; flex:1 1 calc(50% - 8px) !important; }}
        .sales-kpi-value {{ font-size:1.4rem !important; }}
        .sales-kpi-label {{ white-space:normal !important; }}
        [data-testid="column"] {{ width:100% !important; flex:1 1 100% !important; min-width:100% !important; }}
        [data-testid="stPlotlyChart"] {{ padding:4px !important; }}
    }}
    @media (min-width:769px) and (max-width:1024px) {{
        .block-container {{ padding:1rem 1rem 2rem !important; max-width:100% !important; }}
        .metric-value {{ font-size:2rem !important; }}
        .page-header h1 {{ font-size:1.4rem !important; }}
        .stTabs [data-baseweb="tab"] {{ padding:7px 12px !important; font-size:0.82rem !important; }}
        .sales-kpi-value {{ font-size:1.55rem !important; }}
    }}
    </style>
    """, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ВКЛАДКА 1: ФИНАНСЫ
# ═══════════════════════════════════════════════════════════════════════════════
def tab_finance(selected_months):
    df = get_monthly_df(selected_months)
    deals_df = get_deals_df(selected_months)

    выручка_факт = df[df["Статья"] == "Выручка"]["Факт р/с"].sum() if not df.empty else 0
    дебиторка    = deals_df[deals_df["Статус"] == "Ожидается"]["Комиссия"].sum() if not deals_df.empty else 0

    мар_расходы = 0
    if not df.empty:
        мар_расходы = df[df["Статья"] == "Маркетинг и СММ"]["Факт р/с"].sum()

    st.markdown('<div class="section-title">📊 Ключевые финансовые показатели</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(metric_html("Фактическая выручка", fmt_rub(выручка_факт, short=True), color_class="green"), unsafe_allow_html=True)
    with c2:
        st.markdown(metric_html("Маркетинговые расходы", fmt_rub(мар_расходы, short=True)), unsafe_allow_html=True)
    with c3:
        st.markdown(metric_html("Дебиторская задолженность", fmt_rub(дебиторка, short=True), color_class="red"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown('<div class="section-title">📉 Отклонение факта от плана по статьям расходов</div>', unsafe_allow_html=True)
    _chart_plan_fact(selected_months)

    st.markdown('<div class="section-title">💰 Динамика выручки по месяцам</div>', unsafe_allow_html=True)
    _chart_revenue_monthly()


def _chart_plan_fact(selected_months):
    sections_map = {
        "Маркетинг и СММ":        "Маркетинг",
        "ЗП":                     "Зарплаты",
        "Услуги":                 "Услуги",
        "Подписки":               "Подписки",
        "Административные расходы":"Адм. расходы",
        "Офис":                   "Офис",
        "Офисы":                  "Офисы",
        "Командировки":           "Командировки",
    }
    rows_data = []
    for m in selected_months:
        for item in MONTHLY_DATA.get(m, []):
            if item["Уровень"] != "summary" or item["Статья"] not in sections_map:
                continue
            p, f = item["План"], item["Факт р/с"]
            if p == 0:
                continue
            label = f"{sections_map[item['Статья']]} ({MONTHS_RU.get(m,m)})" if len(selected_months) > 1 else sections_map[item['Статья']]
            rows_data.append({"Статья": label, "Отклонение": (f - p) / p * 100, "Факт": f, "План": p})

    if not rows_data:
        st.info("Нет данных.")
        return

    df = pd.DataFrame(rows_data).sort_values("Отклонение", ascending=True)
    bar_colors = [COLORS["danger"] if v > 0 else COLORS["success"] for v in df["Отклонение"]]
    text_labels = [f"{'↑ +' if v > 0 else '↓ '}{v:.1f}%" for v in df["Отклонение"]]
    max_abs = max(abs(df["Отклонение"].max()), abs(df["Отклонение"].min()), 5)
    x_range = [-max_abs * 1.55, max_abs * 1.55]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Отклонение от плана",
        y=df["Статья"], x=df["Отклонение"], orientation="h",
        marker=dict(color=bar_colors, line=dict(color="rgba(0,0,0,0.08)", width=0.5), opacity=0.88),
        text=text_labels, textposition="outside", textfont=dict(size=11, family="Inter"),
        customdata=list(zip(df["План"], df["Факт"])),
        hovertemplate="<b>%{y}</b><br>Отклонение: %{x:.1f}%<br>План: %{customdata[0]:,.0f} ₽<br>Факт: %{customdata[1]:,.0f} ₽<extra></extra>",
    ))
    fig.update_layout(
        **PLOT_LAYOUT, height=max(360, len(df) * 48 + 100), showlegend=False,
        margin=dict(l=12, r=80, t=44, b=16),
        xaxis=dict(title="Отклонение от плана, %", range=x_range, showgrid=True, gridcolor="#F0EDE6",
                   zeroline=True, zerolinecolor="#718096", zerolinewidth=2, ticksuffix="%"),
        yaxis=dict(showgrid=False, tickfont=dict(size=11)),
        annotations=[
            dict(x=x_range[0]*0.7, y=1.06, xref="x", yref="paper", text="← Экономия",
                 showarrow=False, font=dict(color=COLORS["success"], size=11), xanchor="center"),
            dict(x=x_range[1]*0.7, y=1.06, xref="x", yref="paper", text="Перерасход →",
                 showarrow=False, font=dict(color=COLORS["danger"], size=11), xanchor="center"),
        ],
    )
    st.plotly_chart(fig, width="stretch")


def _chart_revenue_monthly():
    months_ru, plans, facts = [], [], []
    for m in MONTHS_ORDER:
        for item in MONTHLY_DATA.get(m, []):
            if item["Статья"] == "Выручка" and item["Уровень"] == "summary":
                months_ru.append(MONTHS_RU[m])
                plans.append(item["План"])
                facts.append(item["Факт р/с"])

    fig = go.Figure()
    fig.add_trace(go.Bar(name="План", x=months_ru, y=plans,
        marker=dict(color="rgba(201,168,76,0.22)", line=dict(color=COLORS["accent"], width=1.5)),
        text=[f"{v/1e6:.1f} млн" for v in plans], textposition="inside",
        textfont=dict(color=COLORS["accent_dark"], size=12),
        hovertemplate="<b>%{x}</b><br>План: %{y:,.0f} ₽<extra></extra>"))
    fig.add_trace(go.Bar(name="Факт", x=months_ru, y=facts,
        marker=dict(color=COLORS["accent"], line=dict(color=COLORS["accent_dark"], width=0.8)),
        text=[f"{v/1e6:.1f} млн" for v in facts], textposition="inside",
        textfont=dict(color="white", size=12),
        hovertemplate="<b>%{x}</b><br>Факт: %{y:,.0f} ₽<extra></extra>"))
    fig.update_layout(
        **PLOT_LAYOUT, barmode="group", height=320, margin=dict(l=12, r=12, t=44, b=16),
        yaxis=dict(title="Выручка, ₽", tickformat=",.0f", showgrid=True, gridcolor="#F0EDE6"),
        xaxis=dict(showgrid=False),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1,
                    bgcolor="rgba(255,255,255,0.9)", bordercolor=COLORS["border"], borderwidth=1))
    st.plotly_chart(fig, width="stretch")


# ═══════════════════════════════════════════════════════════════════════════════
# ВКЛАДКА 2: МАРКЕТИНГ
# ═══════════════════════════════════════════════════════════════════════════════
def tab_marketing(selected_months):
    # Считаем лиды
    total_leads = sum(
        cnt for ch, months in LEADS_BY_CHANNEL_MONTH.items()
        for m, cnt in months.items() if m in selected_months
    )
    cpl_vals = [CPL_BY_MONTH[m] for m in selected_months if m in CPL_BY_MONTH]
    avg_cpl  = sum(cpl_vals) / len(cpl_vals) if cpl_vals else 0
    insta_leads = sum(LEADS_BY_CHANNEL_MONTH["Instagram"].get(m, 0) for m in selected_months)
    insta_share = insta_leads / total_leads * 100 if total_leads > 0 else 0

    st.markdown('<div class="section-title">📱 Ключевые показатели маркетинга</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(metric_html("Всего лидов", f"{total_leads:,}".replace(",", " ")), unsafe_allow_html=True)
    with c2:
        st.markdown(metric_html("Средний CPL", f"{avg_cpl:,.0f} ₽"), unsafe_allow_html=True)
    with c3:
        st.markdown(metric_html("Доля Instagram", f"{insta_share:.1f}%", f"▲ {insta_leads} лидов", color_class="insta"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_left, col_right = st.columns([2, 3])
    with col_left:
        st.markdown('<div class="section-title">🍩 Доли каналов привлечения</div>', unsafe_allow_html=True)
        _chart_channels_donut(selected_months)
    with col_right:
        st.markdown('<div class="section-title">📅 Средняя стоимость лида по месяцам</div>', unsafe_allow_html=True)
        _chart_cpl_monthly()

    st.markdown('<div class="section-title">📈 CPL по каналам привлечения</div>', unsafe_allow_html=True)
    _chart_cpl_channels()


def _chart_channels_donut(selected_months):
    color_map = {
        "Instagram":     COLORS["instagram"],
        "Яндекс.Директ": COLORS["yandex"],
        "Авито":         COLORS["avito"],
        "VK Таргет":     COLORS["vk"],
        "Циан":          COLORS["cian"],
        "Telegram":      COLORS["telegram"],
    }
    channels = list(LEADS_BY_CHANNEL_MONTH.keys())
    counts   = [sum(LEADS_BY_CHANNEL_MONTH[ch].get(m, 0) for m in selected_months) for ch in channels]
    pairs    = [(ch, cnt) for ch, cnt in zip(channels, counts) if cnt > 0]
    labels   = [p[0] for p in pairs]
    values   = [p[1] for p in pairs]
    colors   = [color_map.get(ch, "#95A5A6") for ch in labels]

    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.58,
        marker=dict(colors=colors, line=dict(color="white", width=2.5)),
        textinfo="label+percent",
        hovertemplate="<b>%{label}</b><br>Лидов: %{value}<br>%{percent}<extra></extra>",
        pull=[0.06 if l == "Instagram" else 0 for l in labels],
        textfont=dict(size=11), direction="clockwise", sort=True,
    ))
    total = sum(values)
    fig.add_annotation(
        text=f"<b style='font-size:20px'>{total}</b><br><span style='font-size:11px;color:{COLORS['muted']}'>лидов</span>",
        x=0.5, y=0.5, showarrow=False, align="center", font=dict(color=COLORS["primary"]),
    )
    fig.update_layout(**PLOT_LAYOUT, showlegend=True,
                      legend=dict(orientation="v", x=1.02, y=0.5, font=dict(size=10)),
                      height=380, margin=dict(l=0, r=110, t=20, b=20))
    st.plotly_chart(fig, width="stretch")


def _chart_cpl_monthly():
    months_ru  = [MONTHS_RU[m] for m in MONTHS_ORDER]
    cpl_values = [CPL_BY_MONTH[m] for m in MONTHS_ORDER]
    max_val    = max(cpl_values)
    bar_colors = [COLORS["accent"] if v == max_val else "rgba(201,168,76,0.50)" for v in cpl_values]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=months_ru, y=cpl_values, marker_color=bar_colors,
        marker_line=dict(color=COLORS["accent_dark"], width=0.8),
        text=[f"<b>{v:,} ₽</b>" for v in cpl_values],
        textposition="outside", textfont=dict(size=12, color=COLORS["accent_dark"]),
        showlegend=False, hovertemplate="<b>%{x}</b><br>CPL: %{y:,} ₽<extra></extra>",
    ))
    fig.update_layout(**PLOT_LAYOUT, height=280, margin=dict(l=12, r=12, t=30, b=16),
                      yaxis=dict(range=[0, max_val * 1.35], showgrid=True, gridcolor="#F0EDE6",
                                 ticksuffix=" ₽", tickformat=",.0f"),
                      xaxis=dict(showgrid=False))
    st.plotly_chart(fig, width="stretch")


def _chart_cpl_channels():
    channels = ["Instagram", "Яндекс.Директ", "Авито", "VK Таргет", "Циан"]
    channel_colors = {
        "Instagram":     COLORS["instagram"],
        "Яндекс.Директ": "#D97706",
        "Авито":         COLORS["avito"],
        "VK Таргет":     COLORS["vk"],
        "Циан":          COLORS["cian"],
    }
    channel_dash = {
        "Instagram": "solid", "Яндекс.Директ": "dot",
        "Авито": "dashdot",   "VK Таргет": "dash", "Циан": "longdash",
    }
    months_ru = [MONTHS_RU[m] for m in MONTHS_ORDER]

    fig = go.Figure()
    for ch in channels:
        vals = [CPL_BY_CHANNEL_MONTH[ch].get(m) for m in MONTHS_ORDER]
        is_insta = (ch == "Instagram")
        fig.add_trace(go.Scatter(
            name=ch, x=months_ru, y=vals,
            mode="lines+markers+text",
            line=dict(color=channel_colors[ch], width=3 if is_insta else 2, dash=channel_dash[ch]),
            marker=dict(size=10 if is_insta else 8, color=channel_colors[ch], line=dict(color="white", width=2)),
            text=[f"{v:,.0f} ₽" if v else "" for v in vals],
            textposition="top center",
            textfont=dict(size=10, color=channel_colors[ch]),
            hovertemplate=f"<b>{ch}</b><br>%{{x}}: %{{y:,.0f}} ₽<extra></extra>",
        ))

    fig.add_hline(y=2600, line_dash="dot", line_color=COLORS["danger"], line_width=1.5,
                  annotation_text="⚠️ 2 600 ₽ — пик", annotation_position="top left",
                  annotation_font=dict(color=COLORS["danger"], size=10))
    fig.update_layout(
        **PLOT_LAYOUT, height=480,
        yaxis=dict(title="CPL, ₽", range=[300, 3400], showgrid=True, gridcolor="#F0EDE6",
                   ticksuffix=" ₽", tickformat=",.0f", dtick=400),
        xaxis=dict(showgrid=False),
        legend=dict(orientation="h", yanchor="bottom", y=-0.16, xanchor="center", x=0.5,
                    font=dict(size=11), bgcolor="rgba(255,255,255,0.9)",
                    bordercolor=COLORS["border"], borderwidth=1),
        margin=dict(l=12, r=12, t=30, b=70),
    )
    st.plotly_chart(fig, width="stretch")


# ═══════════════════════════════════════════════════════════════════════════════
# ВКЛАДКА 3: ПРОДАЖИ
# ═══════════════════════════════════════════════════════════════════════════════
def _get_city(typ):
    if "СПб" in typ:   return "Санкт-Петербург"
    if "Мск" in typ:   return "Москва"
    if "Крд" in typ:   return "Краснодар"
    if "Элит" in typ:  return "Элитная недвижимость"
    return "Прочее"


def tab_sales(selected_months):
    deals_df = get_deals_df(selected_months)

    total_leads = sum(
        cnt for ch, months in LEADS_BY_CHANNEL_MONTH.items()
        for m, cnt in months.items() if m in selected_months
    )

    closed_df  = deals_df[deals_df["Статус"] == "Выплачено"] if not deals_df.empty else pd.DataFrame()
    active_df  = deals_df[deals_df["Статус"] == "Ожидается"] if not deals_df.empty else pd.DataFrame()
    n_closed   = len(closed_df)
    sum_closed = closed_df["Факт"].sum() if not closed_df.empty else 0
    sum_active = active_df["Комиссия"].sum() if not active_df.empty else 0
    conv       = n_closed / total_leads * 100 if total_leads > 0 else 0

    # KPI карточки
    st.markdown('<div class="section-title">🎯 Ключевые показатели продаж</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="sales-kpi-row" style="display:flex;flex-wrap:wrap;gap:10px;margin-bottom:18px;">
      <div class="sales-kpi-card"><div class="sales-kpi-label">Сумма всех активных сделок</div><div class="sales-kpi-value blue">{fmt_rub(sum_active)}</div></div>
      <div class="sales-kpi-card"><div class="sales-kpi-label">Успешно закрытые сделки</div><div class="sales-kpi-value teal">{fmt_rub(sum_closed)}</div></div>
      <div class="sales-kpi-card"><div class="sales-kpi-label">Конверсия в продажи</div><div class="sales-kpi-value teal">{conv:.1f}%</div></div>
      <div class="sales-kpi-card"><div class="sales-kpi-label">Закрытых сделок</div><div class="sales-kpi-value blue">{n_closed}</div></div>
    </div>
    """, unsafe_allow_html=True)

    # Воронка
    st.markdown('<div class="section-title">🔽 Воронка продаж</div>', unsafe_allow_html=True)
    _chart_funnel(total_leads, n_closed, selected_months)

    # Рейтинг брокеров
    st.markdown('<div class="section-title">🏆 Рейтинг брокеров по плановой комиссии</div>', unsafe_allow_html=True)
    _chart_broker_rating(deals_df)

    # Рейтинг городов
    st.markdown('<div class="section-title">🗺️ Рейтинг по городам / сегментам</div>', unsafe_allow_html=True)
    _chart_city_rating(deals_df)


def _chart_funnel(total_leads, closed, selected_months):
    if total_leads == 0:
        st.info("Нет данных о лидах.")
        return

    is_only_april = (selected_months == ["апрель"])
    conv_meet = 0.22 if is_only_april else 0.30
    conv_book = 0.35 if is_only_april else 0.42
    conv_deal = 0.45 if is_only_april else 0.52

    if closed == 0:
        closed = max(1, int(total_leads * conv_meet * conv_book * conv_deal))

    paid     = max(closed, int(closed / 0.75))
    approved = max(paid,   int(paid   / conv_deal))
    bookings = max(approved, int(approved / conv_book))
    meetings = max(bookings, int(bookings / conv_book))
    leads_n  = max(meetings, total_leads)

    stages = ["Новые лиды", "Встречи", "Брони", "Согласование", "Оплата получена", "Закрытые сделки"]
    values = [leads_n, meetings, bookings, approved, paid, closed]
    FUNNEL_COLORS = ["#4A90D9", "#7B5EA7", "#E8C84A", "#E87C4A", "#E84A8A", "#5AAD5A"]

    fig = go.Figure(go.Funnel(
        y=stages, x=values,
        textposition="inside", textinfo="label+value",
        textfont=dict(size=13, color="white", family="Inter", weight="bold"),
        hovertemplate=(
            "<b>%{label}</b><br>Количество: <b>%{value}</b><br>"
            "От начала: <b>%{percentInitial:.1%}</b><br>"
            "От предыдущего: <b>%{percentPrevious:.1%}</b><extra></extra>"
        ),
        marker=dict(color=FUNNEL_COLORS, line=dict(width=2.5, color="white")),
        connector=dict(line=dict(color="rgba(200,200,200,0.4)", width=1), fillcolor="rgba(248,248,248,0.6)"),
        opacity=0.95,
    ))

    annotations = []
    if is_only_april:
        annotations = [dict(x=0.5, y=-0.08, xref="paper", yref="paper",
                            text="⚠️ Апрель: снижение качества трафика", showarrow=False,
                            font=dict(color=COLORS["danger"], size=10), align="center")]
    fig.update_layout(**PLOT_LAYOUT, height=460, margin=dict(l=10, r=10, t=20, b=30),
                      annotations=annotations,
                      yaxis=dict(tickfont=dict(size=12, family="Inter", color=COLORS["primary"])))
    st.plotly_chart(fig, width="stretch")


def _chart_broker_rating(deals_df):
    if deals_df.empty:
        st.info("Нет данных.")
        return

    broker_stats = (
        deals_df.groupby("Менеджер")
        .agg(Комиссия=("Комиссия", "sum"), Сделок=("ID", "count"), Тип=("Тип", "first"))
        .reset_index()
        .sort_values("Комиссия", ascending=True)
    )

    def get_region(row):
        t = str(row["Тип"])
        if "Мск" in t:     return "Москва"
        if "Крд" in t:     return "Краснодар"
        if "Элит" in t:    return "Элита"
        return "СПб"

    broker_stats["Регион"] = broker_stats.apply(get_region, axis=1)
    region_palette = {"СПб": COLORS["accent"], "Москва": "#2563EB", "Краснодар": COLORS["success"], "Элита": "#7C3AED"}
    bar_colors = [region_palette.get(r, COLORS["muted"]) for r in broker_stats["Регион"]]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=broker_stats["Комиссия"], y=broker_stats["Менеджер"], orientation="h",
        marker=dict(color=bar_colors, line=dict(color="rgba(255,255,255,0.5)", width=0.5)),
        text=[f"{v/1_000_000:.2f} млн" if v >= 1_000_000 else f"{v/1_000:.0f} тыс" for v in broker_stats["Комиссия"]],
        textposition="outside", textfont=dict(size=10),
        hovertemplate="<b>%{y}</b><br>Комиссия: %{x:,.0f} ₽<extra></extra>",
        showlegend=False,
    ))
    for region, color in region_palette.items():
        fig.add_trace(go.Bar(x=[None], y=[None], orientation="h", name=region, marker_color=color, showlegend=True))

    height = max(520, len(broker_stats) * 28 + 100)
    fig.update_layout(
        **PLOT_LAYOUT, height=height,
        xaxis=dict(title="Плановая комиссия, ₽", tickformat=",.0f", showgrid=True, gridcolor="#F0EDE6", zeroline=False),
        yaxis=dict(showgrid=False, tickfont=dict(size=10)),
        margin=dict(l=12, r=120, t=20, b=20),
        legend=dict(title=dict(text="Регион", font=dict(size=10)), orientation="v", x=1.01, y=1.0,
                    font=dict(size=10), bgcolor="rgba(255,255,255,0.9)", bordercolor=COLORS["border"], borderwidth=1),
        barmode="overlay",
    )
    st.plotly_chart(fig, width="stretch")


def _chart_city_rating(deals_df):
    if deals_df.empty:
        st.info("Нет данных.")
        return

    deals_df = deals_df.copy()
    deals_df["Город"] = deals_df["Тип"].apply(lambda t: _get_city(str(t)))
    city_stats = (
        deals_df.groupby("Город")
        .agg(Комиссия=("Комиссия","sum"), Сделок=("ID","count"),
             Закрытых=("Статус", lambda x: (x=="Выплачено").sum()), Факт=("Факт","sum"))
        .reset_index().sort_values("Комиссия", ascending=False)
    )

    CITY_COLORS = {"Санкт-Петербург":"#4A90D9","Москва":"#7B5EA7","Элитная недвижимость":"#C9A84C","Краснодар":"#5AAD5A","Прочее":"#A0AEC0"}
    CITY_EMOJI  = {"Санкт-Петербург":"🌊","Москва":"🏙️","Элитная недвижимость":"👑","Краснодар":"🌿","Прочее":"📍"}
    max_komis = city_stats["Комиссия"].max()

    cols = st.columns(len(city_stats))
    for i, (_, row) in enumerate(city_stats.iterrows()):
        color  = CITY_COLORS.get(row["Город"], "#A0AEC0")
        emoji  = CITY_EMOJI.get(row["Город"], "📍")
        bar_w  = row["Комиссия"] / max_komis * 100 if max_komis else 0
        komis_m = row["Комиссия"] / 1_000_000
        with cols[i]:
            st.markdown(f"""
            <div style="background:#FFFFFF;border-radius:12px;padding:16px 18px 14px;
                        box-shadow:0 2px 10px rgba(0,0,0,0.07);border:1px solid #EDE8DE;border-top:4px solid {color};">
              <div style="font-size:1.3rem;margin-bottom:4px">{emoji}</div>
              <div style="font-size:0.72rem;color:#718096;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;margin-bottom:2px">{row['Город']}</div>
              <div style="font-size:1.9rem;font-weight:700;color:{color};line-height:1.1;letter-spacing:-0.02em">{komis_m:.2f}<span style="font-size:0.9rem;font-weight:500"> млн ₽</span></div>
              <div style="margin:10px 0 6px;background:#F0EDE6;border-radius:4px;height:6px;overflow:hidden">
                <div style="height:6px;border-radius:4px;background:{color};width:{bar_w:.0f}%"></div>
              </div>
              <div style="margin-top:6px"><span style="font-size:0.78rem;color:#718096">{row['Сделок']} сделок</span></div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    fig = go.Figure()
    colors_list = [CITY_COLORS.get(c, "#A0AEC0") for c in city_stats["Город"]]
    fig.add_trace(go.Bar(
        x=city_stats["Комиссия"], y=city_stats["Город"], orientation="h",
        marker=dict(color=colors_list, line=dict(color="white", width=1.5)),
        text=[f"{v/1e6:.2f} млн ₽" for v in city_stats["Комиссия"]],
        textposition="outside", textfont=dict(size=11),
        hovertemplate="<b>%{y}</b><br>Комиссия: %{x:,.0f} ₽<extra></extra>",
        showlegend=False,
    ))
    fig.update_layout(
        **PLOT_LAYOUT, height=240, margin=dict(l=12, r=160, t=20, b=16),
        xaxis=dict(title="Плановая комиссия, ₽", tickformat=",.0f", showgrid=True, gridcolor="#F0EDE6", zeroline=False),
        yaxis=dict(showgrid=False, tickfont=dict(size=12)),
    )
    st.plotly_chart(fig, width="stretch")


# ═══════════════════════════════════════════════════════════════════════════════
# САЙДБАР
# ═══════════════════════════════════════════════════════════════════════════════
def render_sidebar():
    with st.sidebar:
        st.markdown(
            "<div style='text-align:center;padding:10px 0 4px'>"
            "<span style='font-size:2.4rem'>🏢</span>"
            "<div style='font-size:1rem;font-weight:700;color:#C9A84C;margin-top:4px'>Надо Брать</div>"
            "</div>",
            unsafe_allow_html=True
        )
        st.markdown("<hr style='border-color:rgba(201,168,76,0.3);margin:10px 0'>", unsafe_allow_html=True)

        st.markdown(
            "<div style='font-size:0.75rem;color:#C9A84C;font-weight:700;"
            "text-transform:uppercase;letter-spacing:0.08em;margin-bottom:6px'>📅 Период</div>",
            unsafe_allow_html=True
        )
        month_option = st.selectbox(
            label="Период", options=["Все периоды"] + MONTHS_ORDER,
            index=0, label_visibility="collapsed", key="month_select",
        )
        selected_months = MONTHS_ORDER if month_option == "Все периоды" else [month_option]

        st.markdown("<hr style='border-color:rgba(201,168,76,0.3);margin:14px 0'>", unsafe_allow_html=True)

        # Резервы
        total_res = sum(RESERVES.values())
        items_html = "".join(
            f'<div class="reserve-item"><span class="r-name">{k}</span><span class="r-val">{v:,.2f} ₽</span></div>'
            for k, v in RESERVES.items()
        )
        st.markdown(f"""
        <div class="reserve-box">
            <h4>🏦 Резервные фонды</h4>
            {items_html}
            <div class="reserve-total">
                <span class="r-name">ИТОГО</span>
                <span class="r-val">{total_res:,.2f} ₽</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<hr style='border-color:rgba(201,168,76,0.3);margin:14px 0'>", unsafe_allow_html=True)
        st.markdown(
            "<div style='font-size:0.72rem;color:#888;text-align:center'>2026 © Надо Брать</div>",
            unsafe_allow_html=True
        )

    return selected_months


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    inject_css()
    selected_months = render_sidebar()

    st.markdown("""
    <div class="page-header">
        <div class="gold-line"></div>
        <h1>🏢 Агентство недвижимости «Надо Брать»</h1>
        <p>Интерактивный BI-дашборд</p>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["💰 Финансовые KPI", "📱 KPI маркетинга", "🏆 KPI продаж"])

    period_label = (
        "Все периоды (январь — апрель)"
        if selected_months == MONTHS_ORDER
        else MONTHS_RU.get(selected_months[0], selected_months[0])
    )
    st.caption(f"🗓️ Период: **{period_label}**")

    with tab1:
        tab_finance(selected_months)
    with tab2:
        tab_marketing(selected_months)
    with tab3:
        tab_sales(selected_months)


if __name__ == "__main__":
    main()
