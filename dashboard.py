import json
import os
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import load_workbook
from datetime import datetime

st.set_page_config(page_title="Гант Гарант 2026", layout="wide")

# ===== НАСТРОЙКИ =====
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1DlgMbUkXySIBtQIT8n0k0hdN5qJeYxBFpc-R8a9FYm4/edit?usp=sharing"
SHEET_NAME = "Диаграмма Ганта Гарнт 2026"

# ===== СТИЛИ =====
st.markdown("""
<style>
.card {
    padding: 22px;
    border-radius: 18px;
    color: #111;
    font-size: 20px;
    font-weight: 600;
    box-shadow: 0 4px 14px rgba(0,0,0,0.08);
}
.card-number {
    font-size: 38px;
    font-weight: 800;
    margin-top: 8px;
}
.red { background: #ffd6d6; }
.yellow { background: #fff1b8; }
.green { background: #d8f5d0; }
.gray { background: #eeeeee; }
</style>
""", unsafe_allow_html=True)



def read_google_sheet(sheet_url):
    if os.path.exists("credentials.json"):
        # Локальный запуск на твоём компьютере
        gc = gspread.service_account(filename="credentials.json")
    else:
        # Запуск в Streamlit Cloud
        service_account_info = dict(st.secrets["gcp_service_account"])
        service_account_info["private_key"] = service_account_info["private_key"].replace("\\n", "\n")
        gc = gspread.service_account_from_dict(service_account_info)

    sh = gc.open_by_url(sheet_url)
    ws = sh.worksheet(SHEET_NAME)

    data = ws.get_all_values()
    return pd.DataFrame(data)

def parse_date(value):
    if pd.isna(value) or value == "":
        return pd.NaT
    return pd.to_datetime(value, errors="coerce")


def parse_percent(value):
    if pd.isna(value) or value == "":
        return pd.NA

    if isinstance(value, str):
        value = value.replace("%", "").replace(",", ".").strip()
        try:
            number = float(value)
        except ValueError:
            return pd.NA
    elif isinstance(value, (int, float)):
        number = float(value)
    else:
        return pd.NA

    if number > 1:
        return number / 100

    return number


def prepare_data(df):
    gantt = pd.DataFrame()

    # A = 0, B = 1, H = 7, I = 8, L = 11, M = 12
    gantt["A"] = df.iloc[:, 0]
    gantt["B"] = df.iloc[:, 1]
    gantt["H_deadline"] = df.iloc[:, 7]
    gantt["I_fact"] = df.iloc[:, 8]
    gantt["L_status"] = df.iloc[:, 11]
    gantt["M_progress"] = df.iloc[:, 12]

    # с 6 строки Excel
    gantt = gantt.iloc[5:].copy()

    gantt["deadline"] = gantt["H_deadline"].apply(parse_date)
    gantt["fact_date"] = gantt["I_fact"].apply(parse_date)
    gantt["progress"] = gantt["M_progress"].apply(parse_percent)

    gantt["status"] = gantt["L_status"].astype(str).str.strip()
    gantt["task_name"] = gantt["B"].astype(str).str.strip()

    # Этап — когда в A число
    gantt["is_stage"] = pd.to_numeric(gantt["A"], errors="coerce").notna()

    gantt["stage_num"] = pd.NA
    gantt["stage_name"] = pd.NA

    gantt.loc[gantt["is_stage"], "stage_num"] = gantt.loc[gantt["is_stage"], "A"]
    gantt.loc[gantt["is_stage"], "stage_name"] = gantt.loc[gantt["is_stage"], "B"]

    gantt["stage_num"] = gantt["stage_num"].ffill()
    gantt["stage_name"] = gantt["stage_name"].ffill()

    # Задача — строка, где A текстовая и не пустая
    gantt["is_task"] = (
        gantt["A"].notna()
        & ~gantt["is_stage"]
        & (gantt["A"].astype(str).str.strip() != "")
    )

    return gantt


def calculate_kpi(gantt, full_df):
    today = pd.Timestamp.today().normalize()
    three_days = today + pd.Timedelta(days=3)

    has_task_name = gantt["B"].notna() & (gantt["B"].astype(str).str.strip() != "")
    not_completed = gantt["status"].str.lower() != "завершено"

    overdue = gantt[
        (gantt["deadline"] < today)
        & gantt["deadline"].notna()
        & not_completed
        & has_task_name
    ].shape[0]

    risk = gantt[
        (gantt["deadline"] >= today)
        & (gantt["deadline"] <= three_days)
        & gantt["deadline"].notna()
        & not_completed
        & has_task_name
    ].shape[0]

    completed = gantt[
        (gantt["status"].str.lower() == "завершено")
        & has_task_name
    ].shape[0]

    progress_values = full_df.iloc[2:, 12].apply(parse_percent).dropna()
    project_progress = progress_values.mean() if len(progress_values) else 0

    return overdue, risk, completed, project_progress


# ===== ИНТЕРФЕЙС =====
st.title("📊 Гант Гарант 2026")

sheet_url = st.text_input(
    "Ссылка на Google Таблицу",
    value=GOOGLE_SHEET_URL
)

full_df = read_google_sheet(sheet_url)
gantt = prepare_data(full_df)

overdue, risk, completed, project_progress = calculate_kpi(gantt, full_df)

st.subheader("🎛️ Дашборд")

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.markdown(f"""
    <div class="card red">
        🔴 Просроченные
        <div class="card-number">{overdue}</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="card yellow">
        ⚠️ Риск 3 дня
        <div class="card-number">{risk}</div>
    </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown(f"""
    <div class="card green">
        ✅ Завершено
        <div class="card-number">{completed}</div>
    </div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown(f"""
    <div class="card gray">
        📈 Прогресс проекта
        <div class="card-number">{project_progress:.0%}</div>
    </div>
    """, unsafe_allow_html=True)

st.divider()

# ===== СВОДНАЯ ПО ЭТАПАМ =====
st.subheader("🚀 Сводная по этапам")

tasks = gantt[gantt["is_task"]].copy()

stages = sorted(tasks["stage_name"].dropna().unique())

selected_stages = st.multiselect(
    "Фильтр по этапам",
    options=stages,
    default=stages
)

filtered_tasks = tasks[tasks["stage_name"].isin(selected_stages)]

stage_summary = (
    filtered_tasks
    .groupby(["stage_num", "stage_name"], dropna=False)
    .agg(
        **{
            "Кол-во задач": ("task_name", "count"),
            "Прогресс": ("progress", "mean")
        }
    )
    .reset_index()
)

stage_summary["Прогресс"] = stage_summary["Прогресс"].fillna(0)
stage_summary["Прогресс, %"] = (stage_summary["Прогресс"] * 100).round(1)

if not stage_summary.empty:
    st.dataframe(
        stage_summary[
            ["stage_num", "stage_name", "Кол-во задач", "Прогресс, %"]
        ].rename(columns={
            "stage_num": "Номер этапа",
            "stage_name": "Название этапа"
        }),
        width="stretch"
    )

    st.bar_chart(
        stage_summary.set_index("stage_name")["Прогресс, %"]
    )
else:
    st.info("Нет данных по этапам для отображения.")

st.divider()

# ===== ПЛАН VS ФАКТ =====
st.subheader("📈 План vs факт")

plan_fact = filtered_tasks[filtered_tasks["fact_date"].notna()].copy()

if not plan_fact.empty:
    plan_fact["В срок"] = (
        plan_fact["fact_date"] <= plan_fact["deadline"]
    ).astype(int)

    plan_fact["С задержкой"] = (
        plan_fact["fact_date"] > plan_fact["deadline"]
    ).astype(int)

    plan_fact_summary = (
        plan_fact
        .groupby(["stage_num", "stage_name"], dropna=False)
        .agg(
            **{
                "В срок": ("В срок", "sum"),
                "С задержкой": ("С задержкой", "sum")
            }
        )
        .reset_index()
    )

    st.dataframe(
        plan_fact_summary.rename(columns={
            "stage_num": "Этап",
            "stage_name": "Название"
        }),
        width="stretch"
    )

    if not plan_fact_summary.empty:
        st.bar_chart(
            plan_fact_summary.set_index("stage_name")[["В срок", "С задержкой"]]
        )
else:
    st.info("Нет данных для блока План vs факт.")

st.caption(f"Источник: {sheet_url}")
st.caption(f"Последнее обновление страницы: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}")