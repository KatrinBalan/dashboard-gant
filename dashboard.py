import os
from datetime import datetime

import gspread
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Гант Гарант 2026", layout="wide")

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1DlgMbUkXySIBtQIT8n0k0hdN5qJeYxBFpc-R8a9FYm4/edit?usp=sharing"
SHEET_NAME = "Диаграмма Ганта Гарнт 2026"

st.markdown("""
<style>
.card {
    padding: 24px;
    border-radius: 18px;
    color: #111;
    font-size: 30px;
    font-weight: 600;
    box-shadow: 0 4px 14px rgba(0,0,0,0.08);
}
.card-number {
    font-size: 48px;
    font-weight: 800;
    margin-top: 10px;
}
.red { background: #ffd6d6; }
.yellow { background: #fff1b8; }
.green { background: #d8f5d0; }
.gray { background: #eeeeee; }
</style>
""", unsafe_allow_html=True)


def read_google_sheet(sheet_url):
    if os.path.exists("credentials.json"):
        gc = gspread.service_account(filename="credentials.json")
    else:
        service_account_info = dict(st.secrets["gcp_service_account"])
        service_account_info["private_key"] = service_account_info["private_key"].replace("\\n", "\n")
        gc = gspread.service_account_from_dict(service_account_info)

    sh = gc.open_by_url(sheet_url)
    ws = sh.worksheet(SHEET_NAME)

    # ВАЖНО: UNFORMATTED_VALUE сохраняет числа как числа
    data = ws.get("A1:O", value_render_option="UNFORMATTED_VALUE")
    return pd.DataFrame(data)


def parse_date(value):
    if pd.isna(value) or str(value).strip() == "":
        return pd.NaT
    return pd.to_datetime(value, dayfirst=True, errors="coerce")


def parse_percent(value):
    if pd.isna(value) or str(value).strip() == "":
        return pd.NA

    value = str(value).replace("%", "").replace(",", ".").strip()

    try:
        number = float(value)
    except ValueError:
        return pd.NA

    return number / 100 if number > 1 else number


def is_number(value):
    try:
        float(value)
        return True
    except Exception:
        return False


def prepare_data(df):
    raw = df.copy().reset_index(drop=True)

    # На всякий случай расширяем таблицу до 15 колонок, если Google вернул меньше
    while raw.shape[1] < 15:
        raw[raw.shape[1]] = ""

    gantt = pd.DataFrame()

    # A = 0, B = 1, C = 2, H = 7, I = 8, L = 11, M = 12
    gantt["A"] = raw.iloc[:, 0]
    gantt["B"] = raw.iloc[:, 1]
    gantt["C_function"] = raw.iloc[:, 2]
    gantt["H_deadline"] = raw.iloc[:, 7]
    gantt["I_fact"] = raw.iloc[:, 8]
    gantt["L_status"] = raw.iloc[:, 11]
    gantt["M_progress"] = raw.iloc[:, 12]

    gantt["row_number"] = gantt.index + 1

    gantt["A_text"] = gantt["A"].astype(str).str.strip()
    gantt["B_text"] = gantt["B"].astype(str).str.strip()
    gantt["function"] = gantt["C_function"].astype(str).str.strip()

    # Убираем служебные пустые значения
    for col in ["A_text", "B_text", "function"]:
        gantt[col] = gantt[col].replace(["nan", "None", "<NA>"], "")

    # Строки до 6-й строки не учитываем, как в формуле A6:A
    gantt = gantt[gantt["row_number"] >= 6].copy()

    # Этап = в A целое число: 0, 1, 2, 3...
    gantt["A_number"] = pd.to_numeric(
        gantt["A_text"].str.replace(",", ".", regex=False),
        errors="coerce"
    )

    gantt["is_stage"] = (
        gantt["A_number"].notna()
        & (gantt["A_number"] % 1 == 0)
        & (gantt["B_text"] != "")
    )

    # Задача = A не пустая и не число
    gantt["is_task"] = (
        (gantt["A_text"] != "")
        & ~gantt["is_stage"]
    )

    # ПРОСМОТР: протягиваем номер и название этапа вниз
    gantt["stage_num"] = pd.NA
    gantt["stage_name"] = pd.NA

    gantt.loc[gantt["is_stage"], "stage_num"] = gantt.loc[gantt["is_stage"], "A_number"].astype("Int64")
    gantt.loc[gantt["is_stage"], "stage_name"] = gantt.loc[gantt["is_stage"], "B_text"]

    gantt["stage_num"] = gantt["stage_num"].ffill()
    gantt["stage_name"] = gantt["stage_name"].ffill()

    gantt["task_name"] = gantt["B_text"]
    gantt["deadline"] = gantt["H_deadline"].apply(parse_date)
    gantt["fact_date"] = gantt["I_fact"].apply(parse_date)
    gantt["status"] = gantt["L_status"].astype(str).str.strip()
    gantt["progress"] = gantt["M_progress"].apply(parse_percent)

    return gantt


def calculate_kpi(gantt):
    today = pd.Timestamp.today().normalize()
    three_days = today + pd.Timedelta(days=3)

    tasks = gantt[gantt["is_task"]].copy()

    has_name = tasks["task_name"].astype(str).str.strip() != ""
    not_done = tasks["status"].str.lower() != "завершено"

    overdue = tasks[
        has_name
        & not_done
        & tasks["deadline"].notna()
        & (tasks["deadline"] < today)
    ].shape[0]

    risk = tasks[
        has_name
        & not_done
        & tasks["deadline"].notna()
        & (tasks["deadline"] >= today)
        & (tasks["deadline"] <= three_days)
    ].shape[0]

    completed = tasks[
        has_name
        & (tasks["status"].str.lower() == "завершено")
    ].shape[0]

    progress_values = tasks["progress"].dropna()
    project_progress = progress_values.mean() if len(progress_values) else 0

    return overdue, risk, completed, project_progress


# ===== ЗАГРУЗКА ДАННЫХ =====
st.title("📊 Гант Гарант 2026")
st.markdown("""
<style>

/* ===== ФИЛЬТРЫ (жёсткое переопределение) ===== */
div[data-baseweb="select"] span[data-baseweb="tag"] {
    background-color: #00E755 !important;
    color: #003138 !important;
    border: none !important;
}

/* текст внутри */
div[data-baseweb="select"] span[data-baseweb="tag"] span {
    color: #003138 !important;
}

/* крестик */
div[data-baseweb="select"] span[data-baseweb="tag"] svg {
    fill: #003138 !important;
}

/* hover */
div[data-baseweb="select"] span[data-baseweb="tag"]:hover {
    background-color: #00c94a !important;
}

</style>
""", unsafe_allow_html=True)
try:
    full_df = read_google_sheet(GOOGLE_SHEET_URL)
    gantt = prepare_data(full_df)

    overdue, risk, completed, project_progress = calculate_kpi(gantt)

except Exception as e:
    st.error("Не удалось загрузить данные из Google Sheets.")
    st.code(str(e))
    st.stop()


# ===== ПЕРЕКЛЮЧАТЕЛЬ =====
col_left, col_toggle, col_right = st.columns([1.1, 0.35, 2.5])

with col_left:
    st.markdown("### 🎛️ Дашборд")

with col_toggle:
    plan_fact_mode = st.toggle(
        "mode",
        value=False,
        label_visibility="collapsed"
    )

with col_right:
    st.markdown("### 📈 План VS ФАКТ")

page = "📈 План VS ФАКТ" if plan_fact_mode else "🎛️ Дашборд"


# ===== KPI =====
k1, k2, k3, k4 = st.columns(4)

with k1:
    st.markdown(
        f'<div class="card red">🔴 Просроченные<div class="card-number">{overdue}</div></div>',
        unsafe_allow_html=True
    )

with k2:
    st.markdown(
        f'<div class="card yellow">⚠️ Риск 3 дня<div class="card-number">{risk}</div></div>',
        unsafe_allow_html=True
    )

with k3:
    st.markdown(
        f'<div class="card green">✅ Завершено<div class="card-number">{completed}</div></div>',
        unsafe_allow_html=True
    )

with k4:
    st.markdown(
        f'<div class="card gray">📈 Прогресс проекта<div class="card-number">{project_progress:.0%}</div></div>',
        unsafe_allow_html=True
    )

st.divider()


# ===== ОБЩИЕ ДАННЫЕ =====
tasks = gantt[gantt["is_task"]].copy()

tasks = tasks[
    tasks["task_name"].notna()
    & (tasks["task_name"].astype(str).str.strip() != "")
]

all_stages = sorted(
    tasks[
        tasks["stage_name"].notna()
        & (tasks["stage_name"].astype(str).str.strip() != "")
        & (~tasks["stage_name"].astype(str).str.lower().isin(["nan", "<na>", "none"]))
    ]["stage_name"].unique()
)


if page == "🎛️ Дашборд":
    left_col, right_col = st.columns(2)

    with left_col:
        stages = sorted(tasks["stage_name"].dropna().unique())

        selected_stages = st.multiselect(
            "Фильтр по этапам",
            options=stages,
            default=stages
        )

        filtered_stage_tasks = tasks[tasks["stage_name"].isin(selected_stages)]

        st.subheader("🚀 Сводная по этапам")

        stage_summary = (
            filtered_stage_tasks
            .groupby(["stage_num", "stage_name"], dropna=False)
            .agg({
                "task_name": "count",
                "progress": "mean"
            })
            .reset_index()
        )

        stage_summary["progress"] = stage_summary["progress"].fillna(0)
        stage_summary["Прогресс, %"] = (
            stage_summary["progress"] * 100
        ).map(lambda x: f"{x:.2f}%")

        if not stage_summary.empty:
            st.table(
                stage_summary.rename(columns={
                    "stage_num": "Номер этапа",
                    "stage_name": "Название этапа",
                    "task_name": "Кол-во задач"
                })[
                    ["Номер этапа", "Название этапа", "Кол-во задач", "Прогресс, %"]
                ]
            )
        else:
            st.info("Нет данных по выбранным этапам.")

    with right_col:
        functions = sorted(tasks["function"].dropna().unique())

        selected_functions = st.multiselect(
            "Фильтр по функциям",
            options=functions,
            default=functions
        )

        filtered_function_tasks = tasks[tasks["function"].isin(selected_functions)]

        st.subheader("🧩 Сводная по функциям")

        function_summary = (
            filtered_function_tasks
            .groupby("function", dropna=False)
            .agg({
                "task_name": "count",
                "progress": "mean"
            })
            .reset_index()
        )

        function_summary["progress"] = function_summary["progress"].fillna(0)
        function_summary["Прогресс, %"] = (
            function_summary["progress"] * 100
        ).map(lambda x: f"{x:.2f}%")

        if not function_summary.empty:
            st.table(
                function_summary.rename(columns={
                    "function": "Функция",
                    "task_name": "Кол-во задач"
                })[
                    ["Функция", "Кол-во задач", "Прогресс, %"]
                ]
            )
        else:
            st.info("Нет данных по выбранным функциям.")


# ===== PLAN VS FACT =====
if page == "📈 План VS ФАКТ":
    st.subheader("📈 Plan vs Fact по этапам")

    selected_stages = st.multiselect(
        "Фильтр по этапам",
        options=all_stages,
        default=all_stages
    )

    filtered_tasks = tasks[tasks["stage_name"].isin(selected_stages)]

    plan_fact = filtered_tasks[filtered_tasks["fact_date"].notna()].copy()

    if not plan_fact.empty:
        both_dates = plan_fact["deadline"].notna() & plan_fact["fact_date"].notna()

        plan_fact["В срок"] = (
            both_dates
            & (plan_fact["fact_date"] <= plan_fact["deadline"])
        ).astype(int)

        plan_fact["С задержкой"] = (
            both_dates
            & (plan_fact["fact_date"] > plan_fact["deadline"])
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
            width="stretch",
            hide_index=True
        )

        st.bar_chart(
            plan_fact_summary.set_index("stage_name")[["В срок", "С задержкой"]]
        )
    else:
        st.info("Нет данных для блока Plan vs Fact по этапам.")


st.caption("Источник: Google Sheets")
st.caption(f"Последнее обновление: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}")