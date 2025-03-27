import streamlit as st
st.set_page_config(page_title="Дашборд воронки бота", layout="wide")

import pandas as pd

# Путь к Excel-файлу, содержащему несколько листов
file_path = "Специалист ТБ - 2 Оператор ТБ - Входящий ВА (7).xlsx"

###############################################################################
# ФУНКЦИИ ДЛЯ ВЫЧИСЛЕНИЙ И ОТРИСОВКИ
###############################################################################

def get_bar_width(value, max_val, base_width=80, min_width=5):
    """Возвращает ширину (в %) для полосы, пропорциональную value относительно max_val."""
    if max_val <= 0:
        return f"{min_width}%"
    ratio = value / max_val
    width = ratio * base_width
    if width < min_width:
        width = min_width
    return f"{width}%"

def draw_bar(label, count, color, max_val, font_size="20px", margin="10px auto"):
    """Отрисовка одиночной полосы (для активаций и топ-ответов кросс-продаж)."""
    bar_width = get_bar_width(count, max_val)
    bar_html = f"""
    <div style="
         background-color: {color};
         padding: 10px;
         margin: {margin};
         width: {bar_width};
         border-radius: 5px;
         text-align: center;
         color: white;">
        <span style="font-size: {font_size};"><b>{label}: {count}</b></span>
    </div>
    """
    st.markdown(bar_html, unsafe_allow_html=True)

def short_label(column_name, max_len=50):
    """Сокращает очень длинный заголовок столбца."""
    s = " ".join(column_name.splitlines()).strip()
    return s if len(s) <= max_len else s[:max_len] + "..."

def draw_double_bar(main_label, get_count, other_count, max_val, color="purple"):
    """Отрисовка двойной полосы (для этапа деталей кросс-продаж: «получить» и «другое»)."""
    get_width = get_bar_width(get_count, max_val)
    other_width = get_bar_width(other_count, max_val)
    block_html = f"""
    <div style="margin-bottom: 20px;">
      <div style="font-size:20px; font-weight:bold; text-align:center; margin-bottom: 10px;">{main_label}</div>
      <div style="display: flex; justify-content: center; gap: 20px;">
        <div style="background-color: {color}; color: white; padding: 10px; border-radius: 5px; text-align:center; width: {get_width};">
            {get_count}
        </div>
        <div style="background-color: {color}; color: white; padding: 10px; border-radius: 5px; text-align:center; width: {other_width};">
            {other_count}
        </div>
      </div>
      <div style="display: flex; justify-content: center; gap: 20px; margin-top: 5px;">
          <div style="width: {get_width}; text-align: center; font-size:18px;">Получить</div>
          <div style="width: {other_width}; text-align: center; font-size:18px;">Другое</div>
      </div>
    </div>
    """
    st.markdown(block_html, unsafe_allow_html=True)

def draw_stage_row(stage_name, responses, max_val, color="green"):
    """
    Отрисовка этапа «Описание вакансий»: слева надпись "Описание вакансий:",
    справа несколько полос (только число внутри) с подписями под ними.
    responses: список кортежей [(label, count), ...].
    """
    stage_html = f"""
    <div style="display: flex; align-items: center; margin-bottom: 20px;">
      <div style="width: 200px; text-align: right; font-weight: bold; margin-right: 20px;">
         {stage_name}:
      </div>
      <div style="display: flex; gap: 20px; align-items: center;">
    """
    for label, count in responses:
        bar_width = get_bar_width(count, max_val)
        stage_html += f"""
         <div style="display: flex; flex-direction: column; align-items: center;">
           <div style="background-color: {color}; width: {bar_width}; padding: 10px; border-radius: 5px; text-align: center; color: white;">
             {count}
           </div>
           <div style="margin-top: 5px; font-size: 18px;">{label}</div>
         </div>
        """
    stage_html += "</div></div>"
    st.markdown(stage_html, unsafe_allow_html=True)

def process_sheet(df_raw):
    """
    Обрабатывает один лист (df_raw) и возвращает словарь с:
    - названием бота,
    - рассчитанными метриками,
    - всеми необходимыми значениями для построения дашборда.
    """
    # Если строк меньше 3, пропускаем
    if len(df_raw) < 3:
        return None

    # Извлекаем название бота
    bot_name = df_raw.iloc[0, 0]

    # Заголовки столбцов – вторая строка
    headers = df_raw.iloc[1]

    # Данные – начиная с третьей строки
    df = pd.DataFrame(df_raw.iloc[2:].values, columns=headers)

    # --- 1. Этап: Активаций ---
    candidate_activation_cols = []
    for col in df.columns:
        if isinstance(col, str):
            if "HeadHunter" in col:
                candidate_activation_cols.append(col)
            elif "Добро пожаловать к новым возможностям" in col:
                candidate_activation_cols.append(col)

    activation_col = None
    activation_count = 0
    if candidate_activation_cols:
        counts = {col: df[col].notna().sum() for col in candidate_activation_cols}
        activation_col = max(counts, key=counts.get)
        activation_count = counts[activation_col]

    # --- 2. Этап: Описание вакансий ---
    job_desc_candidate = None
    max_keyword_count = 0
    for col in df.columns:
        if col in candidate_activation_cols:
            continue
        series = df[col].astype(str)
        keyword_count = series.str.contains("Что еще", case=False, na=False).sum() + \
                        series.str.contains("Мне нравится", case=False, na=False).sum()
        if keyword_count > max_keyword_count:
            max_keyword_count = keyword_count
            job_desc_candidate = col

    if job_desc_candidate:
        likes_count = df[job_desc_candidate].astype(str).str.contains("Мне нравится", case=False, na=False).sum()
        what_else_count = df[job_desc_candidate].astype(str).str.contains("Что еще", case=False, na=False).sum()
        not_answered_count = activation_count - (likes_count + what_else_count)
    else:
        likes_count = what_else_count = not_answered_count = 0

    # --- 3. Этап: Кросс‑продажи ---
    cross_selling_title = """/#*Пока ждем собеседование или думаете над заполнением анкеты, сделайте важный шаг!*

🎉 Мы подготовили для вас 2* бонуса*, которые, надеемся, будут вам полезны. 

🚀 *Уверены, что вы скажите нам за это отдельное спасибо!*
Выберите! #/"""

    cross_selling_col = None
    for col in df.columns:
        if isinstance(col, str) and col.strip() == cross_selling_title.strip():
            cross_selling_col = col
            break

    if cross_selling_col:
        cs_counts = df[cross_selling_col].value_counts(dropna=True)
        top3 = cs_counts.head(3)
        cs_no_response = (likes_count + what_else_count) - top3.sum()
    else:
        top3 = None
        cs_no_response = 0

    # --- 4. Этап: Детали кросс‑продаж ---
    cs_detail_candidates = {}
    for col in df.columns:
        if col in candidate_activation_cols or col == job_desc_candidate or col == cross_selling_col:
            continue
        series_lower = df[col].astype(str).str.lower()
        if "получить" in series_lower.values and "другое" in series_lower.values:
            cnt = (series_lower == "получить").sum() + (series_lower == "другое").sum()
            cs_detail_candidates[col] = cnt

    cs_detail_cols = sorted(cs_detail_candidates, key=lambda x: cs_detail_candidates[x], reverse=True)[:2]
    cs_detail_counts = {}
    for col in cs_detail_cols:
        series_lower = df[col].astype(str).str.lower().str.strip()
        cs_detail_counts[col] = {
            "получить": int((series_lower == "получить").sum()),
            "другое": int((series_lower == "другое").sum())
        }

    # --- Собираем все значения для поиска max ---
    all_values = [
        activation_count,
        likes_count,
        what_else_count,
        not_answered_count,
        cs_no_response
    ]
    if top3 is not None:
        all_values.extend(top3.values)
    for col in cs_detail_cols:
        detail = cs_detail_counts[col]
        all_values.append(detail["получить"])
        all_values.append(detail["другое"])

    global_max = max(all_values) if all_values else 1

    return {
        "bot_name": bot_name,
        "activation_count": activation_count,
        "likes_count": likes_count,
        "what_else_count": what_else_count,
        "not_answered_count": not_answered_count,
        "top3": top3,
        "cs_no_response": cs_no_response,
        "cs_detail_cols": cs_detail_cols,
        "cs_detail_counts": cs_detail_counts,
        "global_max": global_max
    }

###############################################################################
# ОСНОВНОЙ КОД: СЧИТЫВАЕМ ВСЕ ЛИСТЫ, СОХРАНЯЕМ РЕЗУЛЬТАТЫ, СТРОИМ ДАШБОРД
###############################################################################

# Считываем все листы Excel
xl = pd.ExcelFile(file_path)
bots_data = []

for sheet_name in xl.sheet_names:
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    metrics = process_sheet(df_raw)
    if metrics is not None:
        bots_data.append(metrics)

# Если ни одного листа не удалось обработать, завершаемся
if not bots_data:
    st.write("Не найдено ни одного корректного листа в файле.")
    st.stop()

# В боковой панели выбираем бота (по названию из A1)
bot_names = [d["bot_name"] for d in bots_data]
selected_bot_name = st.sidebar.selectbox("Выберите бота:", bot_names)

# Находим данные выбранного бота
selected_data = next(d for d in bots_data if d["bot_name"] == selected_bot_name)

# Извлекаем все метрики для удобства
bot_name = selected_data["bot_name"]
activation_count = selected_data["activation_count"]
likes_count = selected_data["likes_count"]
what_else_count = selected_data["what_else_count"]
not_answered_count = selected_data["not_answered_count"]
top3 = selected_data["top3"]
cs_no_response = selected_data["cs_no_response"]
cs_detail_cols = selected_data["cs_detail_cols"]
cs_detail_counts = selected_data["cs_detail_counts"]
global_max = selected_data["global_max"]

# --- Построение дашборда ---

# Заголовок – название бота по центру
st.markdown(f"<h1 style='text-align: center;'><b>{bot_name}</b></h1>", unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# --- Этап 1: Активаций (синяя полоса) ---
draw_bar("Активаций", activation_count, "blue", global_max)


# --- Этап 2: Описание вакансий ---
st.markdown("<h3 style='text-align: center;'><b>Описание вакансий</b></h1>", unsafe_allow_html=True)
draw_bar("Мне нравится", likes_count, "green", global_max)
draw_bar("Что еще", what_else_count, "green", global_max, font_size="18px", margin="0px auto 20px auto")
draw_bar("Не ответили", not_answered_count, "green", global_max, font_size="18px", margin="0px auto 20px auto")


# --- Этап 3: Кросс‑продажи (оранжевая палитра) ---
st.markdown("<h3 style='text-align: center;'><b>Кросс‑продажи</b></h3>", unsafe_allow_html=True)
if top3 is not None:
    for response, count in top3.items():
        draw_bar(response, count, "orange", global_max)
draw_bar("Не ответили", cs_no_response, "orange", global_max, font_size="18px", margin="0px auto 20px auto")

# --- Этап 4: Детали кросс‑продаж ---
if cs_detail_cols:
    for idx, col in enumerate(cs_detail_cols, start=1):
        detail = cs_detail_counts[col]
        col_short = short_label(col)
        main_label = f"Кросс продажа {idx} ({col_short})"
        draw_double_bar(main_label, detail["получить"], detail["другое"], global_max, color="purple")
