import streamlit as st
from datetime import datetime, timedelta
from calendar import isleap
from docx import Document
from io import BytesIO
import pandas as pd

st.set_page_config(page_title="Генератор спецификации", layout="wide")
st.title("📄 Генератор спецификации по программам")

PROGRAM_OPTIONS = ["С1", "КБ", "КЛ"]

# Храним строки в session_state
if "rows" not in st.session_state:
    st.session_state.rows = []

# ➕ Добавить строку
if st.button("➕ Добавить строку"):
    st.session_state.rows.append({
        "name": PROGRAM_OPTIONS[0],
        "start_date": datetime.today().date(),
        "end_date": datetime.today().date(),
        "count": 1,
        "price_annual": 0.0
    })

# Форма ввода
valid_rows = []
for i, row in enumerate(st.session_state.rows):
    cols = st.columns([1.2, 1, 1, 1, 1])
    with cols[0]:
        row["name"] = st.selectbox(f"Программа {i+1}", PROGRAM_OPTIONS, key=f"name_{i}")
    with cols[1]:
        row["start_date"] = st.date_input(f"Начало {i+1}", value=row["start_date"], format="DD.MM.YYYY", key=f"start_{i}")
    with cols[2]:
        row["end_date"] = st.date_input(f"Окончание {i+1}", value=row["end_date"], format="DD.MM.YYYY", key=f"end_{i}")
    with cols[3]:
        row["count"] = st.number_input(f"Кол-во {i+1}", min_value=1, step=1, value=row["count"], key=f"count_{i}")
    with cols[4]:
        row["price_annual"] = st.number_input(f"₽ за 12 мес {i+1}", min_value=0.0, step=100.0, value=row["price_annual"], key=f"price_{i}")

    if row["start_date"] <= row["end_date"] and row["price_annual"] > 0:
        valid_rows.append(row)

# 💰 Расчёт по дням
def calculate_price(start_date, end_date, annual_price):
    total = 0.0
    current = start_date
    while current <= end_date:
        year_days = 366 if isleap(current.year) else 365
        total += annual_price / year_days
        current += timedelta(days=1)
    return round(total, 2)

# 📄 Генерация спецификации
if valid_rows and st.button("📄 Сгенерировать спецификацию"):
    doc = Document()
    doc.add_heading("Спецификация", level=1)

    # Заголовок таблицы Word
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text = "№"
    hdr[1].text = "Наименование программы для ЭВМ"
    hdr[2].text = "Кол-во лицензий"
    hdr[3].text = "Срок, на который предоставляется право"
    hdr[4].text = "Стоимость лицензии, руб. РФ"
    hdr[5].text = "Сумма, руб. РФ"

    st.markdown("### 🧾 Расчёт по позициям:")

    result_data = []
    for idx, p in enumerate(valid_rows, 1):
        start_dt = datetime.combine(p["start_date"], datetime.min.time())
        end_dt = datetime.combine(p["end_date"], datetime.min.time())
        per_license = calculate_price(start_dt, end_dt, p["price_annual"])
        total_price = round(per_license * p["count"], 2)

        start_str = p["start_date"].strftime('%d.%m.%Y')
        end_str = p["end_date"].strftime('%d.%m.%Y')
        period_str = f"от {start_str} до {end_str} гг."

        # Word
        row = table.add_row().cells
        row[0].text = str(idx)
        row[1].text = f"Программа для ЭВМ {p['name']}"
        row[2].text = str(p["count"])
        row[3].text = period_str
        row[4].text = f"{per_license:.2f}"
        row[5].text = f"{total_price:.2f}"

        # Интерфейс
        result_data.append({
            "№": idx,
            "Наименование программы для ЭВМ": f"Программа для ЭВМ {p['name']}",
            "Кол-во лицензий": p["count"],
            "Срок": period_str,
            "Стоимость лицензии, руб. РФ": f"{per_license:.2f}",
            "Сумма, руб. РФ": f"{total_price:.2f}"
        })

    df = pd.DataFrame(result_data)
    st.table(df)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="📥 Скачать спецификацию (.docx)",
        data=buffer,
        file_name="спецификация.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# Очистка
if st.button("🗑️ Очистить всё"):
    st.session_state.rows = []
    st.success("Все строки удалены.")
