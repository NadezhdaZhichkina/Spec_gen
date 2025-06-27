import streamlit as st
from datetime import datetime, timedelta
from calendar import isleap
from docx import Document
from io import BytesIO

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

# Форма строк (редактируемых)
valid_rows = []
for i, row in enumerate(st.session_state.rows):
    cols = st.columns([1.2, 1, 1, 1, 1])
    with cols[0]:
        row["name"] = st.selectbox(f"Программа {i+1}", PROGRAM_OPTIONS, key=f"name_{i}")
    with cols[1]:
        row["start_date"] = st.date_input(f"Начало {i+1}", value=row["start_date"], key=f"start_{i}")
    with cols[2]:
        row["end_date"] = st.date_input(f"Окончание {i+1}", value=row["end_date"], key=f"end_{i}")
    with cols[3]:
        row["count"] = st.number_input(f"Кол-во {i+1}", min_value=1, step=1, value=row["count"], key=f"count_{i}")
    with cols[4]:
        row["price_annual"] = st.number_input(f"₽ за 12 мес {i+1}", min_value=0.0, step=100.0, value=row["price_annual"], key=f"price_{i}")

    # Валидация
    if row["start_date"] <= row["end_date"] and row["price_annual"] > 0:
        valid_rows.append(row)

# 💰 Калькулятор с учётом високосных лет
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

    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text = "№"
    hdr[1].text = "Программа"
    hdr[2].text = "Срок действия"
    hdr[3].text = "Стоимость 1 лицензии"
    hdr[4].text = "Кол-во"
    hdr[5].text = "Стоимость всего"

    st.markdown("### 🧾 Расчёт по позициям:")

    for idx, p in enumerate(valid_rows, 1):
        start_dt = datetime.combine(p["start_date"], datetime.min.time())
        end_dt = datetime.combine(p["end_date"], datetime.min.time())
        per_license = calculate_price(start_dt, end_dt, p["price_annual"])
        total_price = round(per_license * p["count"], 2)

        # Форматируем даты
        start_str = p["start_date"].strftime('%d.%m.%Y')
        end_str = p["end_date"].strftime('%d.%m.%Y')

        # Записываем в Word таблицу
        row = table.add_row().cells
        row[0].text = str(idx)
        row[1].text = f"Программа для ЭВМ {p['name']}"
        row[2].text = f"с {start_str} по {end_str}"
        row[3].text = f"{per_license:,.2f}".replace(",", " ").replace(".", ",") + " ₽"
        row[4].text = str(p["count"])
        row[5].text = f"{total_price:,.2f}".replace(",", " ").replace(".", ",") + " ₽"

        # Показываем пользователю
        st.markdown(
            f"**{idx}.** Программа для ЭВМ {p['name']} с {start_str} по {end_str}, "
            f"{p['count']} шт. — {per_license:,.2f} ₽ за 1, **{total_price:,.2f} ₽ всего**"
            .replace(",", " ").replace(".", ",")
        )

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
