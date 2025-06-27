import streamlit as st
from datetime import datetime, timedelta
from calendar import isleap
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Генератор спецификации", layout="wide")
st.title("📄 Генератор спецификации по программам")

PROGRAM_OPTIONS = ["С1", "КБ", "КЛ"]

# Хранилище строк лицензий
if "programs" not in st.session_state:
    st.session_state.programs = []

# Форма в одну строку
st.markdown("### ➕ Добавить позицию")

with st.form("add_row", clear_on_submit=True):
    cols = st.columns([1.2, 1, 1, 1, 1])
    with cols[0]:
        program_name = st.selectbox("Программа", PROGRAM_OPTIONS)
    with cols[1]:
        start_date = st.date_input("Начало", key="start")
    with cols[2]:
        end_date = st.date_input("Окончание", key="end")
    with cols[3]:
        license_count = st.number_input("Кол-во", min_value=1, step=1)
    with cols[4]:
        price_annual = st.number_input("₽ за 12 мес/1 лиц", min_value=0.0, step=100.0)

    submit = st.form_submit_button("Добавить")

    if submit:
        if start_date > end_date:
            st.error("❌ Дата начала позже даты окончания!")
        else:
            st.session_state.programs.append({
                "name": program_name,
                "start_date": datetime.combine(start_date, datetime.min.time()),
                "end_date": datetime.combine(end_date, datetime.min.time()),
                "count": license_count,
                "price_annual": price_annual
            })
            st.success("✅ Позиция добавлена!")

# Показываем текущие позиции
if st.session_state.programs:
    st.markdown("### 📋 Добавленные позиции:")
    for idx, p in enumerate(st.session_state.programs):
        st.markdown(
            f"**{idx + 1}.** {p['name']}: с {p['start_date'].strftime('%d.%m.%Y')} по {p['end_date'].strftime('%d.%m.%Y')}, "
            f"{p['count']} шт. по {p['price_annual']:.2f} ₽"
        )

# Функция расчёта стоимости
def calculate_price(start_date, end_date, annual_price):
    total = 0.0
    current = start_date
    while current <= end_date:
        year_days = 366 if isleap(current.year) else 365
        total += annual_price / year_days
        current += timedelta(days=1)
    return round(total, 2)

# Генерация спецификации
if st.button("📄 Сгенерировать спецификацию"):
    if not st.session_state.programs:
        st.warning("Добавьте хотя бы одну позицию.")
    else:
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

        for idx, p in enumerate(st.session_state.programs, 1):
            per_license = calculate_price(p["start_date"], p["end_date"], p["price_annual"])
            total_price = round(per_license * p["count"], 2)

            row = table.add_row().cells
            row[0].text = str(idx)
            row[1].text = f"Программа для ЭВМ «{p['name']}»"
            row[2].text = f"с {p['start_date'].strftime('%d.%m.%Y')} по {p['end_date'].strftime('%d.%m.%Y')}"
            row[3].text = f"{per_license:,.2f}".replace(",", " ").replace(".", ",") + " ₽"
            row[4].text = str(p["count"])
            row[5].text = f"{total_price:,.2f}".replace(",", " ").replace(".", ",") + " ₽"

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="📥 Скачать спецификацию (.docx)",
            data=buffer,
            file_name="спецификация.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# Кнопка очистки
if st.button("🗑️ Очистить всё"):
    st.session_state.programs = []
    st.success("Все позиции удалены.")
