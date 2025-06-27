import streamlit as st
from datetime import datetime, timedelta
from calendar import isleap
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Генератор спецификации", layout="centered")
st.title("📄 Генератор спецификации по лицензиям")

# Ввод общей конечной даты
end_date_input = st.date_input("📅 Конечная дата действия всех лицензий (включительно)")
end_date = datetime.combine(end_date_input, datetime.min.time()) if end_date_input else None

# Хранилище лицензий
if "licenses" not in st.session_state:
    st.session_state.licenses = []

# ➕ Добавление лицензии
with st.form("add_license_form"):
    col1, col2 = st.columns(2)
    with col1:
        start_date_input = st.date_input("Дата начала лицензии", key="start_date")
    with col2:
        price_annual = st.number_input("Стоимость за 12 месяцев (₽)", min_value=0.0, step=100.0, key="price")

    submitted = st.form_submit_button("➕ Добавить лицензию")
    if submitted and start_date_input and price_annual and end_date:
        start_date = datetime.combine(start_date_input, datetime.min.time())
        if start_date > end_date:
            st.error("❌ Дата начала позже конечной!")
        else:
            st.session_state.licenses.append({
                "start_date": start_date,
                "price_annual": price_annual
            })
            st.success("✅ Лицензия добавлена!")

# Отображение добавленных лицензий
if st.session_state.licenses:
    st.markdown("### 📋 Добавленные лицензии:")
    for idx, lic in enumerate(st.session_state.licenses):
        st.markdown(
            f"**{idx + 1}.** С {lic['start_date'].strftime('%d.%m.%Y')} по {end_date.strftime('%d.%m.%Y')} — "
            f"{lic['price_annual']:.2f} ₽ / год"
        )

# 🔢 Функция с учётом високосных лет
def calculate_total_price(start_date, end_date, annual_price):
    total_price = 0.0
    current_date = start_date
    while current_date <= end_date:
        year_length = 366 if isleap(current_date.year) else 365
        daily_rate = annual_price / year_length
        total_price += daily_rate
        current_date += timedelta(days=1)
    return round(total_price, 2)

# 📄 Генерация спецификации
if st.button("📄 Сгенерировать спецификацию") and end_date and st.session_state.licenses:
    doc = Document()
    doc.add_heading("Спецификация", level=1)

    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    table.autofit = True
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "№"
    hdr_cells[1].text = "Период действия лицензии"
    hdr_cells[2].text = "Стоимость (₽)"

    for idx, lic in enumerate(st.session_state.licenses, 1):
        start = lic["start_date"]
        total_price = calculate_total_price(start, end_date, lic["price_annual"])

        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = f"с {start.strftime('%d.%m.%Y')} по {end_date.strftime('%d.%m.%Y')}"
        row_cells[2].text = f"{total_price:,.2f}".replace(",", " ").replace(".", ",")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="📥 Скачать спецификацию (.docx)",
        data=buffer,
        file_name="спецификация.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# 🔁 Очистить список
if st.button("🔁 Очистить лицензии"):
    st.session_state.licenses = []
    st.success("Список лицензий очищен.")
