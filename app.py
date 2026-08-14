import streamlit as st
import pandas as pd
import datetime
import calendar

st.set_page_config(page_title="Генератор звітів", layout="wide")

st.title("Генератор звітів")

# Configuration in sidebar
st.sidebar.header("Налаштування планів")
plan_instagram = st.sidebar.number_input("План Instagram", value=4500000, step=100000)
plan_site = st.sidebar.number_input("План сайт", value=1500000, step=100000)
zalyzhok = st.sidebar.number_input("Залишок за попередній тиждень", value=500000, step=10000)
zakryto_tyzhden = st.sidebar.number_input("Закрито за тиждень", value=786300, step=10000)

uploaded_file = st.file_uploader("Завантажте Excel файл (.xlsx)", type=["xlsx"])

def format_space(num):
    return f"{int(num):,}".replace(",", " ")

def format_dot(num):
    return f"{int(num):,}".replace(",", ".")

if uploaded_file is not None:
    xl = pd.ExcelFile(uploaded_file)
    sheet_name = xl.sheet_names[0] # taking the latest sheet (first one)
    df = xl.parse(sheet_name)
    
    st.write(f"**Парситься сторінка:** {sheet_name}")
    
    # Identify manager blocks
    manager_blocks = []
    for idx in range(len(df.columns)):
        val = str(df.iloc[0, idx]).strip()
        if val.startswith("№ замовлення"):
            start_col = idx - 1
            if start_col >= 0 and start_col + 2 < len(df.columns):
                # Try to get manager name from second row (index 1)
                manager_name = str(df.iloc[1, start_col]).strip()
                if manager_name == "nan" or manager_name == "":
                    manager_name = "Стажери"
                manager_blocks.append({
                    "name": manager_name,
                    "start_col": start_col
                })
    
    managers_data = {}
    
    st.subheader("Плани менеджерів")
    manager_plans = {}
    default_plans = {
        "Вікторія": 1200000,
        "Тетяна": 1000000,
        "Катерина": 1800000,
        "Анна": 1500000,
        "Стажери": 500000
    }
    
    # Layout for manager plans
    cols = st.columns(min(len(manager_blocks), 5) or 1)
    
    for i, block in enumerate(manager_blocks):
        m_name = block["name"]
        def_plan = default_plans.get(m_name, 1000000)
        with cols[i % len(cols)]:
            manager_plans[m_name] = st.number_input(f"План: {m_name}", value=def_plan, step=100000)
        
    total_manager_plans = sum(manager_plans.values())
    total_plan_month = plan_instagram + plan_site
    
    if total_manager_plans > total_plan_month:
        st.warning(f"Увага: Загальна сума планів менеджерів ({format_space(total_manager_plans)}) перевищує загальний план на місяць ({format_space(total_plan_month)})!")
        
    if st.button("Згенерувати звіт", type="primary"):
        total_zakryto = 0
        zakryto_instagram = 0
        zakryto_site = 0
        
        # Process each block
        for block in manager_blocks:
            m_name = block["name"]
            start_col = block["start_col"]
            
            # The columns in block: 0: п/н, 1: № замовлення, 2: сума, 3: сайт/інстаграм (optional)
            suma_col = df.columns[start_col + 2]
            
            has_source_col = (start_col + 3) < len(df.columns)
            
            # Skip the header rows (index 0 and 1)
            df_manager = df.iloc[2:].copy()
            
            df_manager[suma_col] = pd.to_numeric(df_manager[suma_col], errors='coerce')
            df_manager = df_manager.dropna(subset=[suma_col])
            
            m_zakryto = df_manager[suma_col].sum()
            total_zakryto += m_zakryto
            
            # Check for Site ("сайт")
            m_site_sum = 0
            if has_source_col:
                source_col = df.columns[start_col + 3]
                site_mask = df_manager[source_col].notna() & (df_manager[source_col].astype(str).str.lower().str.contains("сайт"))
                m_site_sum = df_manager[site_mask][suma_col].sum()
                
            m_insta_sum = m_zakryto - m_site_sum
            
            zakryto_site += m_site_sum
            zakryto_instagram += m_insta_sum
            
            managers_data[m_name] = {
                "zakryto": m_zakryto
            }
            
        now = datetime.datetime.now()
        current_date_str = now.strftime("%d.%m.%Y")
        days_in_month = calendar.monthrange(now.year, now.month)[1]
        current_day = now.day
        
        plan_day = ((total_plan_month / 4) + zalyzhok) / 7
        
        month_progress = (total_plan_month / days_in_month) * current_day
        total_diff = total_zakryto - month_progress
        total_diff_sign = "+" if total_diff > 0 else ""
        
        plan_week = total_plan_month / 4
        zakryto_tyzhden_percent = (zakryto_tyzhden / plan_week) * 100 if plan_week > 0 else 0
        
        # Build text
        report = []
        report.append("Добрий вечір ✨")
        report.append(f"Станом на {current_date_str} р:")
        report.append("")
        report.append(f"План на день: {format_space(plan_day)} грн")
        report.append(f"Закрито: {format_space(total_zakryto)} грн")
        report.append(f"{total_diff_sign}{format_space(total_diff)} грн")
        report.append("")
        report.append(f"План на тиждень: {format_dot(plan_week)} грн")
        report.append(f"Закрито: {format_space(zakryto_tyzhden)} грн / {int(zakryto_tyzhden_percent)}%")
        report.append("")
        report.append(f"План Instagram: {format_dot(plan_instagram)} грн")
        report.append(f"Закрито: {format_space(zakryto_instagram)} грн")
        report.append(f"План сайт: {format_dot(plan_site)} грн")
        report.append(f"Закрито: {format_space(zakryto_site)} грн")
        report.append("")
        
        for m_name, data in managers_data.items():
            m_plan = manager_plans[m_name]
            m_zakryto = data["zakryto"]
            m_progress = (m_plan / days_in_month) * current_day
            m_diff = m_zakryto - m_progress
            m_diff_sign = "+" if m_diff > 0 else ""
            
            # Format: name: plan if name has colon? The user example shows "Вікторія 1.200.000 грн" or "Катерина: 1.800.000 грн"
            # We'll just use "Name: Plan" for consistency or "Name Plan"
            report.append(f"{m_name}: {format_dot(m_plan)} грн")
            report.append(f"Закрито: {format_space(m_zakryto)} грн")
            report.append(f"{m_diff_sign}{format_space(m_diff)} грн")
            report.append("")
            
        report_text = "\n".join(report)
        
        st.subheader("Згенерований звіт")
        st.code(report_text, language="text")
        st.success("Звіт успішно згенеровано! Використовуйте кнопку копіювання у правому верхньому куті блоку тексту.")
