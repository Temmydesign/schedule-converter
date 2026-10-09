import streamlit as st
import pandas as pd
import numpy as np
import datetime

# --- Page Configuration ---
st.set_page_config(
    page_title="AI Schedule Converter",
    page_icon="📅",
    layout="centered"
)

st.title("📅 Integrated L2 Schedule Converter")
st.write("Upload your Excel Schedule Matrix to instantly generate a policy-compliant CSV for **Microsoft Project** and **Primavera P6**.")

# --- Processing Function ---
def convert_schedule(uploaded_file, project_title, start_date):
    xls = pd.ExcelFile(uploaded_file)
    
    # Auto-detect schedule sheet name
    matching_sheets = [s for s in xls.sheet_names if 'L2 Schedule' in s or 'Schedule' in s]
    selected_sheet = matching_sheets[0] if matching_sheets else xls.sheet_names[0]
    
    df_raw = pd.read_excel(xls, sheet_name=selected_sheet, skiprows=2)
    
    # Standardize header
    df_raw.columns = df_raw.iloc[0]
    df_data = df_raw.iloc[1:].dropna(subset=['Activity'])
    
    # Check if 'System' column exists, otherwise assign default 'TRP'
    if 'System' not in df_data.columns:
        df_data['System'] = 'TRP'
        
    tasks_list = []
    
    # 1. Top-Level Project Title Header (Outline Level 1)
    tasks_list.append({
        'Task Name': project_title,
        'Duration': '',
        'Start Date': start_date.strftime("%Y-%m-%d"),
        'Finish Date': '2027-04-30',
        'Predecessors': '',
        'Outline Level': 1
    })
    
    system_groups = df_data.groupby('System', sort=False)
    id_map = {}
    temp_id = 2
    
    for sys_name, group in system_groups:
        # System Summary Row (Level 2)
        sys_sum_old_id = f"SUM_{sys_name}"
        id_map[sys_sum_old_id] = temp_id
        
        tasks_list.append({
            'Task Name': f"System: {sys_name}",
            'Duration': '',
            'Start Date': str(group['Start date'].min()).split(' ')[0],
            'Finish Date': str(group['Finish date'].max()).split(' ')[0],
            'Predecessors': '',
            'Outline Level': 2
        })
        temp_id += 1
        
        sub_tasks = []
        for idx, r in group.iterrows():
            dur_val = r['Dur (wks)'] if 'Dur (wks)' in r else r.get('Dur_Wks', 0)
            dur_str = f"{int(dur_val)} wks" if pd.notna(dur_val) and int(dur_val) > 0 else "0 wks"
            
            sub_tasks.append({
                'Task Name': r['Activity'],
                'Duration': dur_str,
                'Start Date': str(r['Start date']).split(' ')[0],
                'Finish Date': str(r['Finish date']).split(' ')[0],
                'Outline Level': 3
            })
            
        latest_act_id = None
        latest_finish = None
        
        for i, task in enumerate(sub_tasks):
            task_old_id = f"ACT_{sys_name}_{i+1}"
            id_map[task_old_id] = temp_id
            
            f_dt = pd.to_datetime(task['Finish Date'])
            if latest_finish is None or f_dt >= latest_finish:
                latest_finish = f_dt
                latest_act_id = temp_id
                
            preds = []
            if i == 0:
                preds.append("3SS")
                
            tasks_list.append({
                'Task Name': task['Task Name'],
                'Duration': task['Duration'],
                'Start Date': task['Start Date'],
                'Finish Date': task['Finish Date'],
                'Predecessors': ", ".join(preds),
                'Outline Level': 3
            })
            temp_id += 1
            
        # Completion Milestone
        m_name = f"{sys_name} completed"
        m_pred = f"{latest_act_id}FS"
        
        tasks_list.append({
            'Task Name': m_name,
            'Duration': '0 wks',
            'Start Date': str(latest_finish).split(' ')[0],
            'Finish Date': str(latest_finish).split(' ')[0],
            'Predecessors': m_pred,
            'Outline Level': 3
        })
        temp_id += 1

    df_out = pd.DataFrame(tasks_list)
    df_out['ID'] = range(1, len(df_out) + 1)
    
    # Generate WBS
    wbs = ['1']
    s_cnt = 0
    a_cnt = 0
    for idx in range(1, len(df_out)):
        r = df_out.iloc[idx]
        if r['Outline Level'] == 2:
            s_cnt += 1
            a_cnt = 0
            wbs.append(f"1.{s_cnt}")
        else:
            a_cnt += 1
            wbs.append(f"1.{s_cnt}.{a_cnt}")
            
    df_out['WBS'] = wbs
    cols = ['ID', 'Task Name', 'Duration', 'Start Date', 'Finish Date', 'Predecessors', 'Outline Level', 'WBS']
    return df_out[cols]

# --- UI Layout ---
st.markdown("---")
project_name = st.text_input("Project Name (Top Summary Header)", "NMHIRP Conceptual Design Study - Trans Ramos Pipeline L2 Schedule")
start_d = st.date_input("Project Start Date", datetime.date(2026, 11, 2))

uploaded_file = st.file_uploader("Upload Excel Schedule File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_result = convert_schedule(uploaded_file, project_name, start_d)
        st.success("✅ Schedule successfully converted!")
        
        st.subheader("Data Preview")
        st.dataframe(df_result.head(10))
        
        # Download Button
        csv_data = df_result.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Import-Ready CSV",
            data=csv_data,
            file_name="Converted_L2_Schedule_Import_Ready.csv",
            mime="text/csv",
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
