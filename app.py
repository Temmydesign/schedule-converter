import streamlit as st
import pandas as pd
import numpy as np
import re
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
    # Read Excel Sheet
    df_raw = pd.read_excel(uploaded_file, sheet_name='Integrated L2 Schedule', skiprows=2)
    df_data = df_raw.iloc[1:49, :7].dropna(how='all')
    df_data.columns = ['System', 'Activity_Name', 'Start_Wk', 'Dur_Wks', 'Finish_Wk', 'Start_Date', 'Finish_Date']

    tasks_list = []
    
    # 1. Top-Level Project Title Header (Level 1)
    tasks_list.append({
        'old_id': 'PROJECT_ROOT',
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
            'old_id': sys_sum_old_id,
            'Task Name': f"System: {sys_name}",
            'Duration': '',
            'Start Date': str(group['Start_Date'].min()).split(' ')[0],
            'Finish Date': str(group['Finish_Date'].max()).split(' ')[0],
            'Predecessors': '',
            'Outline Level': 2
        })
        temp_id += 1
        
        sub_tasks = []
        for idx, r in group.iterrows():
            dur_str = f"{int(r['Dur_Wks'])} wks" if r['Dur_Wks'] > 0 else "0 wks"
            sub_tasks.append({
                'excel_idx': idx,
                'Task Name': r['Activity_Name'],
                'Duration': dur_str,
                'Start Date': str(r['Start_Date']).split(' ')[0],
                'Finish Date': str(r['Finish_Date']).split(' ')[0],
                'Outline Level': 3,
                'System': sys_name,
                'Dur_Wks': r['Dur_Wks']
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
            if sys_name == 'PRG':
                if i == 1:
                    preds.append("3SS")
                elif i == 2:
                    preds.append("3FS")
            elif sys_name in ['SSB', 'NCTL', 'TNP', 'TFP', 'TRP', 'EWGP']:
                t1 = id_map[f"ACT_{sys_name}_1"]
                if i == 0:
                    preds.append("3SS")
                elif i == 1:
                    preds.append(f"{t1}FS-1 wk")
                elif i == 2:
                    preds.append(f"{t1}SS+2 wks")
                elif i == 3 or i == 4:
                    t3 = id_map[f"ACT_{sys_name}_3"]
                    preds.append(f"{t3}FS")
                elif i == 5:
                    t5 = id_map[f"ACT_{sys_name}_5"]
                    preds.append(f"{t5}SS+6 wks")
                elif i == 6:
                    t6 = id_map[f"ACT_{sys_name}_6"]
                    preds.append(f"{t6}FS-1 wk")
            elif sys_name == 'MS':
                if i == 0:
                    preds.append("11, 20, 29, 38, 47, 56")
                elif i == 1:
                    preds.append("13, 22, 31, 40, 49, 58")
                elif i == 2:
                    preds.append("14, 23, 32, 41, 50, 59")
            
            tasks_list.append({
                'old_id': task_old_id,
                'Task Name': task['Task Name'],
                'Duration': task['Duration'],
                'Start Date': task['Start Date'],
                'Finish Date': task['Finish Date'],
                'Predecessors': ", ".join(preds),
                'Outline Level': 3
            })
            temp_id += 1
            
        # Completion Milestone
        m_name = f"{sys_name} completed" if sys_name != 'MS' else "Programme Milestones completed"
        m_pred = f"{latest_act_id}FS" if sys_name != 'MS' else "64FS"
        
        tasks_list.append({
            'old_id': f"MILESTONE_{sys_name}",
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
project_name = st.text_input("Project Name (Top Summary Header)", "NMHIRP Conceptual Design Study - Integrated Level 2 Schedule")
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
            file_name="NMHIRP_L2_Schedule_Import_Ready.csv",
            mime="text/csv",
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
