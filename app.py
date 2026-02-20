import streamlit as st
import pandas as pd
import re

# ================= 1. 网页基础设置 =================
st.set_page_config(page_title="教师课时管理系统", page_icon="📚", layout="wide")

st.markdown("""
<style>
    div.stButton > button {
        white-space: nowrap !important; 
        font-size: 13px !important;     
        padding: 2px 8px !important;    
        min-height: 28px !important; 
        height: 28px !important;
        width: 100% !important;         
        background-color: #e2efda;      
        color: #333333;
        border: 1px solid #a9d08e;
        border-radius: 3px;
    }
    div.stButton > button:hover {
        background-color: #c6e0b4;
        color: black;
        border-color: #548235;
    }
    .row-title {
        font-size: 13px;
        font-weight: bold;
        color: #385723;
        text-align: left;               
        padding-top: 5px;
        white-space: nowrap;
    }
    [data-testid="column"] { padding: 0 4px !important; }
</style>
""", unsafe_allow_html=True)

st.title("📚 教师排课表智能读取与统计系统")

if 'all_sheets' not in st.session_state:
    st.session_state['all_sheets'] = None
if 'current_sheet' not in st.session_state:
    st.session_state['current_sheet'] = None

# ================= 2. 数据清洗引擎 =================
def clean_excel_data(df):
    header_idx = -1
    for i in range(min(10, len(df))):
        row_str = str(df.iloc[i].values)
        if any(keyword in row_str for keyword in ["姓名", "科目", "班级", "教师", "序号", "早自", "类别", "课数"]):
            header_idx = i
            break
            
    if header_idx != -1:
        raw_cols = df.iloc[header_idx].tolist()
        df = df.iloc[header_idx + 1:].reset_index(drop=True)
    else:
        raw_cols = df.columns.tolist() 
        
    new_cols = []
    for idx, col in enumerate(raw_cols):
        col_str = str(col).strip()
        if pd.isna(col) or col_str.lower() in ['nan', 'none', 'nat', '', 'unnamed']:
            base_name = f"未命名_{idx+1}"
        elif "unnamed" in col_str.lower():
            base_name = f"未命名_{idx+1}"
        else:
            base_name = col_str
            
        final_name = base_name
        counter = 1
        while final_name in new_cols:
            final_name = f"{base_name}_{counter}"
            counter += 1
        new_cols.append(final_name)
        
    df.columns = new_cols
    df = df.dropna(how='all', axis=1).dropna(how='all', axis=0)
    return df

# ================= 3. 文件上传 =================
st.sidebar.header("📁 数据中心")
st.sidebar.info("📌 当前版本为只读模式，所有数据均从 Excel 中提取，不会修改原文件。")
uploaded_file = st.sidebar.file_uploader("请上传您的 xlsm/xlsx 文件", type=["xlsm", "xlsx"])

if uploaded_file is not None and st.session_state['all_sheets'] is None:
    try:
        with st.spinner('正在解析并提取课表...'):
            raw_sheets = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
            clean_sheets = {}
            for sheet_name, df in raw_sheets.items():
                clean_sheets[sheet_name] = clean_excel_data(df)
            st.session_state['all_sheets'] = clean_sheets
            st.session_state['current_sheet'] = list(clean_sheets.keys())[0]
            st.sidebar.success("✅ 文件解析成功！")
    except Exception as e:
        st.error(f"严重错误: {e}")

# ================= 4. 动态导航 =================
if st.session_state['all_sheets'] is not None:
    all_sheet_names = list(st.session_state['all_sheets'].keys())
    directory_data = {
        "总表 & 汇总": [], "高一年级": [], "高二年级": [], 
        "高三年级": [], "一对一": [], "其他表单": []
    }
    for name in all_sheet_names:
        if "总" in name or "分表" in name or "汇总" in name: directory_data["总表 & 汇总"].append(name)
        elif "高一" in name: directory_data["高一年级"].append(name)
        elif "高二" in name: directory_data["高二年级"].append(name)
        elif "高三" in name: directory_data["高三年级"].append(name)
        elif "一对一" in name: directory_data["一对一"].append(name)
        else: directory_data["其他表单"].append(name)

    st.markdown("<hr style='margin: 5px 0px;'>", unsafe_allow_html=True)
    for category, buttons in directory_data.items():
        if not buttons: continue 
        empty_space = 10 - len(buttons) if len(buttons) < 10 else 1
        cols = st.columns([1.2] + [1] * len(buttons) + [empty_space]) 
        with cols[0]:
            st.markdown(f'<div class="row-title">{category} :</div>', unsafe_allow_html=True)
        for i, btn_name in enumerate(buttons):
            with cols[i+1]:
                if st.button(btn_name, key=f"nav_{btn_name}"):
                    st.session_state['current_sheet'] = btn_name
    st.markdown("<hr style='margin: 5px 0px;'>", unsafe_allow_html=True)

    # ================= 5. 只读展示区 =================
    current = st.session_state['current_sheet']
    st.markdown(f"#### 👁️ 当前查看 : 【 {current} 】")
    
    df_current = st.session_state['all_sheets'][current].copy()
    
    # 【核心格式化】：把所有的 00:00:00 去掉，把 nan 变为空白
    df_current = df_current.astype(str)
    df_current = df_current.replace({' 00:00:00': ''}, regex=True)
    df_current = df_current.replace({'nan': ''})
    
    # 彻底改为只读模式 st.dataframe，不再使用编辑器
    st.dataframe(df_current, use_container_width=True, height=350)

    # ================= 6. 带时间范围的智能统计区 =================
    st.markdown("---")
    
    # 步骤 1：自动检测第一行里是不是包含日期 (寻找 2025-12-01 这种格式)
    date_cols = {}
    if len(df_current) > 0:
        for col in df_current.columns:
            val_str = str(df_current.loc[0, col]).strip()
            # 如果符合 YYYY-MM-DD 格式，就记录下来它对应的列名
            if re.match(r'^\d{4}-\d{2}-\d{2}$', val_str):
                date_cols[val_str] = col

    # 如果检测到了日期列（这就是你截图里的横向排课表）
    if date_cols:
        st.markdown(f"#### 📅 【{current}】日期范围课时统计")
        st.success("✨ 系统检测到当前为排课表，已开启按日期范围自动提取统计功能！")
        
        dates = sorted(list(date_cols.keys()))
        min_date = pd.to_datetime(dates[0]).date()
        max_date = pd.to_datetime(dates[-1]).date()

        # 生成日期范围选择器
        selected_dates = st.date_input("🗓️ 请选择要统计的日期范围：", [min_date, max_date], min_value=min_date, max_value=max_date)

        if len(selected_dates) == 2:
            start_date, end_date = selected_dates
            
            # 找到在所选时间范围内的真实列名 (如 未命名_15, 未命名_16)
            valid_cols = []
            for d_str, c_name in date_cols.items():
                if start_date <= pd.to_datetime(d_str).date() <= end_date:
                    valid_cols.append(c_name)

            # 从第 3 行开始（跳过日期行和星期行），提取所有排课数据
            all_classes = []
            for col in valid_cols:
                if len(df_current) > 2:
                    cells = df_current[col].iloc[2:].dropna().astype(str).tolist()
                    all_classes.extend(cells)

            # 过滤垃圾词汇，并拆分姓名和课程类型
            records = []
            ignore_words = ['0', '0.0', '', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日', '体育', '班会', '国学', '美术', '音乐']
            
            for item in all_classes:
                item = item.strip()
                if not item or item in ignore_words: 
                    continue
                
                # 【智能拆词】：寻找 "高一", "高二", "高三" 的位置，左边是名字，右边是班级和类型
                idx = max(item.rfind("高一"), item.rfind("高二"), item.rfind("高三"))
                if idx != -1:
                    name = item[:idx]
                    type_str = item[idx:]
                    records.append({'教师姓名': name, '课程类别': type_str, '课时数': 1})
                else:
                    # 如果找不到“高”，尝试直接看最后两三个字（如 "早自"）
                    records.append({'教师姓名': item, '课程类别': '其他课时', '课时数': 1})

            if records:
                # 生成漂亮的透视统计表
                stat_df = pd.DataFrame(records)
                pivot_df = pd.pivot_table(stat_df, values='课时数', index='教师姓名', columns='课程类别', aggfunc='sum', fill_value=0)
                pivot_df['总计'] = pivot_df.sum(axis=1)
                st.dataframe(pivot_df, use_container_width=True)
            else:
                st.info("💡 在您选择的日期范围内，没有找到有效的教师排课记录哦。")

    # 如果不是横向排课表（比如汇总表），走老规矩下拉菜单逻辑
    else:
        st.markdown(f"#### 📊 【{current}】常规课时自动统计")
        available_cols = list(df_current.columns)
        
        def guess_index(keywords):
            for i, col in enumerate(available_cols):
                if any(k in str(col) for k in keywords): return i
            return 0
            
        col1, col2, col3 = st.columns(3)
        with col1: name_col = st.selectbox("👤 【教师姓名】列", available_cols, index=guess_index(['姓名', '教师']))
        with col2: type_col = st.selectbox("🏷️ 【类别】列", available_cols, index=guess_index(['子类', '类别', '科目']))
        with col3: count_col = st.selectbox("🔢 【数量】列", available_cols, index=guess_index(['课数', '课时', '节数']))
            
        try:
            stat_df = df_current.copy()
            stat_df[count_col] = pd.to_numeric(stat_df[count_col], errors='coerce').fillna(0)
            stat_df = stat_df[stat_df[name_col].notna()]
            stat_df = stat_df[stat_df[name_col].astype(str).str.strip() != '']
            
            pivot_df = pd.pivot_table(stat_df, values=count_col, index=name_col, columns=type_col, aggfunc='sum', fill_value=0)
            pivot_df['总计'] = pivot_df.sum(axis=1)
            st.dataframe(pivot_df, use_container_width=True)
        except:
            st.warning("请确保选择了正确的列。")

else:
    st.info("👆 请先在左侧上传您的 Excel 文件！")