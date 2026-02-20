import streamlit as st
import pandas as pd
import io
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

st.title("📚 教师排课智能读取与统计系统")

if 'all_sheets' not in st.session_state:
    st.session_state['all_sheets'] = None
if 'current_sheet' not in st.session_state:
    st.session_state['current_sheet'] = None

# ================= 2. 智能识别与清洗引擎 =================
def clean_excel_data(df):
    is_schedule = False
    for i in range(min(5, len(df))):
        row_str = " ".join(str(x) for x in df.iloc[i].values)
        if "星期" in row_str or re.search(r'\d{4}[-/]\d{2}[-/]\d{2}', row_str):
            is_schedule = True
            break
            
    if is_schedule:
        new_cols = []
        for idx, col in enumerate(df.columns):
            c = str(col).strip()
            if pd.isna(col) or c.lower() in ['nan', '', 'unnamed'] or 'unnamed' in c.lower():
                c = f"未命名_{idx+1}"
            base = c
            counter = 1
            while c in new_cols:
                c = f"{base}_{counter}"
                counter += 1
            new_cols.append(c)
        df.columns = new_cols
        return df.dropna(how='all', axis=1).dropna(how='all', axis=0)
    else:
        header_idx = -1
        for i in range(min(10, len(df))):
            if any(k in str(df.iloc[i].values) for k in ["姓名", "科目", "类别", "课数"]):
                header_idx = i
                break
        if header_idx != -1:
            raw_cols = df.iloc[header_idx].tolist()
            df = df.iloc[header_idx + 1:].reset_index(drop=True)
        else:
            raw_cols = df.columns.tolist() 
            
        new_cols = []
        for idx, col in enumerate(raw_cols):
            c = str(col).strip()
            if pd.isna(col) or c.lower() in ['nan', '', 'unnamed'] or 'unnamed' in c.lower():
                c = f"未命名_{idx+1}"
            base = c
            counter = 1
            while c in new_cols:
                c = f"{base}_{counter}"
                counter += 1
            new_cols.append(c)
        df.columns = new_cols
        return df.dropna(how='all', axis=1).dropna(how='all', axis=0)

# ================= 3. 侧边栏与文件上传 =================
st.sidebar.header("📁 数据中心")
st.sidebar.info("📌 当前为只读模式，网页仅读取并统计，不会修改您的原文件。")
uploaded_file = st.sidebar.file_uploader("请上传您的 xlsm/xlsx 文件", type=["xlsm", "xlsx"])

if uploaded_file is not None and st.session_state['all_sheets'] is None:
    try:
        with st.spinner('正在执行双引擎解析，请稍候...'):
            raw_sheets = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
            clean_sheets = {}
            for sheet_name, df in raw_sheets.items():
                clean_sheets[sheet_name] = clean_excel_data(df)
            st.session_state['all_sheets'] = clean_sheets
            st.session_state['current_sheet'] = list(clean_sheets.keys())[0]
            st.sidebar.success("✅ 文件解析成功！")
    except Exception as e:
        st.error(f"严重错误: {e}")

# ================= 4. 动态顶部导航 =================
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
    
    display_df = df_current.astype(str)
    display_df = display_df.replace({' 00:00:00': ''}, regex=True)
    display_df = display_df.replace({'nan': '', 'None': ''})
    
    st.dataframe(display_df, use_container_width=True, height=350)

    # ================= 6. 双模式统计区 =================
    st.markdown("---")
    
    tab1, tab2 = st.tabs(["📏 横向排课表拆分与统计 (自动提取时间)", "📊 常规清单表统计 (手动选列)"])
    
    # ---------------- TAB 1：带日期透视的段结构提取逻辑 ----------------
    with tab1:
        st.info("💡 系统已自动扫描表格里的日期。请选择包含具体日期的起始列和结束列：")
        
        all_cols = display_df.columns.tolist()
        
        # 【核心黑科技】：为每一列生成带有时间的漂亮名字
        display_options = []
        for col in all_cols:
            date_info = []
            # 扫描这一列的前3行，寻找日期或星期
            for i in range(min(3, len(display_df))):
                val = str(display_df[col].iloc[i]).strip()
                if re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', val) or "星期" in val:
                    if val and val not in date_info:
                        date_info.append(val)
            
            # 如果找到了时间，就把它拼在列名后面展示
            if date_info:
                display_options.append(f"{col} 📅 {' '.join(date_info)}")
            else:
                display_options.append(col)
        
        col1, col2 = st.columns(2)
        with col1:
            default_start_idx = 14 if len(display_options) > 14 else 0
            start_choice = st.selectbox("🚩 第一步：选择【起始】时间/列", options=display_options, index=default_start_idx)
            
        with col2:
            default_end_idx = 20 if len(display_options) > 20 else len(display_options) - 1
            end_choice = st.selectbox("🏁 第二步：选择【结束】时间/列", options=display_options, index=default_end_idx)
            
        # 根据你选择的漂亮名字，找回真实的列名索引
        start_idx = display_options.index(start_choice)
        end_idx = display_options.index(end_choice)
        
        start_col = all_cols[start_idx]
        end_col = all_cols[end_idx]
        
        if start_idx > end_idx:
            st.error("⚠️ 起始时间不能在结束时间的后面哦，请重新选择！")
        else:
            selected_cols = all_cols[start_idx : end_idx + 1]
            st.success(f"✅ 已锁定范围：包含从 **{start_choice}** 到 **{end_choice}** 的共 {len(selected_cols)} 天数据。")
            
            if st.button("🚀 开始拆分并生成统计报表", type="primary"):
                records = []
                ignore_words = ['0', '0.0', '', 'nan', 'none', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日', '体育', '班会', '国学', '美术', '音乐', '大扫除']
                
                for col in selected_cols:
                    for val in display_df[col]:
                        val_str = str(val).strip()
                        if not val_str or val_str.lower() in ignore_words or re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', val_str):
                            continue
                            
                        match = re.match(r'^([\u4e00-\u9fa5a-zA-Z]+?)(高[一二三]|初[一二三]|小[一二三四五六])(.*)$', val_str)
                        if match:
                            name = match.group(1)
                            ctype = match.group(2) + match.group(3)
                        else:
                            known_types = ['早自', '正大', '正小', '晚自', '自大', '自小', '辅导']
                            name = val_str
                            ctype = "常规课"
                            for kt in known_types:
                                if val_str.endswith(kt):
                                    name = val_str[:-len(kt)]
                                    ctype = kt
                                    break
                                    
                        records.append({'教师姓名': name, '课程类别': ctype, '课时数': 1})
                        
                if records:
                    stat_df = pd.DataFrame(records)
                    pivot_df = pd.pivot_table(stat_df, values='课时数', index='教师姓名', columns='课程类别', aggfunc='sum', fill_value=0)
                    pivot_df['总计'] = pivot_df.sum(axis=1)
                    st.success(f"🎉 提取成功！已精准抓取到 {len(records)} 节有效课时。")
                    st.dataframe(pivot_df, use_container_width=True)
                else:
                    st.warning("⚠️ 在您选定的列范围中，没有找到可以统计的课时数据。")

    # ---------------- TAB 2：常规下拉菜单统计逻辑 ----------------
    with tab2:
        available_cols = list(display_df.columns)
        def guess_index(keywords):
            for i, c in enumerate(available_cols):
                if any(k in str(c) for k in keywords): return i
            return 0
            
        col1, col2, col3 = st.columns(3)
        with col1: name_col = st.selectbox("👤 【教师姓名】列", available_cols, index=guess_index(['姓名', '教师', '未命名_2']))
        with col2: type_col = st.selectbox("🏷️ 【类别】列", available_cols, index=guess_index(['子类', '类别', '科目', '未命名_4']))
        with col3: count_col = st.selectbox("🔢 【数量】列", available_cols, index=guess_index(['课数', '课时', '节数', '未命名_7']))
            
        if st.button("📊 生成常规统计"):
            try:
                stat_df = df_current.copy()
                stat_df[count_col] = pd.to_numeric(stat_df[count_col], errors='coerce').fillna(0)
                stat_df = stat_df[stat_df[name_col].notna()]
                stat_df = stat_df[stat_df[name_col].astype(str).str.strip() != '']
                pivot_df = pd.pivot_table(stat_df, values=count_col, index=name_col, columns=type_col, aggfunc='sum', fill_value=0)
                pivot_df['总计'] = pivot_df.sum(axis=1)
                st.dataframe(pivot_df, use_container_width=True)
            except:
                st.warning("无法生成，请确认选对了列名哦！")

else:
    st.info("👆 请先在左侧上传您的 Excel 文件！")