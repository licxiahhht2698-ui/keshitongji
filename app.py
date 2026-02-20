import streamlit as st
import pandas as pd
import io

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

st.title("📚 教师课时智能管理平台")

if 'all_sheets' not in st.session_state:
    st.session_state['all_sheets'] = None
if 'current_sheet' not in st.session_state:
    st.session_state['current_sheet'] = None

# ================= 2. 终极防御数据清洗引擎 =================
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

# ================= 3. 侧边栏与文件上传 =================
st.sidebar.header("📁 数据中心")
uploaded_file = st.sidebar.file_uploader("请上传您的 xlsm/xlsx 文件", type=["xlsm", "xlsx"])

if uploaded_file is not None and st.session_state['all_sheets'] is None:
    try:
        with st.spinner('正在执行终极防崩溃算法解析，请稍候...'):
            raw_sheets = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
            clean_sheets = {}
            for sheet_name, df in raw_sheets.items():
                clean_sheets[sheet_name] = clean_excel_data(df)
            st.session_state['all_sheets'] = clean_sheets
            st.session_state['current_sheet'] = list(clean_sheets.keys())[0]
            st.sidebar.success("✅ 文件清洗并加载成功！")
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

    # ================= 5. 核心编辑区 =================
    current = st.session_state['current_sheet']
    st.markdown(f"#### ✏️ 当前编辑 : 【 {current} 】")
    df_current = st.session_state['all_sheets'][current]
    
    try:
        edited_df = st.data_editor(
            df_current, 
            num_rows="dynamic",
            use_container_width=True,
            height=400,
            key=f"editor_{current}"
        )
        st.session_state['all_sheets'][current] = edited_df
    except Exception as e:
        st.error(f"渲染失败。错误详情: {e}")

    # ================= 6. 智能统计区 (新增了用户自选功能！) =================
    st.markdown("---")
    st.markdown(f"#### 📊 【{current}】各教师课时自动统计")
    
    available_cols = list(edited_df.columns)
    
    if len(available_cols) > 0:
        # 智能猜一下哪几列是我们要的
        def guess_index(keywords):
            for i, col in enumerate(available_cols):
                if any(k in str(col) for k in keywords):
                    return i
            return 0
            
        idx_name = guess_index(['姓名', '教师', '老师'])
        idx_type = guess_index(['子类', '类别', '科目'])
        idx_count = guess_index(['课数', '课时', '节数'])
        
        # 💡 在网页上生成三个下拉菜单！
        st.info("💡 如果下方的统计没出来，或者统计错了，请在这里手动选择对应的列：")
        col1, col2, col3 = st.columns(3)
        with col1:
            name_col = st.selectbox("👤 哪一列是【教师姓名】？", available_cols, index=idx_name, key=f"sel_name_{current}")
        with col2:
            type_col = st.selectbox("🏷️ 哪一列是【课时类别/早自晚自】？", available_cols, index=idx_type, key=f"sel_type_{current}")
        with col3:
            count_col = st.selectbox("🔢 哪一列是【课时数量】？", available_cols, index=idx_count, key=f"sel_count_{current}")
            
        try:
            # 根据你选择的列来进行计算
            stat_df = edited_df.copy()
            stat_df[count_col] = pd.to_numeric(stat_df[count_col], errors='coerce').fillna(0)
            
            # 过滤掉一些没用的空行，让统计表更干净
            stat_df = stat_df[stat_df[name_col].notna()]
            stat_df = stat_df[stat_df[name_col].astype(str).str.strip() != '']
            stat_df = stat_df[stat_df[name_col].astype(str).str.strip() != '0']
            
            # 生成透视表
            pivot_df = pd.pivot_table(
                stat_df, 
                values=count_col, 
                index=name_col, 
                columns=type_col, 
                aggfunc='sum', 
                fill_value=0
            )
            
            pivot_df['总计'] = pivot_df.sum(axis=1)
            st.dataframe(pivot_df, use_container_width=True)
            
        except Exception as e:
            st.warning(f"无法生成统计表，请确保选择的列正确哦。({e})")

    # ---------------- 下载最新数据 ----------------
    st.sidebar.divider()
    st.sidebar.subheader("💾 保存与下载")
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in st.session_state['all_sheets'].items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        processed_data = output.getvalue()
        st.sidebar.download_button("⬇️ 下载最新版 Excel", data=processed_data, file_name="最新课时统计_已清理.xlsx")
    except Exception as e:
        pass