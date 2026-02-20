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

# 初始化网页记忆
if 'all_sheets' not in st.session_state:
    st.session_state['all_sheets'] = None
if 'current_sheet' not in st.session_state:
    st.session_state['current_sheet'] = None

# ================= 2. 强力数据清洗引擎 (修复报错的核心) =================
def clean_excel_data(df):
    """自动寻找真正的表头，并修复空白/重复列名，防止网页崩溃"""
    header_idx = -1
    
    # 1. 寻找包含 "姓名" 或 "科目" 的那一行作为真正的表头
    for i, row in df.iterrows():
        row_str = str(row.values)
        if "姓名" in row_str or "科目" in row_str:
            header_idx = i
            break
            
    if header_idx != -1:
        # 提取真实表头
        raw_cols = df.iloc[header_idx].tolist()
        new_cols = []
        
        # 2. 修复合并单元格带来的空白列和重复列名 (防止 st.data_editor 报错)
        for idx, col in enumerate(raw_cols):
            # 如果是空的，给个默认名字
            col_name = str(col).strip() if pd.notna(col) and str(col).strip() != "" else f"空白列_{idx}"
            # 如果名字重复了，加个后缀
            while col_name in new_cols:
                col_name += "_重复"
            new_cols.append(col_name)
            
        df.columns = new_cols
        # 删掉表头以上的没用排版行
        df = df.iloc[header_idx + 1:].reset_index(drop=True)
        
    # 3. 清理掉全空的行或列
    df = df.dropna(how='all', axis=1).dropna(how='all', axis=0)
    
    # 4. 强制所有列名转为字符串（Streamlit 的硬性要求）
    df.columns = df.columns.astype(str)
    
    return df

# ================= 3. 侧边栏与文件上传 =================
st.sidebar.header("📁 数据中心")
uploaded_file = st.sidebar.file_uploader("请上传您的 xlsm/xlsx 文件", type=["xlsm", "xlsx"])

if uploaded_file is not None and st.session_state['all_sheets'] is None:
    try:
        with st.spinner('正在读取并智能清洗您的 Excel 数据，请稍候...'):
            raw_sheets = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
            clean_sheets = {}
            
            for sheet_name, df in raw_sheets.items():
                clean_sheets[sheet_name] = clean_excel_data(df)
                
            st.session_state['all_sheets'] = clean_sheets
            st.session_state['current_sheet'] = list(clean_sheets.keys())[0]
            st.sidebar.success("✅ 文件清洗并加载成功！")
    except Exception as e:
        st.error(f"读取 Excel 文件时发生致命错误: {e}")

# ================= 4. 动态顶部导航 =================
if st.session_state['all_sheets'] is not None:
    all_sheet_names = list(st.session_state['all_sheets'].keys())
    
    directory_data = {
        "总表 & 汇总": [], "高一年级": [], "高二年级": [], 
        "高三年级": [], "一对一": [], "其他表单": []
    }
    
    for name in all_sheet_names:
        if "总" in name or "分表" in name or "汇总" in name:
            directory_data["总表 & 汇总"].append(name)
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
        # 渲染干净的数据表
        edited_df = st.data_editor(
            df_current, 
            num_rows="dynamic",
            use_container_width=True,
            height=400,
            key=f"editor_{current}" # 增加 key 防止互相干扰
        )
        st.session_state['all_sheets'][current] = edited_df
    except Exception as e:
        st.error(f"表格渲染失败，可能是由于表头格式特殊导致。错误代码: {e}")

    # ================= 6. 智能统计区 =================
    st.markdown("---")
    st.markdown(f"#### 📊 【{current}】各教师课时自动统计")
    
    try:
        # 提取相关列（兼容不同的叫法）
        name_col = next((col for col in edited_df.columns if '姓名' in str(col)), None)
        type_col = next((col for col in edited_df.columns if '子类' in str(col) or '类别' in str(col)), None)
        count_col = next((col for col in edited_df.columns if '课数' in str(col) or '课时' in str(col)), None)

        if name_col and type_col and count_col:
            # 强制转换为数字（把非数字的变成 0）
            edited_df[count_col] = pd.to_numeric(edited_df[count_col], errors='coerce').fillna(0)
            
            # 生成数据透视表
            pivot_df = pd.pivot_table(
                edited_df, 
                values=count_col, 
                index=name_col, 
                columns=type_col, 
                aggfunc='sum', 
                fill_value=0
            )
            
            # 计算每位老师的总计
            pivot_df['总计'] = pivot_df.sum(axis=1)
            st.dataframe(pivot_df, use_container_width=True)
        else:
            st.info("💡 当前表格必须包含【姓名】、【类别/子类】和【课数/课时】的列头，才能自动生成统计报表。")
    except Exception as e:
        st.warning(f"由于数据格式原因，暂无法生成统计表: {e}")

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
        st.sidebar.error("生成下载文件时出错，请检查是否有非法字符。")

else:
    st.info("👆 请先在左侧上传您的 Excel 文件！")