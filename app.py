import streamlit as st
import pandas as pd
import io

# ================= 1. 网页基础设置 =================
st.set_page_config(page_title="教师课时管理系统", page_icon="📚", layout="wide")

# 注入 CSS 魔法：极致压缩按钮空间，超小字体
st.markdown("""
<style>
    /* 极致压缩按钮：11px超小字体，极窄边距，高度仅24px */
    div.stButton > button {
        font-size: 11px !important; 
        padding: 0px 2px !important;
        min-height: 24px !important; 
        height: 24px !important;
        width: 100%;
        background-color: #d8e4bc; 
        color: #333333;
        border: 1px solid #8e9e63;
        border-radius: 2px; /* 让边角更像 Excel 单元格，而不是圆角 */
    }
    div.stButton > button:hover {
        background-color: #c4d79b;
        color: black;
        border-color: #4f6228;
    }
    /* 极致压缩分类标题 */
    .row-title {
        font-size: 12px;
        font-weight: bold;
        color: #604a0e;
        text-align: right;
        padding-top: 3px;
        white-space: nowrap; /* 防止标题换行 */
    }
    /* 缩小列与列之间的间距 */
    [data-testid="column"] {
        padding: 0 2px !important; 
    }
    /* 隐藏顶部多余空白 */
    .block-container {
        padding-top: 2rem;
    }
</style>
""", unsafe_allow_html=True)

st.title("📚 教师课时智能管理平台")

# 初始化网页的记忆
if 'all_sheets' not in st.session_state:
    st.session_state['all_sheets'] = None
if 'current_sheet' not in st.session_state:
    st.session_state['current_sheet'] = "汇总表"

# ================= 2. 侧边栏与文件上传 =================
st.sidebar.header("📁 数据中心")
uploaded_file = st.sidebar.file_uploader("请上传您的 xlsm/xlsx 文件", type=["xlsm", "xlsx"])

if uploaded_file is not None and st.session_state['all_sheets'] is None:
    with st.spinner('正在读取您的 Excel 数据...'):
        st.session_state['all_sheets'] = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
        st.sidebar.success("✅ 文件读取成功！")

# ================= 3. 顶部导航 (极简左对齐横排) =================
directory_data = {
    "总表": ["汇总表", "分表"],
    "高一年级": [f"高一{i}班" for i in range(1, 9)],
    "高二年级": [f"高二{i}班" for i in range(1, 9)],
    "高三年级": ["高三生物1班", "高三生物2班", "高三地理1班", "高三地理2班", "高三政治班"],
    "一对一": ["一对一", "一对一档案"]
}

st.markdown("<hr style='margin: 5px 0px;'>", unsafe_allow_html=True) # 超窄分割线

# 按行（横排）生成目录
for category, buttons in directory_data.items():
    # 【核心左对齐魔法】：[1.2]放标题，[1]*数量放按钮，[10]放一个巨大的空列把所有东西往左挤！
    cols = st.columns([1.2] + [1] * len(buttons) + [10]) 
    
    with cols[0]:
        st.markdown(f'<div class="row-title">{category} :</div>', unsafe_allow_html=True)
        
    for i, btn_name in enumerate(buttons):
        with cols[i+1]:
            # 生成按钮
            if st.button(btn_name, key=btn_name):
                st.session_state['current_sheet'] = btn_name

st.markdown("<hr style='margin: 5px 0px;'>", unsafe_allow_html=True) # 超窄分割线

# ================= 4. 核心编辑区 =================
if st.session_state['all_sheets'] is not None:
    
    current = st.session_state['current_sheet']
    st.markdown(f"#### ✏️ 当前编辑 : 【 {current} 】")
    
    if current in st.session_state['all_sheets']:
        df_current = st.session_state['all_sheets'][current]
        
        edited_df = st.data_editor(
            df_current, 
            num_rows="dynamic",
            use_container_width=True,
            height=550
        )
        st.session_state['all_sheets'][current] = edited_df
        
    else:
        st.warning(f"⚠️ 在上传的 Excel 中没有找到 '{current}' 工作表。")

    # ---------------- 下载最新数据 ----------------
    st.sidebar.divider()
    st.sidebar.subheader("💾 保存与下载")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in st.session_state['all_sheets'].items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    processed_data = output.getvalue()
    st.sidebar.download_button("⬇️ 下载最新版 Excel", data=processed_data, file_name="最新课时统计.xlsx")

else:
    st.info("👆 请先在左侧上传您的 Excel 文件，随后即可点击上方横排按钮切换班级！")