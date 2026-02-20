import streamlit as st
import pandas as pd
import io

# ================= 1. 网页基础设置 & 五号字体样式 =================
st.set_page_config(page_title="教师课时管理系统", page_icon="📚", layout="wide")

# 注入 CSS 魔法：设置按钮为五号字体(14px)，并模仿你截图中的浅绿色风格
st.markdown("""
<style>
    /* 控制按钮的样式：五号字体(14px)，浅绿色背景，棕色边框 */
    div.stButton > button {
        font-size: 14px !important; 
        width: 100%;
        background-color: #d8e4bc; 
        color: #333333;
        border: 1px solid #8e9e63;
        padding: 5px 0px;
        margin-bottom: 2px;
    }
    div.stButton > button:hover {
        background-color: #c4d79b;
        color: black;
        border-color: #4f6228;
    }
    /* 控制列标题的样式 */
    .dir-title {
        text-align: center;
        font-size: 16px;
        font-weight: bold;
        color: #604a0e;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

st.title("📚 教师课时智能管理平台")

# 初始化网页的记忆（当前选中的工作表）
if 'all_sheets' not in st.session_state:
    st.session_state['all_sheets'] = None
if 'current_sheet' not in st.session_state:
    st.session_state['current_sheet'] = "汇总表" # 默认打开的表

# ================= 2. 侧边栏与文件上传 =================
st.sidebar.header("📁 数据中心")
uploaded_file = st.sidebar.file_uploader("首次使用，请先上传您的 xlsm/xlsx 文件", type=["xlsm", "xlsx"])

if uploaded_file is not None and st.session_state['all_sheets'] is None:
    with st.spinner('正在疯狂解析您的 Excel 结构...'):
        st.session_state['all_sheets'] = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
        st.sidebar.success("✅ 文件读取成功！")

# ================= 3. 顶部导航目录 (你的截图结构) =================
# 用一个字典把你的目录结构存起来
directory_data = {
    "总表": ["汇总表", "分表"],
    "高一年级": [f"高一{i}班" for i in range(1, 9)],
    "高二年级": [f"高二{i}班" for i in range(1, 9)],
    "高三年级": ["高三生物1班", "高三生物2班", "高三地理1班", "高三地理2班", "高三政治班"],
    "一对一": ["一对一", "一对一档案"]
}

# 在网页顶部划出 5 个等宽的列
cols = st.columns(5)

# 自动生成这 5 列的按钮
for i, (category, buttons) in enumerate(directory_data.items()):
    with cols[i]:
        # 写入大标题（比如：高一年级）
        st.markdown(f'<div class="dir-title">{category}</div>', unsafe_allow_html=True)
        # 生成这一列下面的所有按钮
        for btn_name in buttons:
            if st.button(btn_name):
                # 如果按钮被点击，就让网页记住当前要看哪个表
                st.session_state['current_sheet'] = btn_name


st.divider() # 画一条分割线

# ================= 4. 核心编辑区 (点击上方按钮后联动) =================
if st.session_state['all_sheets'] is not None:
    
    current = st.session_state['current_sheet']
    st.subheader(f"✏️ 正在编辑: 【{current}】")
    
    # 检查你点击的班级，在你的 Excel 里到底存不存在这个 Sheet
    if current in st.session_state['all_sheets']:
        df_current = st.session_state['all_sheets'][current]
        
        # 生成可编辑的表格
        edited_df = st.data_editor(
            df_current, 
            num_rows="dynamic",
            use_container_width=True,
            height=600
        )
        # 实时保存修改
        st.session_state['all_sheets'][current] = edited_df
        
    else:
        st.warning(f"⚠️ 在您上传的 Excel 文件中，没有找到名为 '{current}' 的工作表哦！请检查 Excel 的底部标签名是否对应。")

    # ---------------- 下载最新数据 ----------------
    st.sidebar.divider()
    st.sidebar.subheader("💾 导出数据")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in st.session_state['all_sheets'].items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    processed_data = output.getvalue()
    st.sidebar.download_button("⬇️ 下载修改后的 Excel", data=processed_data, file_name="最新课时统计.xlsx")

else:
    st.info("👆 请先在左侧上传包含这些班级数据的 Excel 文件哦！")