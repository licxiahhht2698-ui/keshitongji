import streamlit as st
import pandas as pd
import io

# ================= 1. 网页基础设置 & 极致横排 CSS =================
st.set_page_config(page_title="教师课时管理系统", page_icon="📚", layout="wide")

st.markdown("""
<style>
    /* 强制按钮文字横向显示，左对齐，缩小字体和底色 */
    div.stButton > button {
        white-space: nowrap !important; /* 【核心】绝对禁止文字换行，解决竖排问题 */
        font-size: 13px !important;     /* 缩小字体 */
        padding: 2px 8px !important;    /* 缩小内部留白 */
        min-height: 28px !important; 
        height: 28px !important;
        width: 100% !important;         
        background-color: #e2efda;      /* 更淡的浅绿色底色，不突兀 */
        color: #333333;
        border: 1px solid #a9d08e;
        border-radius: 3px;
    }
    div.stButton > button:hover {
        background-color: #c6e0b4;
        color: black;
        border-color: #548235;
    }
    /* 分类标题样式：靠左对齐 */
    .row-title {
        font-size: 13px;
        font-weight: bold;
        color: #385723;
        text-align: left;               /* 【核心】整体左对齐 */
        padding-top: 5px;
        white-space: nowrap;
    }
    /* 调整列间距，紧凑排列 */
    [data-testid="column"] {
        padding: 0 4px !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("📚 教师课时智能管理平台")

# 初始化记忆
if 'all_sheets' not in st.session_state:
    st.session_state['all_sheets'] = None
if 'current_sheet' not in st.session_state:
    st.session_state['current_sheet'] = None

# ================= 2. 侧边栏与文件上传 =================
st.sidebar.header("📁 数据中心")
uploaded_file = st.sidebar.file_uploader("请上传您的 xlsm/xlsx 文件", type=["xlsm", "xlsx"])

if uploaded_file is not None and st.session_state['all_sheets'] is None:
    with st.spinner('正在读取您的 Excel 数据...'):
        st.session_state['all_sheets'] = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
        # 默认打开真实存在的第一个表
        st.session_state['current_sheet'] = list(st.session_state['all_sheets'].keys())[0]
        st.sidebar.success("✅ 文件读取成功！")

# ================= 3. 动态顶部导航 (随Excel自动变化) =================
if st.session_state['all_sheets'] is not None:
    
    # 1. 实时获取你 Excel 里真实存在的所有表名
    all_sheet_names = list(st.session_state['all_sheets'].keys())
    
    # 2. 准备一个空的分类夹
    directory_data = {
        "总表 & 汇总": [],
        "高一年级": [],
        "高二年级": [],
        "高三年级": [],
        "一对一": [],
        "其他表单": []
    }
    
    # 3. 智能分类（无论你怎么增减表，只要名字里带这些字，就会自动归类）
    for name in all_sheet_names:
        if "总" in name or "分表" in name or "汇总" in name:
            directory_data["总表 & 汇总"].append(name)
        elif "高一" in name:
            directory_data["高一年级"].append(name)
        elif "高二" in name:
            directory_data["高二年级"].append(name)
        elif "高三" in name:
            directory_data["高三年级"].append(name)
        elif "一对一" in name:
            directory_data["一对一"].append(name)
        else:
            directory_data["其他表单"].append(name)

    st.markdown("<hr style='margin: 5px 0px;'>", unsafe_allow_html=True)
    
    # 4. 渲染导航栏 (整体左对齐)
    for category, buttons in directory_data.items():
        if not buttons: 
            continue # 如果这个类别下没有表，就直接跳过不显示，保持界面干净
            
        # 左对齐魔法：[1.2]是标题宽，[1]*按钮数是按钮宽，最后加个[10]大空白把它们全部挤到左边！
        # 做了安全处理，防止按钮太多超出列的限制
        empty_space = 10 - len(buttons) if len(buttons) < 10 else 1
        cols = st.columns([1.2] + [1] * len(buttons) + [empty_space]) 
        
        with cols[0]:
            st.markdown(f'<div class="row-title">{category} :</div>', unsafe_allow_html=True)
            
        for i, btn_name in enumerate(buttons):
            with cols[i+1]:
                if st.button(btn_name, key=f"nav_{btn_name}"):
                    st.session_state['current_sheet'] = btn_name

    st.markdown("<hr style='margin: 5px 0px;'>", unsafe_allow_html=True)

    # ================= 4. 核心编辑区 =================
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
    st.info("👆 请先在左侧上传您的 Excel 文件，随后系统会自动生成专属导航目录！")