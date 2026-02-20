import streamlit as st
import pandas as pd
import io

# ================= 1. 网页基础设置 =================
st.set_page_config(page_title="教师课时管理系统", page_icon="📚", layout="wide")
st.title("📚 教师课时智能管理平台")

# ================= 2. 网页的“记忆力”(非常重要) =================
# 因为网页每次点击按钮都会刷新，我们需要用 session_state 让网页“记住”你上传和修改过的数据
if 'all_sheets' not in st.session_state:
    st.session_state['all_sheets'] = None

# ================= 3. 侧边栏与文件上传 =================
st.sidebar.header("📁 数据中心")
uploaded_file = st.sidebar.file_uploader("首次使用，请先上传您的 xlsm/xlsx 文件", type=["xlsm", "xlsx"])

# 如果上传了新文件，立刻把它读取进网页的“记忆”里
if uploaded_file is not None and st.session_state['all_sheets'] is None:
    with st.spinner('正在疯狂解析您的 Excel 结构...'):
        # 读取所有的 Sheet
        st.session_state['all_sheets'] = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
        st.sidebar.success("✅ 文件读取成功！")

# ================= 4. 核心功能区 (只有上传数据后才显示) =================
if st.session_state['all_sheets'] is not None:
    
    # 提取所有工作表的名字
    sheet_names = list(st.session_state['all_sheets'].keys())
    
    # 设置功能导航
    menu = st.sidebar.radio("🧭 功能导航", ["✏️ 课时在线编辑", "📊 自动汇总大屏", "💾 下载最新数据"])
    
    # ---------------- 页面 A：课时在线编辑 ----------------
    if menu == "✏️ 课时在线编辑":
        st.subheader("✏️ 在线编辑排课与课时数据")
        st.info("💡 提示：双击下方表格的单元格即可修改内容。你还可以在表格最下方点击 '+' 添加新行！")
        
        # 让用户选择要编辑的月份/Sheet
        target_sheet = st.selectbox("请选择要编辑的月份或工作表:", sheet_names)
        
        # 获取当前 Sheet 的数据
        df_current = st.session_state['all_sheets'][target_sheet]
        
        # 【黑科技登场】生成可编辑表格！num_rows="dynamic" 允许你增加或删除行
        edited_df = st.data_editor(
            df_current, 
            num_rows="dynamic",
            use_container_width=True,
            height=500
        )
        
        # 把你在网页上修改好的数据，重新存回网页的“记忆”里
        st.session_state['all_sheets'][target_sheet] = edited_df
        st.success(f"当前对 {target_sheet} 的修改已临时保存在网页中！")

    # ---------------- 页面 B：自动汇总大屏 ----------------
    elif menu == "📊 自动汇总大屏":
        st.subheader("📊 全校课时智能汇总")
        st.write("这里可以替代你以前 Excel 里的复杂公式，用 Python 直接算！")
        
        # 假设你要汇总刚才选的那个 Sheet (你可以根据实际情况调整)
        target_sheet = st.selectbox("选择要分析的月份:", sheet_names)
        df_to_analyze = st.session_state['all_sheets'][target_sheet]
        
        # 假设你的 Excel 里有一列叫 "教师姓名"，一列叫 "课时数"
        # 这里教你一段 Python 分组求和的魔法 (你需要根据你实际的列名修改中文字符串)
        try:
            st.markdown(f"### {target_sheet} 汇总报表")
            # 找到你表格里的列名，这里需要替换成你 Excel 里真实的表头名字！
            # 例如：teacher_col = "姓名", hours_col = "实际课时"
            teacher_col = st.selectbox("请选择表示【教师姓名】的列", df_to_analyze.columns)
            hours_col = st.selectbox("请选择表示【课时数】的列", df_to_analyze.columns)
            
            # 一行代码完成 Excel 里的按人头汇总！
            summary_df = df_to_analyze.groupby(teacher_col)[hours_col].sum().reset_index()
            
            # 画一个简单的柱状图
            st.bar_chart(data=summary_df, x=teacher_col, y=hours_col)
            # 显示汇总表格
            st.dataframe(summary_df, use_container_width=True)
            
        except Exception as e:
            st.warning("请确保你选择了包含数字的列来进行汇总计算哦！")

    # ---------------- 页面 C：下载最新数据 ----------------
    elif menu == "💾 下载最新数据":
        st.subheader("💾 将修改后的数据导出为 Excel")
        st.warning("⚠️ Streamlit 云端不会永久保存数据！关闭网页前，请务必下载保存你的修改结果！")
        
        # 把内存里的数据打包成一个新的 Excel 文件
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in st.session_state['all_sheets'].items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        processed_data = output.getvalue()
        
        # 生成下载按钮
        st.download_button(
            label="⬇️ 点击下载最新版本的 Excel 文件",
            data=processed_data,
            file_name="教师课时管理_网站更新版.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("等待上传文件中...")
    st.image("https://images.unsplash.com/photo-1434030216411-0b793f4b4173?auto=format&fit=crop&q=80&w=1000", caption="告别繁琐，拥抱高效")