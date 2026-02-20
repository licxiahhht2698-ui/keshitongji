# 文件名: app.py
import streamlit as st
import pandas as pd
import sqlite3

# 1. 网页标题设置
st.set_page_config(page_title="教师课时管理系统", layout="wide")
st.title("📚 教师课时管理系统 (Web初级版)")

# 连接数据库
conn = sqlite3.connect('school.db')

# 2. 在左侧做一个侧边栏菜单
menu = ["📝 录入课时", "📊 课时汇总看板"]
choice = st.sidebar.selectbox("请选择功能", menu)

if choice == "📝 录入课时":
    st.subheader("新增一条课时记录")
    
    # 从数据库里把老师的名字读取出来，变成下拉菜单
    teachers_df = pd.read_sql("SELECT name FROM teachers", conn)
    teacher_list = teachers_df['name'].tolist()

    # 创建一个表单，供你填数据
    with st.form("add_record_form"):
        t_name = st.selectbox("选择教师", teacher_list)
        month = st.selectbox("选择月份", ["2026-01", "2026-02", "2026-03", "2026-04"])
        course = st.text_input("输入课程名称 (例如：高一语文)")
        hours = st.number_input("输入课时数", min_value=0.0, step=0.5)
        
        # 提交按钮
        submit = st.form_submit_button("保存到系统")

        if submit:
            # 你点击保存后，把数据写进数据库
            c = conn.cursor()
            c.execute("INSERT INTO records (teacher_name, month, course, hours) VALUES (?,?,?,?)", 
                      (t_name, month, course, hours))
            conn.commit()
            st.success(f"太棒了！成功为 {t_name} 保存了 {hours} 个课时！")

elif choice == "📊 课时汇总看板":
    st.subheader("查看所有教师的课时记录")
    
    # 把数据库里的记录全拿出来，直接显示成漂亮的表格
    df = pd.read_sql("SELECT * FROM records", conn)
    
    if df.empty:
        st.info("目前还没有录入任何课时记录哦。")
    else:
        st.dataframe(df, use_container_width=True)