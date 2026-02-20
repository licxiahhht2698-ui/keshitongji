import streamlit as st
import pandas as pd
import io
import re

# ================= 1. 网页基础设置 =================
st.set_page_config(page_title="教师课时管理系统", page_icon="📚", layout="wide")

st.markdown("""
<style>
    div.stButton > button {
        white-space: nowrap !important; font-size: 13px !important;     
        padding: 2px 8px !important; min-height: 28px !important; 
        height: 28px !important; width: 100% !important;         
        background-color: #e2efda; color: #333333;
        border: 1px solid #a9d08e; border-radius: 3px;
    }
    div.stButton > button:hover { background-color: #c6e0b4; color: black; border-color: #548235; }
    .row-title { font-size: 13px; font-weight: bold; color: #385723; text-align: left; padding-top: 5px; white-space: nowrap; }
    [data-testid="column"] { padding: 0 4px !important; }
</style>
""", unsafe_allow_html=True)

st.title("📚 教师排课智能读取与精准统计系统")

if 'all_sheets' not in st.session_state: st.session_state['all_sheets'] = None
if 'current_sheet' not in st.session_state: st.session_state['current_sheet'] = None

# ================= 2. 智能识别与清洗引擎 =================
def clean_excel_data(df):
    is_schedule = False
    for i in range(min(5, len(df))):
        row_str = " ".join(str(x) for x in df.iloc[i].values)
        if "星期" in row_str or re.search(r'\d{4}[-/]\d{2}[-/]\d{2}', row_str):
            is_schedule = True; break
            
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
                header_idx = i; break
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
st.sidebar.info("📌 当前为只读模式，网页仅读取并统计，绝对不会修改您的原文件。")
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
    display_df = df_current.astype(str).replace({' 00:00:00': ''}, regex=True).replace({'nan': '', 'None': ''})
    st.dataframe(display_df, use_container_width=True, height=350)

    # ================= 6. 核心统计算法库 =================
    def parse_class_string(val_str):
        """最强容错提取算法：去除所有空格，提取课时倍数，精准拆分教师和课程"""
        val_str = str(val_str).replace(" ", "") # 抹除所有导致失效的内层空格
        
        # 排除无用词汇
        ignore = ['0', '0.0', 'nan', 'none', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日', '体育', '班会', '国学', '美术', '音乐', '大扫除']
        if not val_str or val_str.lower() in ignore or re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', val_str):
            return None
            
        # 智能提取尾部数字（例如：早自2，代表两节课）
        count = 1.0
        m_num = re.search(r'(\d+(?:\.\d+)?)$', val_str)
        if m_num:
            if m_num.start() == 0: return None # 全是数字则跳过
            count = float(m_num.group(1))
            val_str = val_str[:m_num.start()] # 去掉数字，只留名字
            
        # 核心拆分法则 1：认准“高、初、小”
        match = re.match(r'^(.*?)(高[一二三]|初[一二三]|小[一二三四五六])(.*)$', val_str)
        if match:
            return {'教师姓名': match.group(1), '课程类别': match.group(2) + match.group(3), '课时数': count}
            
        # 核心拆分法则 2：强制匹配常用课名
        known_types = ['早自', '正大', '正小', '晚自', '自大', '自小', '辅导', '正课', '早读', '晚修']
        for kt in known_types:
            if val_str.endswith(kt):
                return {'教师姓名': val_str[:-len(kt)], '课程类别': kt, '课时数': count}
                
        # 最后的兜底：如果完全无法识别，将整个名字记下来，归类为“其他课”
        if len(val_str) >= 2:
            return {'教师姓名': val_str, '课程类别': '常规课', '课时数': count}
        return None

    # ================= 7. 双模式统计区 =================
    st.markdown("---")
    tab1, tab2 = st.tabs(["📏 精准双重过滤统计 (锁定范围 + 自由日历)", "📊 常规清单表统计 (手动选列)"])
    
    with tab1:
        st.info("💡 请先锁定排课表横向段，然后您可以通过日历自由设定时间（日历不再限制日期）。")
        all_cols = display_df.columns.tolist()
        
        # 极其贴心：为下拉菜单里的列名拼上日期（如果找得到的话）
        display_options = []
        for col in all_cols:
            date_info = ""
            for i in range(min(4, len(display_df))):
                val = str(display_df[col].iloc[i]).strip()
                m = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', val)
                if m: date_info = m.group(1); break
            display_options.append(f"{col} 📅 {date_info}" if date_info else col)

        # 1. 结构锁定
        col_a, col_b = st.columns(2)
        with col_a:
            default_start = 14 if len(all_cols) > 14 else 0
            start_choice = st.selectbox("🚩 第一步：选择【起始】列", options=display_options, index=default_start)
        with col_b:
            default_end = 20 if len(all_cols) > 20 else len(all_cols) - 1
            end_choice = st.selectbox("🏁 第二步：选择【结束】列", options=display_options, index=default_end)
            
        start_idx, end_idx = display_options.index(start_choice), display_options.index(end_choice)
        
        if start_idx > end_idx:
            st.error("⚠️ 起始列不能在结束列的后面！")
        else:
            locked_cols = all_cols[start_idx : end_idx + 1]
            
            # 拿到锁定区域内含有的日期
            col_dates = {}
            for col in locked_cols:
                for i in range(min(3, len(display_df))):
                    val = str(display_df[col].iloc[i]).strip()
                    match = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', val)
                    if match:
                        try:
                            col_dates[col] = pd.to_datetime(match.group(1)).date()
                            break
                        except: pass
            
            # 2. 自由日历（去掉了强制封印）
            if col_dates:
                st.markdown("##### 🗓️ 第三步：设定需要统计的时间")
                min_d, max_d = min(col_dates.values()), max(col_dates.values())
                
                # 解除了 min_value 和 max_value 的限制，让你自由选择
                date_range = st.date_input("选择时间范围（系统默认选中本段全部时间）", [min_d, max_d])
                
                if len(date_range) >= 1:
                    filter_start = date_range[0]
                    filter_end = date_range[1] if len(date_range) == 2 else date_range[0]
                    
                    final_target_cols = [c for c, d in col_dates.items() if filter_start <= d <= filter_end]
                    
                    if not final_target_cols:
                        st.warning("⚠️ 在你选择的时间范围内，指定的列中没有排课数据哦。")
                    else:
                        st.success(f"✅ 将对以下日期的列进行统计：**{', '.join([str(col_dates[c]) for c in final_target_cols])}**")
                        
                        if st.button("🚀 极速拆分并生成统计报表", type="primary"):
                            records = []
                            for col in final_target_cols:
                                for val in display_df[col]:
                                    parsed = parse_class_string(val)
                                    if parsed:
                                        parsed['来源日期'] = str(col_dates[col]) # 记录一下是哪天的，方便排错
                                        parsed['原始单元格'] = val
                                        records.append(parsed)
                                        
                            if records:
                                stat_df = pd.DataFrame(records)
                                pivot_df = pd.pivot_table(stat_df, values='课时数', index='教师姓名', columns='课程类别', aggfunc='sum', fill_value=0)
                                pivot_df['总计'] = pivot_df.sum(axis=1)
                                
                                st.success(f"🎉 统计完毕！共计提取到 {stat_df['课时数'].sum()} 节课时。")
                                st.dataframe(pivot_df, use_container_width=True)
                                
                                # 🔍 【新增防漏抓透视镜】：让你明白系统到底抓了什么
                                with st.expander("🔍 觉得算得不准？点这里查看抓取明细 (Debug)"):
                                    st.write("系统从你的 Excel 中拆解出了以下记录，如有遗漏说明你的 Excel 拼写不符合规则：")
                                    st.dataframe(stat_df)
                            else:
                                st.warning("没有找到可以统计的课时数据。")
            else:
                st.warning("⚠️ 在你锁定的列范围中，没有找到 YYYY-MM-DD 格式的日期。")

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