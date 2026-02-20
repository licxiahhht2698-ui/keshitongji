import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Font, Alignment

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
    div[data-testid="stDownloadButton"] > button {
        background-color: #ffe699 !important; border-color: #ffc000 !important; font-weight: bold;
    }
    div[data-testid="stDownloadButton"] > button:hover { background-color: #ffd966 !important; }
</style>
""", unsafe_allow_html=True)

st.title("📚 教师排课智能读取与精准统计系统")

if 'all_sheets' not in st.session_state: st.session_state['all_sheets'] = None
if 'current_sheet' not in st.session_state: st.session_state['current_sheet'] = None
if 'global_mode' not in st.session_state: st.session_state['global_mode'] = False

# ================= 新增：带大表头的 Excel 导出引擎 =================
def convert_df_to_excel(df, sheet_name, title):
    """生成带有大字号合并表头的正式 Excel 报表"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 从第3行开始写数据，把前2行留给大表头
        df.to_excel(writer, sheet_name=sheet_name, startrow=2)
        worksheet = writer.sheets[sheet_name]
        
        # 写入大标题
        cell = worksheet.cell(row=1, column=1, value=title)
        cell.font = Font(size=16, bold=True, color="000000")
        
        # 合并单元格让标题居中
        max_col = len(df.columns) + 1 
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
    return output.getvalue()

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

# ================= 6. 核心统计算法库 =================
def parse_class_string(val_str):
    val_str = str(val_str).replace(" ", "") 
    ignore = ['0', '0.0', 'nan', 'none', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日', '体育', '班会', '国学', '美术', '音乐', '大扫除']
    if not val_str or val_str.lower() in ignore or re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', val_str) or re.search(r'^第[一二三四五六七八九十]+周', val_str):
        return None
        
    count = 1.0
    m_num = re.search(r'(\d+(?:\.\d+)?)$', val_str)
    if m_num:
        if m_num.start() == 0: return None
        count = float(m_num.group(1))
        val_str = val_str[:m_num.start()] 
        
    match = re.match(r'^([\u4e00-\u9fa5a-zA-Z]+?)(高[一二三]|初[一二三]|小[一二三四五六])(.*)$', val_str)
    if match: return {'教师姓名': match.group(1), '课程类别': match.group(2) + match.group(3), '课时数': count}
        
    known_types = ['早自', '正大', '正小', '晚自', '自大', '自小', '辅导', '正课', '早读', '晚修']
    for kt in known_types:
        if val_str.endswith(kt): return {'教师姓名': val_str[:-len(kt)], '课程类别': kt, '课时数': count}
            
    if len(val_str) >= 2: return {'教师姓名': val_str, '课程类别': '常规课', '课时数': count}
    return None

# ================= 3. 侧边栏与全局汇总配置 =================
st.sidebar.header("📁 数据中心")
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

# 【新增核心功能】：侧边栏全校一键汇总
if st.session_state['all_sheets'] is not None:
    st.sidebar.markdown("---")
    st.sidebar.subheader("🌐 全校整体统计设置")
    st.sidebar.info("统一划定所有班级的排课列与时间段，一键生成全校总表。")
    
    col_g1, col_g2 = st.sidebar.columns(2)
    with col_g1: g_start_idx = st.number_input("起始列数字", min_value=1, value=15, help="第15列即 未命名_15")
    with col_g2: g_end_idx = st.number_input("结束列数字", min_value=1, value=21, help="第21列即 未命名_21")
    
    g_dates = st.sidebar.date_input("全校统计时间段", [])
    
    if st.sidebar.button("🚀 一键生成全校总表", use_container_width=True, type="primary"):
        if len(g_dates) < 1:
            st.sidebar.error("请先选择完整的时间段！")
        else:
            st.session_state['global_mode'] = True
            st.session_state['g_start'] = g_start_idx
            st.session_state['g_end'] = g_end_idx
            st.session_state['g_dates'] = g_dates

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
                    # 点击任何单一班级，都会退出全校总表模式
                    st.session_state['global_mode'] = False 
    st.markdown("<hr style='margin: 5px 0px;'>", unsafe_allow_html=True)

    # ================= 5. 分支判断：显示全校总表 还是 单个班级表 =================
    if st.session_state['global_mode']:
        # ---------------- 全局统计视图 ----------------
        g_dates = st.session_state['g_dates']
        f_start = g_dates[0]
        f_end = g_dates[1] if len(g_dates) == 2 else g_dates[0]
        
        st.markdown(f"### 🌐 全校教师课时总汇 📅 【{f_start} 至 {f_end}】")
        st.info(f"系统正在扫描所有班级表格的第 {st.session_state['g_start']} 列 到 第 {st.session_state['g_end']} 列...")
        
        all_records = []
        # 要跳过的汇总表（防止数据重复统计）
        skip_words = ['总表', '分表', '汇总'] 
        
        for s_name, s_df in st.session_state['all_sheets'].items():
            if any(kw in s_name for kw in skip_words): continue
            
            # 安全切片：把输入的数字列转成代码索引 (比如15列是索引14)
            start_i = max(0, st.session_state['g_start'] - 1)
            end_i = min(len(s_df.columns), st.session_state['g_end'])
            if start_i >= end_i: continue
                
            locked_cols = s_df.columns[start_i:end_i]
            
            for col in locked_cols:
                current_date = None
                for val in s_df[col]:
                    val_str = str(val).strip()
                    m = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', val_str)
                    if m:
                        try: current_date = pd.to_datetime(m.group(1)).date()
                        except: pass
                        continue
                    
                    if current_date and (f_start <= current_date <= f_end):
                        parsed = parse_class_string(val_str)
                        if parsed:
                            parsed['来源班级'] = s_name
                            parsed['来源日期'] = str(current_date)
                            all_records.append(parsed)
                            
        if all_records:
            stat_df = pd.DataFrame(all_records)
            # 全校透视：算出每个老师所有科目的总计
            pivot_df = pd.pivot_table(stat_df, values='课时数', index='教师姓名', columns='课程类别', aggfunc='sum', fill_value=0)
            pivot_df['总计'] = pivot_df.sum(axis=1)
            
            st.success(f"🎉 扫描完毕！全校共有 {len(stat_df['教师姓名'].unique())} 位老师在此期间上了课，总计 {stat_df['课时数'].sum()} 节。")
            st.dataframe(pivot_df, use_container_width=True)
            
            title = f"【全校汇总】课时统计报表 ({f_start} 至 {f_end})"
            excel_data = convert_df_to_excel(pivot_df, sheet_name="全校总计", title=title)
            st.download_button(
                label="⬇️ 导出《全校课时总表》为 Excel",
                data=excel_data,
                file_name=f"全校课时总表_{f_start}至{f_end}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            with st.expander("🔍 查看全校抓取底层明细 (用于排错)"):
                st.dataframe(stat_df)
        else:
            st.warning("⚠️ 在指定的日期和列范围内，全校表格中均未抓取到有效课时数据，请检查侧边栏设置！")
            
    else:
        # ---------------- 单一班级视图 (你原来用的界面) ----------------
        current = st.session_state['current_sheet']
        st.markdown(f"#### 👁️ 当前查看 : 【 {current} 】")
        
        df_current = st.session_state['all_sheets'][current].copy()
        display_df = df_current.astype(str).replace({' 00:00:00': ''}, regex=True).replace({'nan': '', 'None': ''})
        st.dataframe(display_df, use_container_width=True, height=350)

        st.markdown("---")
        tab1, tab2 = st.tabs(["📏 【周课表专用】垂直穿插统计", "📊 【常规明细表】手动选列统计"])
        
        with tab1:
            all_cols = display_df.columns.tolist()
            col_a, col_b = st.columns(2)
            with col_a:
                default_start = 14 if len(all_cols) > 14 else 0
                start_choice = st.selectbox("🚩 起始列 (星期一)", options=all_cols, index=default_start)
            with col_b:
                default_end = 20 if len(all_cols) > 20 else len(all_cols) - 1
                end_choice = st.selectbox("🏁 结束列 (星期日)", options=all_cols, index=default_end)
                
            start_idx, end_idx = all_cols.index(start_choice), all_cols.index(end_choice)
            
            if start_idx > end_idx:
                st.error("⚠️ 起始列不能在结束列的后面！")
            else:
                locked_cols = all_cols[start_idx : end_idx + 1]
                all_dates_in_range = set()
                for col in locked_cols:
                    for val in display_df[col]:
                        val_str = str(val).strip()
                        m = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', val_str)
                        if m:
                            try: all_dates_in_range.add(pd.to_datetime(m.group(1)).date())
                            except: pass
                
                if all_dates_in_range:
                    min_d, max_d = min(all_dates_in_range), max(all_dates_in_range)
                    date_range = st.date_input(f"🗓️ 该班级区域共扫描到 {len(all_dates_in_range)} 天的数据，请划定提取区间：", [min_d, max_d])
                    
                    if len(date_range) >= 1:
                        filter_start = date_range[0]
                        filter_end = date_range[1] if len(date_range) == 2 else date_range[0]
                        
                        if st.button("🚀 开始本班垂直扫描提取", type="primary"):
                            records = []
                            for col in locked_cols:
                                current_date = None
                                for val in display_df[col]:
                                    val_str = str(val).strip()
                                    m = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', val_str)
                                    if m:
                                        try: current_date = pd.to_datetime(m.group(1)).date()
                                        except: pass
                                        continue
                                    
                                    if current_date and (filter_start <= current_date <= filter_end):
                                        parsed = parse_class_string(val_str)
                                        if parsed:
                                            parsed['来源日期'] = str(current_date)
                                            parsed['原始录入'] = val_str
                                            records.append(parsed)
                                            
                            if records:
                                stat_df = pd.DataFrame(records)
                                pivot_df = pd.pivot_table(stat_df, values='课时数', index='教师姓名', columns='课程类别', aggfunc='sum', fill_value=0)
                                pivot_df['总计'] = pivot_df.sum(axis=1)
                                
                                st.success(f"🎉 统计完毕！【{current}】共计 {stat_df['课时数'].sum()} 节课时。")
                                st.dataframe(pivot_df, use_container_width=True)
                                
                                # 【核心功能：大表头导出】
                                formal_title = f"【{current}】课时统计报表 ({filter_start} 至 {filter_end})"
                                excel_data = convert_df_to_excel(pivot_df, sheet_name=current, title=formal_title)
                                
                                st.download_button(
                                    label=f"⬇️ 导出带表头的《{current}报表》为 Excel",
                                    data=excel_data,
                                    file_name=f"{current}_课时统计_{filter_start}至{filter_end}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                                
                                with st.expander("🔍 点这里查看提取明细账单"):
                                    st.dataframe(stat_df)
                            else:
                                st.warning("未找到可识别的课时。")
                else:
                    st.warning("⚠️ 没有扫描到包含日期的行！")

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
                    formal_title = f"【{current}】常规课时统计"
                    excel_data = convert_df_to_excel(pivot_df, sheet_name=current, title=formal_title)
                    st.download_button(
                        label="⬇️ 导出当前常规报表",
                        data=excel_data,
                        file_name=f"{current}_常规课时.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except:
                    st.warning("无法生成，请确认选对了列名哦！")
else:
    st.info("👆 请先在左侧上传您的 Excel 文件！")