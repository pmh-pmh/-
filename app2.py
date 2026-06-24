import streamlit as st
import pandas as pd
import plotly.express as px
import re

# ==================== 页面配置 ====================
st.set_page_config(page_title="省队运动员成绩系统", layout="wide")
st.title("🏃 省队运动员成绩查询系统")

# ==================== 标准项目列表 ====================
STANDARD_EVENTS = [
    "100米", "200米", "400米", "800米", "1500米", "5000米", "10000米",
    "110米栏", "400米栏", "3000米障碍",
    "跳高", "撑竿跳高", "跳远", "三级跳远",
    "铅球", "铁饼", "链球", "标枪",
    "十项全能", "20公里竞走", "马拉松",
    "4×100米接力", "4×400米接力",
    "100米栏", "七项全能",
    "4×100米混合接力", "4×400米混合接力", "竞走混合团体"
]
STANDARD_EVENTS = list(set(STANDARD_EVENTS))
STANDARD_EVENTS.sort(key=len, reverse=True)

# ==================== 成绩解析函数 ====================
def parse_score_to_seconds(score_str):
    if pd.isna(score_str):
        return None
    if not isinstance(score_str, str):
        try:
            return float(score_str)
        except:
            return None
    s = score_str.strip()
    if '(' in s:
        s = s.split('(')[0].strip()
    if not s:
        return None
    if ':' in s:
        parts = s.split(':')
        if len(parts) == 2:
            try:
                minutes = float(parts[0])
                seconds = float(parts[1])
                return minutes * 60 + seconds
            except:
                return None
        else:
            return None
    else:
        try:
            return float(s)
        except:
            return None

def extract_raw_score(score_str):
    if pd.isna(score_str):
        return ''
    if not isinstance(score_str, str):
        return str(score_str)
    s = score_str.strip()
    if '(' in s:
        s = s.split('(')[0].strip()
    return s

# ==================== 核心项目归类函数 ====================
def map_to_standard_event(orig):
    if not isinstance(orig, str):
        return str(orig)
    s = orig.strip()
    if '60米' in s and '栏' not in s and '接力' not in s:
        return "60米"
    if '异程接力' in s:
        return "异程接力"
    prefixes = ['男子', '女子', '甲组', '乙组', 'U18', 'U20', 'U16', 'U14', '少年', '青年', '成年', '混合']
    for p in prefixes:
        if s.startswith(p):
            s = s[len(p):].strip()
            break
    if '混合' in orig and ('接力' in orig):
        if '4×100' in orig or '4x100' in orig:
            return "4×100米混合接力"
        elif '4×400' in orig or '4x400' in orig:
            return "4×400米混合接力"
        elif '竞走' in orig:
            return "竞走混合团体"
    s = s.replace('x', '×')
    for event in STANDARD_EVENTS:
        if event in s:
            return event
    suffixes = ['预赛', '决赛', '半决赛', '复赛', '资格赛', '第一轮', '第二轮', '第1轮', '第2轮']
    for suf in suffixes:
        if s.endswith(suf):
            s_clean = s[:-len(suf)].strip()
            for event in STANDARD_EVENTS:
                if event in s_clean:
                    return event
            break
    return orig.strip()

# ==================== 数据加载函数 ====================
@st.cache_data
def load_data():
    df = pd.read_excel('浙江省成绩汇总.xlsx', sheet_name='表1', engine='openpyxl')
    df.columns = df.columns.str.strip()
    def clean_reg_id(val):
        if pd.isna(val):
            return ''
        if isinstance(val, (int, float)):
            return str(int(val))
        return str(val).strip()
    if '注册号' in df.columns:
        df['注册号'] = df['注册号'].apply(clean_reg_id)
    else:
        st.error("Excel 中找不到 '注册号' 列，请检查表头名称是否完全一致！")
        st.stop()
    df['成绩_原始'] = df['成绩'].apply(extract_raw_score)
    df['成绩_数值'] = df['成绩'].apply(parse_score_to_seconds)
    df['项目_核心'] = df['项目'].apply(map_to_standard_event)
    if '比赛时间' in df.columns:
        df['比赛时间'] = pd.to_datetime(df['比赛时间'], errors='coerce')
    else:
        st.error("Excel 中找不到 '比赛时间' 列，请检查表头名称！")
        st.stop()
    df = df.sort_values(by='比赛时间')
    return df

df = load_data()

# ==================== 侧边栏搜索 ====================
st.sidebar.header("🔍 筛选条件")
name_input = st.sidebar.text_input("请输入运动员姓名", value="")

if name_input:
    candidate_df = df[df['姓名'].str.contains(name_input, na=False)]
    if not candidate_df.empty:
        candidate_df = candidate_df[candidate_df['注册号'] != '']
        if candidate_df.empty:
            st.sidebar.warning("该运动员的成绩记录中注册号为空，无法识别")
            st.stop()
        candidates_unique = candidate_df[['姓名', '注册号', '代表队']].drop_duplicates(subset=['注册号'])
        candidates_unique['注册号'] = candidates_unique['注册号'].astype(str)
        if len(candidates_unique) == 1:
            selected_reg_id = candidates_unique.iloc[0]['注册号']
            st.sidebar.success(f"当前查看: {candidates_unique.iloc[0]['姓名']} ({candidates_unique.iloc[0]['代表队']})")
        else:
            candidates_unique['label'] = candidates_unique['姓名'] + " (注册号: " + candidates_unique['注册号'] + ", " + candidates_unique['代表队'] + ")"
            options_list = candidates_unique['注册号'].tolist()
            label_map = {str(k): str(v) for k, v in zip(candidates_unique['注册号'], candidates_unique['label'])}
            selected_reg_id = st.sidebar.selectbox(
                "检测到同名/多人，请选择具体的运动员：",
                options=options_list,
                format_func=lambda x: label_map.get(str(x), str(x))
            )
            current_info = candidates_unique[candidates_unique['注册号'] == selected_reg_id]
            if not current_info.empty:
                st.sidebar.success(f"当前查看: {current_info.iloc[0]['姓名']} ({current_info.iloc[0]['代表队']})")
        
        athlete_data = df[df['注册号'] == selected_reg_id]
        if not athlete_data.empty:
            st.subheader(f"📈 {athlete_data['姓名'].iloc[0]} 历史成绩趋势（按标准项目归类）")
            st.caption("💡 点击右侧图例切换项目，拖拽底部滑动条查看不同时间段")
            
            plot_data = athlete_data[athlete_data['成绩_数值'].notna()].copy()
            if not plot_data.empty:
                fig = px.line(
                    plot_data,
                    x='比赛时间',
                    y='成绩_数值',
                    color='项目_核心',
                    markers=True,
                    text='成绩_原始',
                    hover_data={
                        '赛事名称': True,
                        '赛次': True,
                        '比赛时间': True,
                        '项目': True,
                        '项目_核心': False
                    },
                    title="成绩随时间变化趋势"
                )
                
                # 默认隐藏所有曲线
                fig.for_each_trace(lambda t: t.update(visible='legendonly'))
                fig.update_traces(textposition="top center")
                
                # ========== 关键修改：锁定 X 轴范围（禁止鼠标拖拽） ==========
                fig.update_xaxes(
                    type='date',
                    tickformat='%Y-%m',
                    fixedrange=True,                     # 禁止拖拽平移/缩放 X 轴
                    rangeslider=dict(
                        visible=True,
                        thickness=0.08,
                        bgcolor="rgba(220,220,220,0.3)",
                        bordercolor="lightgray",
                        borderwidth=1,
                        yaxis=dict(
                            range=[0, 1],
                            rangemode='fixed'
                        )
                    ),
                    rangeselector_visible=False,
                    title_text="比赛日期"
                )
                
                # 锁定 Y 轴范围（原有）
                fig.update_yaxes(fixedrange=True)
                
                fig.update_layout(
                    yaxis_title="成绩 (秒/米/分秒)",
                    legend_title="标准项目 (点击显示/隐藏)",
                    legend=dict(
                        orientation="v",
                        yanchor="top",
                        y=1,
                        xanchor="left",
                        x=1.02
                    ),
                    height=500
                )
                
                # 隐藏模式栏（右上角工具条）
                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
            else:
                st.warning("该运动员暂无有效数值成绩，无法绘制折线图")
            
            with st.expander("📋 查看该运动员详细数据表格"):
                display_cols = ['赛事名称', '项目', '项目_核心', '赛次', '成绩_原始', '风速', '比赛时间', '赛季年份']
                existing_cols = [c for c in display_cols if c in athlete_data.columns]
                st.dataframe(athlete_data[existing_cols], use_container_width=True)
                
                all_core = athlete_data['项目_核心'].unique()
                all_orig = athlete_data['项目'].unique()
                unmatched = [p for p in all_core if p in all_orig]
                if unmatched:
                    st.info(f"💡 发现以下项目未匹配到标准列表，如需归类请调整 `map_to_standard_event` 函数：{', '.join(unmatched)}")
        else:
            st.warning("未找到该注册号对应的成绩数据")
    else:
        st.sidebar.warning("未找到该运动员，请检查姓名输入")
else:
    st.sidebar.info("请输入姓名开始查询")