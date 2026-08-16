# ============================================================
# --- Web 介面佈局 ---
# ============================================================

st.title("🏫 澳森托嬰中心 智慧教案與適性月計畫系統")

st.markdown(
    "完成每月教案輪值排班後，可一鍵下載 "
    "**【月教案輪值表】**與**【適性發展活動月計畫】**。"
)


# ============================================================
# 壹、基本設定
# ============================================================

st.subheader("壹、基本設定")

# 園所、年份、月份
c1, c2, c3 = st.columns([1.5, 1, 1])

with c1:
    branch = st.selectbox(
        "園所：",
        ["澳森", "澳森文德"]
    )

with c2:
    year_roc = st.number_input(
        "民國年份：",
        min_value=114,
        max_value=125,
        value=115
    )

with c3:
    month = st.selectbox(
        "月份：",
        list(range(1, 13)),
        index=8
    )


# 主帶老師
teachers_input = st.text_input(
    "主帶老師名單（請以逗號分隔）：",
    value="",
    placeholder="例如：Panda, Candy, 綺綺, 樺樺, 小安, 均宜"
)


# ------------------------------------------------------------
# 使用說明：放在「壹、基本設定」裡
# ------------------------------------------------------------

st.markdown("#### 📖 使用說明")

st.markdown(
    """
1. 選取園所、年份及月份  
2. 輸入該月設計教案的人員  
3. 填寫輪值排班表  
4. 完成後檢查，即可一鍵下載輪值表及月計畫
"""
)


# ------------------------------------------------------------
# 老師／特殊活動下拉選單
# ------------------------------------------------------------

builtin_options = [
    "（請選擇）",
    "主任",
    "特約醫師",
    "牙醫師",
    "消防演練",
    "大型活動",
    "國定假日",
    "教案暫停"
]

raw_teachers = [
    t.strip()
    for t in teachers_input.split(",")
    if t.strip()
]

dropdown_options = (
    builtin_options[:2]
    + raw_teachers
    + builtin_options[2:]
)


# ------------------------------------------------------------
# 依分園切換課程位置
# 澳森：週三觸覺、週四體能
# 澳森文德：週三體能、週四觸覺
# ------------------------------------------------------------

if branch == "澳森":

    headers = [
        "主題",
        "星期一：繪本",
        "星期二：小肌肉／認知",
        "星期三：小肌肉／觸覺",
        "星期四：體能課",
        "星期五：藝術創作"
    ]

    wed_title_prefix = "週三：小肌肉／觸覺"
    thu_title_prefix = "週四：體能"

else:

    headers = [
        "主題",
        "星期一：繪本",
        "星期二：小肌肉／認知",
        "星期三：體能課",
        "星期四：小肌肉／觸覺",
        "星期五：藝術創作"
    ]

    wed_title_prefix = "週三：體能"
    thu_title_prefix = "週四：小肌肉／觸覺"


# ============================================================
# 貳、輪值排班
# ============================================================

st.divider()

st.subheader(
    f"貳、{branch} {year_roc} 年 {month} 月輪值排班"
)


# 西元年份
year_ad = year_roc + 1911

# 取得月曆
cal = calendar.monthcalendar(year_ad, month)


# ------------------------------------------------------------
# 只保留有週一～週五日期的週次
# ------------------------------------------------------------

work_weeks = []

for week in cal:

    mon_to_fri = week[0:5]

    if any(day > 0 for day in mon_to_fri):
        work_weeks.append(mon_to_fri)


# ------------------------------------------------------------
# 資料容器
# ------------------------------------------------------------

table_data = []
weekly_schedule_for_plan = []


# ------------------------------------------------------------
# 每週輪值表
# ------------------------------------------------------------

for idx, week in enumerate(work_weeks, start=1):

    with st.expander(
        f"📌 第 {idx} 週排班設定",
        expanded=True
    ):

        col_theme, col_m, col_t, col_w, col_th, col_f = st.columns(
            [1.2, 2.2, 1.3, 1.3, 1.3, 1.3]
        )


        # 日期
        d_mon = f"{month}/{week[0]}" if week[0] > 0 else ""
        d_tue = f"{month}/{week[1]}" if week[1] > 0 else ""
        d_wed = f"{month}/{week[2]}" if week[2] > 0 else ""
        d_thu = f"{month}/{week[3]}" if week[3] > 0 else ""
        d_fri = f"{month}/{week[4]}" if week[4] > 0 else ""


        # ----------------------------------------------------
        # 第一週保留範例
        # ----------------------------------------------------

        if idx == 1:

            theme_default = "顏色"
            book_name_default = "小藍與小黃"
            m_lead_default_idx = 1

        else:

            theme_default = ""
            book_name_default = ""
            m_lead_default_idx = 0


        # ----------------------------------------------------
        # 主題
        # ----------------------------------------------------

        with col_theme:

            theme = st.text_input(
                "週主題",
                key=f"theme_{idx}",
                value=theme_default,
                placeholder="請輸入主題"
            )


        # ----------------------------------------------------
        # 週一：繪本
        # ----------------------------------------------------

        with col_m:

            if d_mon:
                st.markdown(
                    f"**週一：繪本 ({d_mon})**"
                )
            else:
                st.markdown("**週一：繪本**")


            m_lead = st.selectbox(
                "導讀人",
                dropdown_options,
                key=f"m_lead_{idx}",
                index=m_lead_default_idx
            )


            book_name = st.text_input(
                "繪本名稱",
                key=f"book_{idx}",
                value=book_name_default,
                placeholder="請輸入繪本名稱"
            )


            lead_txt = (
                ""
                if m_lead == "（請選擇）"
                else f" {m_lead}"
            )

            book_txt = (
                f"\n{book_name}"
                if book_name.strip()
                else ""
            )

            mon_str = (
                f"{d_mon}{lead_txt}{book_txt}".strip()
                if d_mon
                else ""
            )


        # ----------------------------------------------------
        # 週二：小肌肉／認知
        # ----------------------------------------------------

        with col_t:

            if d_tue:
                st.markdown(
                    f"**週二：小肌肉／認知 ({d_tue})**"
                )
            else:
                st.markdown(
                    "**週二：小肌肉／認知**"
                )


            tue_teacher = st.selectbox(
                "負責人／狀態",
                dropdown_options,
                key=f"tue_{idx}",
                index=0
            )


            tue_txt = (
                ""
                if tue_teacher == "（請選擇）"
                else f" {tue_teacher}"
            )

            tue_str = (
                f"{d_tue}{tue_txt}".strip()
                if d_tue
                else ""
            )


        # ----------------------------------------------------
        # 週三
        # ----------------------------------------------------

        with col_w:

            if d_wed:
                st.markdown(
                    f"**{wed_title_prefix} ({d_wed})**"
                )
            else:
                st.markdown(
                    f"**{wed_title_prefix}**"
                )


            wed_teacher = st.selectbox(
                "負責人／狀態",
                dropdown_options,
                key=f"wed_{idx}",
                index=0
            )


            wed_txt = (
                ""
                if wed_teacher == "（請選擇）"
                else f" {wed_teacher}"
            )

            wed_str = (
                f"{d_wed}{wed_txt}".strip()
                if d_wed
                else ""
            )


        # ----------------------------------------------------
        # 週四
        # ----------------------------------------------------

        with col_th:

            if d_thu:
                st.markdown(
                    f"**{thu_title_prefix} ({d_thu})**"
                )
            else:
                st.markdown(
                    f"**{thu_title_prefix}**"
                )


            thu_teacher = st.selectbox(
                "負責人／狀態",
                dropdown_options,
                key=f"thu_{idx}",
                index=0
            )


            thu_txt = (
                ""
                if thu_teacher == "（請選擇）"
                else f" {thu_teacher}"
            )

            thu_str = (
                f"{d_thu}{thu_txt}".strip()
                if d_thu
                else ""
            )


        # ----------------------------------------------------
        # 週五：藝術創作
        # ----------------------------------------------------

        with col_f:

            if d_fri:
                st.markdown(
                    f"**週五：藝術創作 ({d_fri})**"
                )
            else:
                st.markdown(
                    "**週五：藝術創作**"
                )


            fri_teacher = st.selectbox(
                "負責人／狀態",
                dropdown_options,
                key=f"fri_{idx}",
                index=0
            )


            fri_txt = (
                ""
                if fri_teacher == "（請選擇）"
                else f" {fri_teacher}"
            )

            fri_str = (
                f"{d_fri}{fri_txt}".strip()
                if d_fri
                else ""
            )


        # ----------------------------------------------------
        # 輪值表資料
        # ----------------------------------------------------

        table_data.append(
            [
                theme,
                mon_str,
                tue_str,
                wed_str,
                thu_str,
                fri_str
            ]
        )


        # ----------------------------------------------------
        # 月計畫使用資料
        # ----------------------------------------------------

        valid_dates = [
            day
            for day in week
            if day > 0
        ]

        if valid_dates:

            start_day = valid_dates[0]
            end_day = valid_dates[-1]

            week_range_str = (
                f"{month}/{start_day}～{month}/{end_day}"
            )

        else:

            week_range_str = ""


        roster_names = [
            m_lead,
            tue_teacher,
            wed_teacher,
            thu_teacher,
            fri_teacher
        ]

        roster_names = [
            x
            for x in roster_names
            if x and x != "（請選擇）"
        ]

        roster_summary = "、".join(
            dict.fromkeys(roster_names)
        )


        weekly_schedule_for_plan.append(
            {
                "week_range": week_range_str,
                "theme": (
                    theme
                    if theme.strip()
                    else "主題探索"
                ),
                "book": (
                    book_name
                    if book_name.strip()
                    else "主題繪本導讀"
                ),
                "roster_summary": roster_summary
            }
        )


# ============================================================
# 參、成果匯出
# ============================================================

st.divider()

st.subheader("參、成果匯出")


# ============================================================
# 1. 輪值表下載
# ============================================================

st.markdown(
    f"### 📋 1. 下載【{branch}月教案輪值表】"
)

st.caption(
    "請先確認上方主題、繪本及每日負責人皆填寫正確。"
)


doc_roster_bytes = generate_roster_docx(
    branch,
    year_roc,
    month,
    teachers_input,
    table_data,
    headers
)


st.download_button(
    label=(
        f"📥 下載【{branch} "
        f"{year_roc}年{month}月教案輪值表】"
    ),
    data=doc_roster_bytes,
    file_name=(
        f"{branch}_"
        f"{year_roc}年{month}月教案輪值表.docx"
    ),
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "wordprocessingml.document"
    ),
    use_container_width=True
)


# ============================================================
# 2. 適性月計畫
# 月齡選單移到這裡
# ============================================================

st.markdown("---")

st.markdown(
    f"### 🌱 2. 下載【{branch}適性發展活動月計畫】"
)


# ------------------------------------------------------------
# 依分園決定月齡
# ------------------------------------------------------------

if branch == "澳森":

    available_ages = [
        "0-12個月",
        "12-24個月",
        "24-36個月"
    ]

else:

    available_ages = [
        "7-12個月",
        "12-24個月",
        "24-36個月"
    ]


# 月齡選單現在位於月計畫下載按鈕正上方
age_group = st.selectbox(
    "月計畫年齡階段：",
    available_ages,
    key="month_plan_age_group"
)


st.caption(
    "選擇月齡後，系統會依該年齡階段套用對應的"
    "適性發展指標及活動內容。"
)


# ------------------------------------------------------------
# 產生月計畫
# ------------------------------------------------------------

doc_plan_bytes = generate_month_plan_docx(
    branch,
    year_roc,
    month,
    age_group,
    teachers_input,
    weekly_schedule_for_plan
)


st.download_button(
    label=(
        f"📥 下載【{branch}（{age_group}）"
        f"適性月計畫】"
    ),
    data=doc_plan_bytes,
    file_name=(
        f"{branch}_"
        f"{year_roc}年{month}月_"
        f"{age_group}_適性發展活動月計畫.docx"
    ),
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "wordprocessingml.document"
    ),
    use_container_width=True
)
