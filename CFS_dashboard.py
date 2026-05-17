import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="Tdh Kenya | CFS Data Dashboard",
    layout="wide",
    page_icon="📊"
)

st.markdown("""
<style>
.stMetric {
    background-color: #ffffff;
    padding: 15px;
    border-radius: 10px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    border: 1px solid #f0f2f6;
}
thead tr th {
    background-color: #0083B8 !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)


def clean_yes_no_series(series):
    values = series.astype("string").str.strip().str.lower()
    return values.isin(["yes", "y", "true", "1", "1.0", "selected"])


def get_sorted_options(data, column):
    return sorted(data[column].dropna().unique())


def show_section_conclusion(section_title, data_summary):
    prompt = f"""
    You are reviewing a child protection dashboard for programme decision-making.

    Write a brief conclusion from the data below.
    Use 2 to 3 sentences only.
    Mention the most important pattern, imbalance, or operational implication.
    Do not repeat every number.
    Keep the wording simple and useful for CFS programme staff.

    Section: {section_title}

    Data summary:
    {data_summary}
    """

    conclusion = None

    try:
        api_key = st.secrets.get("OPENAI_API_KEY")
        model = st.secrets.get("OPENAI_MODEL", "gpt-4o-mini")

        if api_key:
            from openai import OpenAI

            client = OpenAI(api_key=api_key)

            response = client.chat.completions.create(
                model=model,
                messages=[
                    {
                        "role": "system",
                        "content": "You generate concise conclusions from child protection dashboard data."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                temperature=0.2,
                max_tokens=140
            )

            conclusion = response.choices[0].message.content.strip()

    except Exception:
        conclusion = None

    if not conclusion:
        conclusion = (
            f"Conclusion: The {section_title.lower()} section should be reviewed using the highest counts, "
            "gender differences, and site-level variations shown above. These patterns can guide follow-up, "
            "resource allocation, and targeted support under the current filters."
        )

    st.text_area(
        f"Section Conclusion - {section_title}",
        value=conclusion,
        height=105,
        disabled=True
    )


def explode_multiselect_column(data, column):
    exploded = data.copy()

    exploded[column] = (
        exploded[column]
        .astype("string")
        .str.strip()
        .str.replace(r"\s*[,;/|]\s*", "|", regex=True)
        .str.split("|")
    )

    exploded = exploded.explode(column)
    exploded[column] = exploded[column].astype("string").str.strip()

    exploded = exploded[
        exploded[column].notna()
        & (exploded[column] != "")
        & (exploded[column].str.lower() != "nan")
    ]

    return exploded


def prepare_checkbox_long_data(data, id_cols, checkbox_cols, name_col, value_col):
    long_df = data[id_cols + checkbox_cols].copy()

    for col in checkbox_cols:
        long_df[col] = clean_yes_no_series(long_df[col])

    long_df = long_df.melt(
        id_vars=id_cols,
        value_vars=checkbox_cols,
        var_name=name_col,
        value_name=value_col
    )

    long_df = long_df[long_df[value_col]]

    long_df[name_col] = (
        long_df[name_col]
        .str.replace("support_", "", regex=False)
        .str.replace("issue_", "", regex=False)
        .str.replace("_", " ", regex=False)
        .str.title()
    )

    return long_df


def add_cumulative_hover(fig, category_label, orientation="v"):
    if orientation == "h":
        category_ref = "%{y}"
    else:
        category_ref = "%{x}"

    fig.update_traces(
        hovertemplate=(
            f"<b>{category_label}:</b> {category_ref}<br>"
            "<b>Gender:</b> %{customdata[0]}<br>"
            "<b>Count:</b> %{customdata[1]}<br>"
            "<b>Scope:</b> Current filtered selection"
            "<extra></extra>"
        )
    )

    return fig


@st.cache_data
def load_data():
    file_path = "./data/CFS_QUESTIONNAIRE_-_Tdh_Kenya_Draft3.xlsx"

    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    text_cols = df.select_dtypes(include="object").columns

    for col in text_cols:
        df[col] = df[col].astype("string").str.strip()

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"])

    df["child_age"] = pd.to_numeric(df["child_age"], errors="coerce")

    bins = [0, 4, 11, 17, 100]
    labels = ["0-4 years", "5-11 years", "12-17 years", "18+"]

    df["age_group"] = pd.cut(
        df["child_age"],
        bins=bins,
        labels=labels,
        right=True,
        include_lowest=True
    )

    return df


try:
    df = load_data()
except Exception as e:
    st.error(f"Error loading data: {e}")
    st.stop()


# --- SIDEBAR FILTERS ---
st.sidebar.header("Filter Hierarchy")

global_min_date = df["date"].min().date()
global_max_date = df["date"].max().date()

st.sidebar.markdown("**1. Date Range**")

start_d = st.sidebar.date_input(
    "From",
    value=global_min_date,
    min_value=global_min_date,
    max_value=global_max_date
)

end_d = st.sidebar.date_input(
    "To",
    value=global_max_date,
    min_value=global_min_date,
    max_value=global_max_date
)

if start_d > end_d:
    st.sidebar.error("'From' date cannot be later than 'To' date.")
    st.error("Invalid date range selected. Please choose a valid 'From' and 'To' date.")
    st.stop()

date_filtered_df = df[df["date"].dt.date.between(start_d, end_d)].copy()

camp_options = get_sorted_options(date_filtered_df, "camp_name")
selected_camps = st.sidebar.multiselect("2. Camp Name", options=camp_options)

camp_filtered_df = date_filtered_df.copy()

if selected_camps:
    camp_filtered_df = camp_filtered_df[
        camp_filtered_df["camp_name"].isin(selected_camps)
    ]

location_options = get_sorted_options(camp_filtered_df, "camp_location")
selected_locations = st.sidebar.multiselect("3. Camp Location", options=location_options)

location_filtered_df = camp_filtered_df.copy()

if selected_locations:
    location_filtered_df = location_filtered_df[
        location_filtered_df["camp_location"].isin(selected_locations)
    ]

staff_options = get_sorted_options(location_filtered_df, "staff_name")
selected_staff = st.sidebar.multiselect("4. Staff Name", options=staff_options)

f_df = location_filtered_df.copy()

if selected_staff:
    f_df = f_df[f_df["staff_name"].isin(selected_staff)]

f_df = f_df.copy()


st.title("🛡️ CFS Data Dashboard - Tdh Kenya")
st.markdown("### Executive Summary Statistics")

if f_df.empty:
    st.warning("No data found for the current filters.")
    st.stop()


# --- 1. EXECUTIVE SUMMARY ---
m1, m2, m3, m4, m5 = st.columns(5)

referral_rate = clean_yes_no_series(f_df["referral_made"]).mean()
disability_rate = clean_yes_no_series(f_df["has_disability"]).mean()
avg_child_age = f_df["child_age"].mean()

m1.metric("Total CFS Visits", len(f_df))
m2.metric("Referral Rate", f"{referral_rate:.1%}")
m3.metric("Disability Prev.", f"{disability_rate:.1%}")
m4.metric("Avg Child Age", round(avg_child_age, 1) if pd.notna(avg_child_age) else "N/A")
m5.metric("Active Staff", f_df["staff_name"].nunique())

show_section_conclusion(
    "Executive Summary Statistics",
    f"""
    Total visits: {len(f_df)}
    Referral rate: {referral_rate:.1%}
    Disability prevalence: {disability_rate:.1%}
    Average child age: {round(avg_child_age, 1) if pd.notna(avg_child_age) else "N/A"}
    Active staff: {f_df["staff_name"].nunique()}
    Date range: {start_d} to {end_d}
    """
)

st.divider()


# --- 2. CPV SUBMISSION TABLE ---
st.subheader("Distribution of CFS data as per the CPVs Data Submission")

cpv_tab = (
    f_df
    .groupby("staff_name")["gender"]
    .value_counts()
    .unstack(fill_value=0)
)

cpv_tab["Total"] = cpv_tab.sum(axis=1)

grand_total_cpv = pd.DataFrame(cpv_tab.sum()).T
grand_total_cpv.index = ["Grand Total"]

st.table(pd.concat([cpv_tab, grand_total_cpv]))

show_section_conclusion(
    "CPV Data Submission",
    cpv_tab.sort_values("Total", ascending=False).head(8).to_string()
)

st.divider()


# --- 3. ATTENDANCE & GENDER BY CFS ---
c1, c2 = st.columns(2)

with c1:
    st.subheader("Is this the first visit to CFS's")

    visit_tab = (
        f_df
        .groupby(["cfs_name", "first_visit", "gender"])
        .size()
        .unstack(level=[1, 2], fill_value=0)
    )

    st.dataframe(visit_tab, use_container_width=True)

    show_section_conclusion(
        "First Visit to CFS",
        visit_tab.head(10).to_string()
    )

with c2:
    st.subheader("Gender of child seeking information by CFS")

    gender_cfs = (
        f_df
        .groupby(["cfs_name", "gender"])
        .size()
        .unstack(fill_value=0)
    )

    gender_cfs["Total"] = gender_cfs.sum(axis=1)

    st.table(gender_cfs)

    show_section_conclusion(
        "Gender by CFS",
        gender_cfs.sort_values("Total", ascending=False).head(10).to_string()
    )

st.divider()


# --- 4. AGE GROUP ---
st.subheader("Age Group Breakdown by CFS & Gender")

age_pivot = (
    f_df
    .groupby(["cfs_name", "age_group", "gender"], observed=False)
    .size()
    .unstack(level=[1, 2], fill_value=0)
)

st.dataframe(age_pivot, use_container_width=True)

age_chart_df = (
    f_df
    .dropna(subset=["age_group"])
    .groupby(["cfs_name", "age_group", "gender"], observed=False)
    .size()
    .reset_index(name="Count")
)

if not age_chart_df.empty:
    fig_age = px.bar(
        age_chart_df,
        x="age_group",
        y="Count",
        color="gender",
        barmode="group",
        facet_col="cfs_name",
        facet_col_wrap=2,
        custom_data=["cfs_name", "gender", "Count"],
        labels={
            "age_group": "Age Group",
            "Count": "Number of Children",
            "gender": "Gender",
            "cfs_name": "CFS Name"
        },
        title="Age Group Distribution by CFS and Gender"
    )

    fig_age.update_traces(
        hovertemplate=(
            "<b>Age Group:</b> %{x}<br>"
            "<b>CFS Name:</b> %{customdata[0]}<br>"
            "<b>Gender:</b> %{customdata[1]}<br>"
            "<b>Count:</b> %{customdata[2]}"
            "<extra></extra>"
        )
    )

    fig_age.update_layout(
        legend_title_text="Gender",
        height=420 + (220 * max(0, age_chart_df["cfs_name"].nunique() - 2))
    )

    fig_age.for_each_annotation(
        lambda annotation: annotation.update(
            text=annotation.text.replace("cfs_name=", "")
        )
    )

    st.plotly_chart(fig_age, use_container_width=True)

show_section_conclusion(
    "Age Group Breakdown by CFS and Gender",
    age_chart_df.sort_values("Count", ascending=False).head(12).to_string(index=False)
)

st.divider()


# --- 5. DISABILITY PREVALENCE & TYPES ---
st.subheader("Disability Prevalence & Disability Types")

d1, d2 = st.columns([1, 1.5])

with d1:
    st.write("**Prevalence With/Without Disability**")

    dis_prev = (
        f_df
        .groupby(["cfs_name", "has_disability", "gender"])
        .size()
        .unstack(level=[1, 2], fill_value=0)
    )

    st.dataframe(dis_prev, use_container_width=True)

with d2:
    st.write("**Specific Disability Types**")

    dis_f = f_df[clean_yes_no_series(f_df["has_disability"])]

    if not dis_f.empty:
        dis_type_chart_df = (
            dis_f
            .groupby(["disability_type", "gender"])
            .size()
            .reset_index(name="Count")
        )

        fig_disability = px.bar(
            dis_type_chart_df,
            x="disability_type",
            y="Count",
            color="gender",
            barmode="stack",
            custom_data=["gender", "Count"],
            labels={
                "disability_type": "Disability Type",
                "Count": "Number of Children",
                "gender": "Gender"
            },
            title="Cumulative Disability Types by Gender"
        )

        fig_disability = add_cumulative_hover(fig_disability, "Disability Type")

        fig_disability.update_layout(
            xaxis_tickangle=-35,
            legend_title_text="Gender"
        )

        st.plotly_chart(fig_disability, use_container_width=True)
    else:
        st.info("No disability records found for the current selection.")

show_section_conclusion(
    "Disability Prevalence and Types",
    f"""
    Disability prevalence by CFS, status, and gender:
    {dis_prev.head(10).to_string()}

    Top disability types:
    {dis_f["disability_type"].value_counts().head(10).to_string() if not dis_f.empty else "No disability records"}
    """
)

st.divider()


# --- 6. NATURE OF ISSUES REPORTED ---
st.subheader("Nature of Issues Reported")

issue_cols = [
    c for c in f_df.columns
    if c.startswith("issue_")
    and c not in ["issue_other_specified", "issue_other", "issue_none"]
]

if issue_cols:
    issue_long = prepare_checkbox_long_data(
        data=f_df,
        id_cols=["cfs_name", "gender"],
        checkbox_cols=issue_cols,
        name_col="issue_type",
        value_col="issue_selected"
    )

    if not issue_long.empty:
        issue_table_col, issue_chart_col = st.columns([1.3, 1])

        with issue_table_col:
            issue_gender_tab = pd.pivot_table(
                issue_long,
                index="issue_type",
                columns=["cfs_name", "gender"],
                values="issue_selected",
                aggfunc="count",
                fill_value=0,
                margins=True,
                margins_name="Grand Total"
            )

            st.dataframe(issue_gender_tab, use_container_width=True)

        with issue_chart_col:
            issue_chart_df = (
                issue_long
                .groupby(["issue_type", "gender"])
                .size()
                .reset_index(name="Count")
            )

            fig_issues = px.bar(
                issue_chart_df,
                x="Count",
                y="issue_type",
                color="gender",
                orientation="h",
                barmode="stack",
                custom_data=["gender", "Count"],
                labels={
                    "issue_type": "Issue Reported",
                    "Count": "Number of Children",
                    "gender": "Gender"
                },
                title="Cumulative Issues Reported by Gender"
            )

            fig_issues = add_cumulative_hover(
                fig_issues,
                "Issue Reported",
                orientation="h"
            )

            fig_issues.update_layout(
                legend_title_text="Gender",
                yaxis={"categoryorder": "total ascending"}
            )

            st.plotly_chart(fig_issues, use_container_width=True)

        show_section_conclusion(
            "Nature of Issues Reported",
            issue_chart_df.sort_values("Count", ascending=False).head(12).to_string(index=False)
        )
    else:
        st.info("No issue records found for the current selection.")
else:
    st.info("No issue columns found in the dataset.")

st.divider()


# --- 7. SUPPORT OFFERED BY CFS SITE ---
st.subheader("Support Offered by CFS Site")

sup_cols = [
    c for c in f_df.columns
    if c.startswith("support_") and c != "support_none"
]

if sup_cols:
    support_long = prepare_checkbox_long_data(
        data=f_df,
        id_cols=["cfs_name", "gender"],
        checkbox_cols=sup_cols,
        name_col="support_type",
        value_col="support_selected"
    )

    if not support_long.empty:
        support_table_col, support_chart_col = st.columns([1.3, 1])

        with support_table_col:
            support_gender_tab = pd.pivot_table(
                support_long,
                index="support_type",
                columns=["cfs_name", "gender"],
                values="support_selected",
                aggfunc="count",
                fill_value=0,
                margins=True,
                margins_name="Grand Total"
            )

            st.dataframe(support_gender_tab, use_container_width=True)

        with support_chart_col:
            support_chart_df = (
                support_long
                .groupby(["support_type", "gender"])
                .size()
                .reset_index(name="Count")
            )

            fig_support = px.bar(
                support_chart_df,
                x="support_type",
                y="Count",
                color="gender",
                barmode="stack",
                custom_data=["gender", "Count"],
                labels={
                    "support_type": "Support Offered",
                    "Count": "Number of Children",
                    "gender": "Gender"
                },
                title="Cumulative Support Offered by Gender"
            )

            fig_support = add_cumulative_hover(fig_support, "Support Offered")

            fig_support.update_layout(
                xaxis_tickangle=-35,
                legend_title_text="Gender"
            )

            st.plotly_chart(fig_support, use_container_width=True)

        show_section_conclusion(
            "Support Offered by CFS Site",
            support_chart_df.sort_values("Count", ascending=False).head(12).to_string(index=False)
        )
    else:
        st.info("No support records found for the current selection.")
else:
    st.info("No support columns found in the dataset.")

st.divider()


# --- 8. GAMES & REFERRALS ---
g1, g2 = st.columns(2)

with g1:
    st.subheader("Games child played/engaged with")

    games_df = explode_multiselect_column(f_df, "activities_engaged")

    if not games_df.empty:
        games_gender_tab = pd.pivot_table(
            games_df,
            index="activities_engaged",
            columns=["cfs_name", "gender"],
            values="date",
            aggfunc="count",
            fill_value=0,
            margins=True,
            margins_name="Grand Total"
        )

        st.dataframe(games_gender_tab, use_container_width=True)

        games_chart_df = (
            games_df
            .groupby(["activities_engaged", "gender"])
            .size()
            .reset_index(name="Count")
        )

        fig_games = px.bar(
            games_chart_df,
            x="activities_engaged",
            y="Count",
            color="gender",
            barmode="stack",
            custom_data=["gender", "Count"],
            labels={
                "activities_engaged": "Game / Activity",
                "Count": "Number of Children",
                "gender": "Gender"
            },
            title="Cumulative Games / Activities by Gender"
        )

        fig_games = add_cumulative_hover(fig_games, "Game / Activity")

        fig_games.update_layout(
            xaxis_tickangle=-35,
            legend_title_text="Gender"
        )

        st.plotly_chart(fig_games, use_container_width=True)

        show_section_conclusion(
            "Games Child Played or Engaged With",
            games_chart_df.sort_values("Count", ascending=False).head(12).to_string(index=False)
        )
    else:
        st.info("No games or activities found for the current selection.")

with g2:
    st.subheader("Referral Destination If Yes")

    ref_yes = f_df[clean_yes_no_series(f_df["referral_made"])].copy()

    if not ref_yes.empty:
        ref_yes["referral_destination"] = (
            ref_yes["referral_destination"]
            .astype("string")
            .str.strip()
        )

        ref_yes = ref_yes[
            ref_yes["referral_destination"].notna()
            & (ref_yes["referral_destination"] != "")
            & (ref_yes["referral_destination"].str.lower() != "nan")
        ]

        if not ref_yes.empty:
            referral_gender_tab = (
                ref_yes
                .groupby(["referral_destination", "gender"])
                .size()
                .unstack(fill_value=0)
            )

            referral_gender_tab["Total"] = referral_gender_tab.sum(axis=1)
            referral_gender_tab = referral_gender_tab.sort_values(
                by="Total",
                ascending=False
            )

            grand_total_referral = pd.DataFrame(referral_gender_tab.sum()).T
            grand_total_referral.index = ["Grand Total"]

            st.dataframe(
                pd.concat([referral_gender_tab, grand_total_referral]),
                use_container_width=True
            )

            referral_chart_df = (
                ref_yes
                .groupby(["referral_destination", "gender"])
                .size()
                .reset_index(name="Count")
            )

            fig_referrals = px.bar(
                referral_chart_df,
                x="referral_destination",
                y="Count",
                color="gender",
                barmode="stack",
                custom_data=["gender", "Count"],
                labels={
                    "referral_destination": "Referral Destination",
                    "Count": "Number of Referrals",
                    "gender": "Gender"
                },
                title="Cumulative Referral Destinations by Gender"
            )

            fig_referrals = add_cumulative_hover(
                fig_referrals,
                "Referral Destination"
            )

            fig_referrals.update_layout(
                xaxis_tickangle=-35,
                legend_title_text="Gender"
            )

            st.plotly_chart(fig_referrals, use_container_width=True)

            show_section_conclusion(
                "Referral Destination If Yes",
                referral_chart_df.sort_values("Count", ascending=False).head(12).to_string(index=False)
            )
        else:
            st.info("Referral records exist, but no referral destination was provided.")
    else:
        st.info("No referrals found for the current selection.")

st.divider()


# --- 9. EXTERNAL REFERRAL AGENCY BREAKDOWN ---
st.subheader("External Referral Agency Breakdown")

ref_yes_external = f_df[clean_yes_no_series(f_df["referral_made"])].copy()

if "referral_destination" not in ref_yes_external.columns:
    st.info("The column 'referral_destination' was not found in the dataset.")
elif "referral_external_agency" not in ref_yes_external.columns:
    st.info("The column 'referral_external_agency' was not found in the dataset.")
else:
    ref_yes_external["referral_destination"] = (
        ref_yes_external["referral_destination"]
        .astype("string")
        .str.strip()
    )

    external_referrals = ref_yes_external[
        ref_yes_external["referral_destination"].str.lower().eq("external")
    ].copy()

    if not external_referrals.empty:
        external_referrals["referral_external_agency"] = (
            external_referrals["referral_external_agency"]
            .astype("string")
            .str.strip()
        )

        external_referrals = external_referrals[
            external_referrals["referral_external_agency"].notna()
            & (external_referrals["referral_external_agency"] != "")
            & (external_referrals["referral_external_agency"].str.lower() != "nan")
        ]

        if not external_referrals.empty:
            e1, e2 = st.columns([1.2, 1])

            with e1:
                external_agency_tab = (
                    external_referrals
                    .groupby(["referral_external_agency", "gender"])
                    .size()
                    .unstack(fill_value=0)
                )

                external_agency_tab["Total"] = external_agency_tab.sum(axis=1)
                external_agency_tab = external_agency_tab.sort_values(
                    by="Total",
                    ascending=False
                )

                grand_total_external = pd.DataFrame(external_agency_tab.sum()).T
                grand_total_external.index = ["Grand Total"]

                st.dataframe(
                    pd.concat([external_agency_tab, grand_total_external]),
                    use_container_width=True
                )

            with e2:
                external_chart_df = (
                    external_referrals
                    .groupby(["referral_external_agency", "gender"])
                    .size()
                    .reset_index(name="Count")
                )

                fig_external = px.bar(
                    external_chart_df,
                    x="referral_external_agency",
                    y="Count",
                    color="gender",
                    barmode="stack",
                    custom_data=["gender", "Count"],
                    labels={
                        "referral_external_agency": "External Agency",
                        "Count": "Number of Referrals",
                        "gender": "Gender"
                    },
                    title="Cumulative External Referral Agencies by Gender"
                )

                fig_external = add_cumulative_hover(
                    fig_external,
                    "External Agency"
                )

                fig_external.update_layout(
                    xaxis_tickangle=-35,
                    legend_title_text="Gender"
                )

                st.plotly_chart(fig_external, use_container_width=True)

            show_section_conclusion(
                "External Referral Agency Breakdown",
                external_chart_df.sort_values("Count", ascending=False).head(12).to_string(index=False)
            )
        else:
            st.info("External referrals exist, but no external agency was provided.")
    else:
        st.info("No external referrals found for the current selection.")


st.markdown("---")
st.caption("Developed for Tdh Kenya - @johnkul_MEAL Officer.")