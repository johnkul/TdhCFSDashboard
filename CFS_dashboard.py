import streamlit as st
import pandas as pd
import plotly.express as px

# --- APP CONFIG ---
st.set_page_config(page_title="Tdh Kenya | CFS Data Dashboard", layout="wide", page_icon="📊")

# --- CUSTOM CSS FOR TABLE STYLING ---
st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border: 1px solid #f0f2f6; }
    thead tr th { background-color: #0083B8 !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# --- DATA LOADING ---
@st.cache_data
def load_data():
    file_path = "./data/CFS_QUESTIONNAIRE_-_Tdh_Kenya_Draft2.xlsx"
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()
    df['date'] = pd.to_datetime(df['date'])
    
    # Age Binning
    bins = [0, 4, 11, 17, 100]
    labels = ['0-4 years', '5-11 years', '12-17 years', '18+']
    df['age_group'] = pd.cut(df['child_age'], bins=bins, labels=labels, right=True)
    return df

try:
    df = load_data()
except Exception as e:
    st.error(f"Error loading data: {e}")
    st.stop()

# --- SIDEBAR: CASCADING FILTERS ---
st.sidebar.header("Filter Hierarchy")
all_camps = sorted(df['camp_name'].unique())
selected_camps = st.sidebar.multiselect("1. Camp Name", options=all_camps)

temp_df = df[df['camp_name'].isin(selected_camps)] if selected_camps else df
all_locs = sorted(temp_df['camp_location'].unique())
selected_locations = st.sidebar.multiselect("2. Camp Location", options=all_locs)

temp_df = temp_df[temp_df['camp_location'].isin(selected_locations)] if selected_locations else temp_df
all_staff = sorted(temp_df['staff_name'].unique())
selected_staff = st.sidebar.multiselect("3. Staff Name", options=all_staff)

min_date, max_date = df['date'].min().date(), df['date'].max().date()
date_selection = st.sidebar.date_input("4. Date Range", [min_date, max_date])

# --- APPLY FINAL FILTER MASK ---
mask = pd.Series([True] * len(df), index=df.index)
if selected_camps: mask &= df['camp_name'].isin(selected_camps)
if selected_locations: mask &= df['camp_location'].isin(selected_locations)
if selected_staff: mask &= df['staff_name'].isin(selected_staff)
if isinstance(date_selection, list) and len(date_selection) == 2:
    start_d, end_d = date_selection
    mask &= (df['date'].dt.date >= start_d) & (df['date'].dt.date <= end_d)

f_df = df[mask]

# --- 1. EXECUTIVE SUMMARY (METRIC CARDS) ---
st.title("🛡️ CFS Data Dashboard - Tdh Kenya")
st.markdown("### Executive Summary Statistics")

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Total CFS Visits", len(f_df))
m2.metric("Referral Rate", f"{(f_df['referral_made'] == 'Yes').mean():.1%}")
m3.metric("Disability Prev.", f"{(f_df['has_disability'] == 'Yes').mean():.1%}")
m4.metric("Avg Child Age", round(f_df['child_age'].mean(), 1))
m5.metric("Active Staff", f_df['staff_name'].nunique())

st.divider()

if f_df.empty:
    st.warning("No data found for the current filters.")
    st.stop()

# --- 2. CPV SUBMISSION TABLE ---
st.subheader("Distribution of CFS data as per the CPVs Data Submission")
cpv_tab = f_df.groupby('staff_name')['gender'].value_counts().unstack(fill_value=0)
cpv_tab['Total'] = cpv_tab.sum(axis=1)
grand_total_cpv = pd.DataFrame(cpv_tab.sum()).T
grand_total_cpv.index = ['Grand Total']
st.table(pd.concat([cpv_tab, grand_total_cpv]))

# --- 3. ATTENDANCE & GENDER BY CFS ---
st.divider()
c1, c2 = st.columns(2)
with c1:
    st.subheader("Is this the first visit to CFS's")
    visit_tab = f_df.groupby(['cfs_name', 'first_visit', 'gender']).size().unstack(level=[1,2], fill_value=0)
    st.dataframe(visit_tab, use_container_width=True)
with c2:
    st.subheader("Gender of child seeking information by CFS")
    gender_cfs = f_df.groupby(['cfs_name', 'gender']).size().unstack(fill_value=0)
    gender_cfs['Total'] = gender_cfs.sum(axis=1)
    st.table(gender_cfs)

# --- 4. AGE GROUP (EXCEL STYLE) ---
st.divider()
st.subheader("Age Group Breakdown (By CFS & Gender)")
age_pivot = f_df.groupby(['cfs_name', 'age_group', 'gender']).size().unstack(level=[1,2], fill_value=0)
st.dataframe(age_pivot, use_container_width=True)

# --- 5. DISABILITY PREVALENCE & TYPES ---
st.divider()
st.subheader("Disability Prevalence & Disability Types")
d1, d2 = st.columns([1, 1.5])
with d1:
    st.write("**Prevalence (With/Without Disability)**")
    dis_prev = f_df.groupby(['cfs_name', 'has_disability', 'gender']).size().unstack(level=[1,2], fill_value=0)
    st.dataframe(dis_prev)
with d2:
    st.write("**Specific Disability Types**")
    dis_f = f_df[f_df['has_disability'] == 'Yes']
    if not dis_f.empty:
        dis_type_tab = dis_f.groupby(['disability_type', 'gender']).size().unstack(fill_value=0)
        st.bar_chart(dis_type_tab)

# --- 6. ISSUES & SUPPORT ---
st.divider()
i1, i2 = st.columns(2)
with i1:
    st.subheader("Nature of Issues Reported")
    issue_cols = [c for c in f_df.columns if c.startswith('issue_') and c not in ['issue_other_specified', 'issue_other', 'issue_none']]
    issue_sums = f_df[issue_cols].sum().sort_values(ascending=True)
    issue_sums.index = [i.replace('issue_', '').replace('_', ' ').title() for i in issue_sums.index]
    st.plotly_chart(px.bar(issue_sums, orientation='h', color_discrete_sequence=['#E53935']), use_container_width=True)
with i2:
    st.subheader("Support Offered by CFS Site")
    sup_cols = [c for c in f_df.columns if c.startswith('support_') and c != 'support_none']
    sup_matrix = f_df.groupby('cfs_name')[sup_cols].sum()
    sup_matrix.columns = [s.replace('support_', '').replace('_', ' ').title() for s in sup_matrix.columns]
    st.dataframe(sup_matrix, use_container_width=True)

# --- 7. GAMES & REFERRALS ---
st.divider()
g1, g2 = st.columns(2)
with g1:
    st.subheader("Games child played/engaged with")
    games_tab = f_df.groupby(['cfs_name', 'activities_engaged']).size().unstack(fill_value=0)
    st.dataframe(games_tab, use_container_width=True)
with g2:
    st.subheader("Referral Destination (If Yes)")
    ref_yes = f_df[f_df['referral_made'] == 'Yes']
    if not ref_yes.empty:
        st.table(ref_yes['referral_destination'].value_counts())
    else:
        st.info("No referrals found for selection.")

st.markdown("---")
st.caption("Developed for Tdh Kenya-@johnkul_MEAL Officer.")