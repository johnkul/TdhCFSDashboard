"""
Optional UI formatting patch for the Tdh Kenya CFS Dashboard.
Copy the relevant functions into app.py if you want stronger table/card formatting.
"""

import pandas as pd
import streamlit as st


def page_header(title: str, subtitle: str | None = None):
    subtitle_html = f"<p>{subtitle}</p>" if subtitle else ""
    st.markdown(
        f"""
        <div class="dashboard-title">
            <h1>{title}</h1>
            {subtitle_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_header(title: str, subtitle: str | None = None):
    st.markdown(f"<div class='section-heading'>{title}</div>", unsafe_allow_html=True)
    if subtitle:
        st.markdown(f"<div class='section-subtitle'>{subtitle}</div>", unsafe_allow_html=True)


def metric_card(label, value, helper=None, tone="primary"):
    tone_map = {
        "primary": "#1d4ed8",
        "success": "#15803d",
        "warning": "#d97706",
        "danger": "#dc2626",
        "neutral": "#64748b",
    }
    border = tone_map.get(tone, tone_map["primary"])
    helper_html = f"<div class='metric-helper'>{helper}</div>" if helper else ""
    st.markdown(
        f"""
        <div class="metric-card" style="border-top:4px solid {border};">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            {helper_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def insight_card(label, value, helper=None):
    helper_html = f"<div class='insight-helper'>{helper}</div>" if helper else ""
    st.markdown(
        f"""
        <div class="insight-card">
            <div class="insight-label">{label}</div>
            <div class="insight-value">{value}</div>
            {helper_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def professional_table_style(table: pd.DataFrame, precision: int = 0):
    """Professional table style with highlighted Grand Total rows/columns."""
    if table.empty:
        return table

    def row_style(row):
        is_total = str(row.name) == "Grand Total" or any(str(v) == "Grand Total" for v in row.values)
        if is_total:
            return [
                "font-weight:900;background-color:#eef2ff;color:#0f172a;border-top:2px solid #1d4ed8;"
                for _ in row
            ]
        return ["" for _ in row]

    def total_col_style(col):
        if str(col.name).lower().startswith("total") or str(col.name).lower().endswith("total"):
            return ["font-weight:800;background-color:#f8fafc;" for _ in col]
        return ["" for _ in col]

    return (
        table.style
        .apply(row_style, axis=1)
        .apply(total_col_style, axis=0)
        .format(precision=precision, thousands=",")
        .set_table_styles([
            {"selector": "thead th", "props": [
                ("background", "linear-gradient(180deg,#1e40af,#172554)"),
                ("color", "#ffffff"),
                ("font-weight", "900"),
                ("padding", "10px 12px"),
                ("border", "1px solid #1e3a8a"),
                ("text-align", "left"),
            ]},
            {"selector": "tbody td", "props": [
                ("color", "#1f2937"),
                ("padding", "9px 12px"),
                ("border", "1px solid #e5edf5"),
            ]},
            {"selector": "tbody tr:nth-child(even)", "props": [
                ("background-color", "#f8fafc"),
            ]},
            {"selector": "tbody tr:hover", "props": [
                ("background-color", "#eff6ff"),
            ]},
            {"selector": "caption", "props": [
                ("caption-side", "top"),
                ("font-weight", "900"),
                ("color", "#172554"),
                ("padding", "0 0 .6rem 0"),
            ]},
        ])
    )


def style_total(table):
    return professional_table_style(table, precision=0)


def style_simple_total(df):
    return professional_table_style(df, precision=0)
