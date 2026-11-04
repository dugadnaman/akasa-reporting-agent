"""Akasa Air Marketing Reports Consolidator & Cleaner Web Application.

Streamlit-based UI for uploading, normalizing, aggregating, and exporting
multi-channel MoEngage campaign attribution reports.
"""

from __future__ import annotations

import os
from typing import Optional

import pandas as pd
import streamlit as st

import pipeline

# Page Configuration
st.set_page_config(
    page_title="Akasa Air — Campaign Report Consolidator",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #FF5500;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #555555;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background: #F8F9FA;
        border-left: 4px solid #FF5500;
        padding: 1rem;
        border-radius: 6px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    .stDownloadButton button {
        background-color: #FF5500 !important;
        color: white !important;
        font-weight: 600 !important;
        padding: 0.5rem 1.5rem !important;
        border-radius: 6px !important;
        border: none !important;
    }
    .stDownloadButton button:hover {
        background-color: #E04800 !important;
    }
</style>
""", unsafe_allow_html=True)


def main():
    st.markdown('<div class="main-header">✈️ Akasa Air — Campaign Report Consolidator</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-header">'
        'Automated normalization, multi-channel aggregation (Email, WhatsApp, App Push, Web Push, Flows), '
        'and Flight Booking (<code>FB_Successful</code>) attribution reporting.'
        '</div>',
        unsafe_allow_html=True
    )

    # Sidebar: Data Source Selection & Settings
    st.sidebar.header("📁 Data Source & Settings")
    data_mode = st.sidebar.radio(
        "Report Input Mode",
        options=["Pre-loaded Sample Data", "Upload Custom Reports"],
        index=0,
        help="Choose whether to run on the bundled MoEngage sample exports or upload your own weekly/monthly files."
    )

    normalize_tags = st.sidebar.checkbox(
        "Normalize Campaign Tags",
        value=True,
        help="Sorts and canonicalizes list tags (e.g. ['AMD', 'OneWayAbandoned_V3'] and ['OneWayAbandoned_V3', 'AMD'] merge into one campaign)."
    )

    email_src: Optional[any] = None
    wa_src: Optional[any] = None
    push_src: Optional[any] = None
    overall_src: Optional[any] = None

    sample_base = os.path.join(os.path.dirname(__file__), "AkasaReportsCleaner", "csvFiles")

    if data_mode == "Pre-loaded Sample Data":
        st.sidebar.success("Using bundled sample exports from `AkasaReportsCleaner/csvFiles/`.")
        email_src = os.path.join(sample_base, "email.xls")
        wa_src = os.path.join(sample_base, "whatsapp.xls")
        push_src = os.path.join(sample_base, "push.xls")
        overall_src = os.path.join(sample_base, "overall.xls")

        missing = [f for f in [email_src, wa_src, push_src, overall_src] if not os.path.exists(f)]
        if missing:
            st.sidebar.error(f"Missing sample files: {missing}")
    else:
        st.sidebar.markdown("Upload MoEngage channel exports:")
        email_file = st.sidebar.file_uploader("1. Email Report (.xls / .xlsx)", type=["xls", "xlsx", "csv"])
        wa_file = st.sidebar.file_uploader("2. WhatsApp Report (.xls / .xlsx)", type=["xls", "xlsx", "csv"])
        push_file = st.sidebar.file_uploader("3. Push Notifications Report (.xls / .xlsx)", type=["xls", "xlsx", "csv"])
        overall_file = st.sidebar.file_uploader("4. Overall Flows Report (.xls / .xlsx)", type=["xls", "xlsx", "csv"])

        email_src = email_file
        wa_src = wa_file
        push_src = push_file
        overall_src = overall_file

    # Check if at least one file is provided
    sources_provided = any(s is not None for s in [email_src, wa_src, push_src, overall_src])

    if not sources_provided:
        st.info("👋 Please select 'Pre-loaded Sample Data' or upload at least one channel export file to begin.")
        return

    # Run Pipeline
    with st.spinner("Processing and consolidating campaign reports..."):
        try:
            df = pipeline.generate_consolidated_report(
                email_file=email_src,
                whatsapp_file=wa_src,
                push_file=push_src,
                overall_file=overall_src,
                normalize_tags=normalize_tags
            )
        except Exception as e:
            st.error(f"Error during report processing: {e}")
            st.exception(e)
            return

    if df.empty:
        st.warning("No data found in the provided reports after filtering.")
        return

    # Top KPI Metrics Cards
    total_sent = int(df["Total Sent"].sum())
    total_conv_users = int(
        df["Goal 1 View Through Converted Users"].sum() +
        df["Goal 1 Click Through Converted Users"].sum()
    )
    direct_click_rev = float(df["Goal 1 Click Through Total Revenue"].sum())
    flow_view_rev = float(df["View Through Revenue (Primary Goal)"].sum())
    unique_camps = df["campaign"].nunique()

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Total Messages Sent", f"{total_sent:,}")
    with col2:
        st.metric("Total Converted Users", f"{total_conv_users:,}")
    with col3:
        st.metric("Direct Click Revenue", f"₹{direct_click_rev:,.0f}")
    with col4:
        st.metric("Flow Journey Revenue", f"₹{flow_view_rev:,.0f}")
    with col5:
        st.metric("Campaigns Analyzed", f"{unique_camps}")

    st.markdown("---")

    # Interactive Visualizations
    tab_overview, tab_data, tab_audit = st.tabs(["📊 Performance Insights", "📋 Data Explorer", "🛠️ Hygiene & Fixes"])

    with tab_overview:
        col_chart1, col_chart2 = st.columns(2)

        with col_chart1:
            st.subheader("Direct Revenue by Channel (₹)")
            channel_rev = df[df["Channel"] != "overAll"].groupby("Channel")["Goal 1 Click Through Total Revenue"].sum().reset_index()
            if not channel_rev.empty and channel_rev["Goal 1 Click Through Total Revenue"].sum() > 0:
                st.bar_chart(channel_rev.set_index("Channel")["Goal 1 Click Through Total Revenue"], color="#FF5500")
            else:
                st.info("No direct click revenue recorded for messaging channels.")

        with col_chart2:
            st.subheader("Top Campaigns by Direct Revenue (₹)")
            top_camps = df[df["Channel"] != "overAll"].groupby("campaign")["Goal 1 Click Through Total Revenue"].sum().reset_index()
            top_camps = top_camps.sort_values(by="Goal 1 Click Through Total Revenue", ascending=False).head(5)
            if not top_camps.empty and top_camps["Goal 1 Click Through Total Revenue"].sum() > 0:
                st.bar_chart(top_camps.set_index("campaign")["Goal 1 Click Through Total Revenue"], color="#1A237E")
            else:
                st.info("No campaign direct revenue recorded.")

        # Flow Journeys summary if present
        flow_df = df[df["Channel"] == "overAll"]
        if not flow_df.empty and flow_df["View Through Revenue (Primary Goal)"].sum() > 0:
            st.subheader("Top User Journey Flows by View-Through Revenue (₹)")
            top_flows = flow_df.sort_values(by="View Through Revenue (Primary Goal)", ascending=False).head(5)
            st.bar_chart(top_flows.set_index("campaign")["View Through Revenue (Primary Goal)"], color="#2E7D32")

    with tab_data:
        st.subheader("Consolidated Campaign Table")

        # Filters
        fcol1, fcol2, fcol3 = st.columns([2, 1, 1])
        with fcol1:
            search_query = st.text_input("🔍 Search Campaign Tag", "")
        with fcol2:
            channels = sorted(df["Channel"].unique().tolist())
            selected_channels = st.multiselect("Filter Channel", options=channels, default=channels)
        with fcol3:
            hide_zero_revenue = st.checkbox("Hide ₹0 Revenue Rows", value=False)

        filtered_df = df[df["Channel"].isin(selected_channels)].copy()
        if search_query:
            filtered_df = filtered_df[filtered_df["campaign"].str.contains(search_query, case=False, na=False)]
        if hide_zero_revenue:
            rev_mask = (filtered_df["Goal 1 Click Through Total Revenue"] > 0) | (filtered_df["View Through Revenue (Primary Goal)"] > 0)
            filtered_df = filtered_df[rev_mask]

        st.dataframe(filtered_df, width="stretch", height=450)

        # Download Button
        st.markdown("### 📥 Export Executive Report")
        col_down1, col_down2 = st.columns([1, 3])
        with col_down1:
            excel_bytes = pipeline.export_to_excel_bytes(filtered_df)
            st.download_button(
                label="📥 Download Excel (.xlsx)",
                data=excel_bytes,
                file_name="final_campaign_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_down2:
            st.caption("Formatted with executive styling, auto-column widths, and clean blanks for non-applicable metrics.")

    with tab_audit:
        st.subheader("Automated Data Hygiene & Fix Log")
        st.markdown("""
        The engine automatically applies the following fixes:
        - **Tag Canonicalization**: Resolves list order permutations (e.g., `['AMD', 'OneWayAbandoned_V3']` vs `['OneWayAbandoned_V3', 'AMD']`) so split rows are properly merged.
        - **Column Name Unification**: Standardizes `Total clicks` and `Total Clicks` into a single consolidated column.
        - **App Push Mobile Aggregation**: Merges Android and iOS sent, impressions, clicks, conversions, and revenues.
        - **Web Push Mapping**: Isolates Web Push rows and aligns metrics with standard channel naming.
        - **Filter Logic Correction**: Ensures zero-activity draft/inactive campaigns are cleanly filtered before grouping.
        - **Primary Goal Focus**: Strips secondary tracking goals (`Goal 2`, `Goal 3`) to isolate net ticket conversions (`FB_Successful`).
        """)


if __name__ == "__main__":
    main()
