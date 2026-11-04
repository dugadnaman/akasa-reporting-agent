"""Akasa Air Marketing Reports Consolidator & Journey Tracker Web Application.

Streamlit-based UI for uploading, normalizing, aggregating, and exporting
multi-channel MoEngage campaign attribution reports and Journey Tracker (Book 1).
"""

from __future__ import annotations

import os
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st
import importlib
import pipeline
importlib.reload(pipeline)
st.set_page_config(
    page_title="Akasa Air — Campaign Report Consolidator & Journey Tracker",
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
    .cost-card {
        background: #F0F4F8;
        border-left: 4px solid #1F4E79;
        padding: 0.8rem;
        border-radius: 4px;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)


def main():
    st.markdown('<div class="main-header">✈️ Akasa Air — Campaign Consolidator & Journey Tracker</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-header">'
        'Automated normalization, multi-channel aggregation (Email, WhatsApp, App Push, Web Push, Flows), '
        'Flight Booking (<code>FB_Successful</code>) attribution, and executive Journey Tracker (Book 1) reporting with cost & ROI formulas.'
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

    # Sidebar: Cost Rates Configuration
    st.sidebar.markdown("---")
    st.sidebar.subheader("💰 Send Cost Rates")
    email_rate = st.sidebar.number_input(
        "Email Rate (₹ / Sent)",
        min_value=0.0,
        max_value=10.0,
        value=0.007,
        step=0.001,
        format="%.4f",
        help="Cost per sent email used in live Excel cost formula."
    )
    wa_rate = st.sidebar.number_input(
        "WhatsApp Rate (₹ / Sent)",
        min_value=0.0,
        max_value=10.0,
        value=0.007,
        step=0.001,
        format="%.4f",
        help="Cost per sent WhatsApp message used in live Excel cost formula."
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

    # Initialize Mappings in Session State
    if "journey_mappings" not in st.session_state:
        st.session_state.journey_mappings = dict(pipeline.DEFAULT_JOURNEY_MAPPINGS)

    # Ensure all current tags have mapping entries
    current_tags = df["campaign"].astype(str).unique().tolist()
    for t in current_tags:
        if t not in st.session_state.journey_mappings:
            st.session_state.journey_mappings[t] = t

    # Transform to Journey Tracker
    journey_df, j_warnings = pipeline.transform_to_journey_tracker(
        df,
        mappings=st.session_state.journey_mappings,
        email_rate=email_rate,
        wa_rate=wa_rate
    )

    # Main Tabs
    tab_journey, tab_overview, tab_data, tab_audit = st.tabs([
        "🎯 Journey Tracker (Book 1)",
        "📊 Performance Insights",
        "📋 Consolidated Campaign Report",
        "🛠️ Hygiene & Fixes"
    ])

    # TAB 1: JOURNEY TRACKER (BOOK 1)
    with tab_journey:
        st.subheader("🎯 Akasa Journey Tracker Format (Book 1)")
        st.markdown(
            "Converts the MoEngage campaign output into the official **Journey Tracker format** with send cost modeling, "
            "delivery & CTR rates, Average Order Value (AOV), and live Excel formulas for Cost and ROI."
        )

        # Mapping Editor Expander
        with st.expander("✏️ Customize Journey Name Mappings", expanded=False):
            st.caption("Map raw campaign tags from MoEngage exports into friendly journey names. Rows with identical journey names are automatically summed.")
            map_cols = st.columns(2)
            updated_mappings = dict(st.session_state.journey_mappings)
            for idx, tag in enumerate(sorted(current_tags)):
                col_target = map_cols[idx % 2]
                val = col_target.text_input(
                    label=f"Tag: {tag}",
                    value=st.session_state.journey_mappings.get(tag, tag),
                    key=f"map_{tag}"
                )
                updated_mappings[tag] = val

            mcol1, mcol2 = st.columns([1, 4])
            with mcol1:
                if st.button("Apply Mappings", type="primary"):
                    st.session_state.journey_mappings = updated_mappings
                    st.rerun()
            with mcol2:
                if st.button("Reset to Default Mappings"):
                    st.session_state.journey_mappings = dict(pipeline.DEFAULT_JOURNEY_MAPPINGS)
                    st.rerun()

        # Re-compute with current active mappings
        journey_df, j_warnings = pipeline.transform_to_journey_tracker(
            df,
            mappings=st.session_state.journey_mappings,
            email_rate=email_rate,
            wa_rate=wa_rate
        )

        # Journey Tracker Summary KPIs
        total_cost = journey_df["Cost"].dropna().sum()
        total_ct_rev = journey_df["CT Revenue"].sum()
        overall_roi = round(total_ct_rev / total_cost, 1) if total_cost > 0 else 0.0

        jcol1, jcol2, jcol3, jcol4 = st.columns(4)
        with jcol1:
            st.metric("Total Send Cost (Email + WA)", f"₹{total_cost:,.2f}")
        with jcol2:
            st.metric("Total Direct Click Revenue", f"₹{total_ct_rev:,.0f}")
        with jcol3:
            st.metric("Direct Marketing ROI", f"{overall_roi:,.1f}x")
        with jcol4:
            st.metric("Unique Journeys Tracked", f"{journey_df['Journey Name'].replace('', None).dropna().nunique()}")

        # Display Journey Tracker Table
        display_journey_df = journey_df.copy()
        # Drop internal underscore columns for UI display
        display_cols = [c for c in display_journey_df.columns if not c.startswith("_")]
        display_journey_df = display_journey_df[display_cols]

        st.dataframe(display_journey_df, width="stretch", height=450)

        # Download Journey Tracker Excel
        st.markdown("### 📥 Export Journey Tracker (Book 1)")
        j_down_col1, j_down_col2 = st.columns([1, 3])
        with j_down_col1:
            today_str = datetime.today().strftime('%d%b%Y')
            j_excel_bytes = pipeline.export_journey_tracker_to_excel_bytes(
                journey_df,
                email_rate=email_rate,
                wa_rate=wa_rate
            )
            st.download_button(
                label="📥 Download Journey Report (.xlsx)",
                data=j_excel_bytes,
                file_name=f"Journey_Report_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with j_down_col2:
            st.caption("Styled in Akasa corporate dark blue (`#1F4E79`) with live formulas: `=ROUND(rate*Sent,0)` and `=IFERROR(ROUND(CT_Revenue/Cost,0),-1)`.")

    # TAB 2: PERFORMANCE INSIGHTS
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

    # TAB 3: CONSOLIDATED CAMPAIGN REPORT
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
        st.markdown("### 📥 Export Consolidated Campaign Report")
        col_down1, col_down2 = st.columns([1, 3])
        with col_down1:
            excel_bytes = pipeline.export_to_excel_bytes(filtered_df)
            st.download_button(
                label="📥 Download final_campaign_report (.xlsx)",
                data=excel_bytes,
                file_name="final_campaign_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_down2:
            st.caption("Formatted with executive styling, auto-column widths, and clean blanks for non-applicable metrics.")

    # TAB 4: HYGIENE & AUDIT LOG
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
        - **Integrated Journey Tracker (Book 1)**: Computes cost modeling, AOV, CTR, Delivery %, and injects live Excel formulas for Cost and ROI.
        """)


if __name__ == "__main__":
    main()
