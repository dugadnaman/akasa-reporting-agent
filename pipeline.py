"""Core data processing and aggregation pipeline for Akasa Air campaign reports.

Normalizes multi-channel MoEngage marketing exports (Email, WhatsApp, Push, Flows)
into a unified executive campaign attribution report.
"""

from __future__ import annotations

import ast
import io
import re
from typing import Any, BinaryIO, Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd


def normalize_campaign_tag(val: Any, untagged_label: str = "[]") -> str:
    """Normalize raw stringified tags into canonical sorted string representations.
    
    Resolves tag permutations like "['AMD', 'OneWayAbandoned_V3']" and
    "['OneWayAbandoned_V3', 'AMD']" to a single canonical group.
    """
    if pd.isna(val):
        return untagged_label
    val_str = str(val).strip()
    if not val_str or val_str in ("[]", "['']", '[""]', "set()"):
        return untagged_label
    try:
        parsed = ast.literal_eval(val_str)
        if isinstance(parsed, (list, tuple, set)):
            cleaned = sorted(str(x).strip() for x in parsed if str(x).strip())
            return str(cleaned) if cleaned else untagged_label
        return str(parsed)
    except Exception:
        items = re.findall(r"['\"](.*?)['\"]", val_str)
        if items:
            cleaned = sorted(set(i.strip() for i in items if i.strip()))
            return str(cleaned) if cleaned else untagged_label
        return val_str


def _read_df(file_source: Union[str, BinaryIO, bytes]) -> pd.DataFrame:
    """Read an Excel or CSV file into a pandas DataFrame."""
    if isinstance(file_source, bytes):
        file_source = io.BytesIO(file_source)
    try:
        return pd.read_excel(file_source)
    except Exception:
        if hasattr(file_source, "seek"):
            file_source.seek(0)
        return pd.read_csv(file_source)


def process_email_channel(
    file_source: Union[str, BinaryIO, bytes],
    normalize_tags: bool = True
) -> List[Dict[str, Any]]:
    """Extract and aggregate Email channel metrics."""
    df = _read_df(file_source)
    metrics = [
        "Total Sent", "Total Open", "Total clicks",
        "Goal 1 View Through Converted Users",
        "Goal 1 Click Through Converted Users",
        "Goal 1 Click Through Conversion Events",
        "Goal 1 Click Through Total Revenue"
    ]
    existing = [c for c in metrics if c in df.columns]
    if existing:
        active = df[(df[existing] != 0).any(axis=1)].copy()
    else:
        active = df.copy()

    tag_col = "Tag Category: Uncategorized" if "Tag Category: Uncategorized" in active.columns else active.columns[3]
    active["campaign"] = active[tag_col].apply(normalize_campaign_tag if normalize_tags else lambda x: str(x).strip())

    grp = active.groupby("campaign")[existing].sum()
    records = []
    for camp, row in grp.iterrows():
        records.append({
            "campaign": camp,
            "Channel": "Email",
            "Total Sent": int(row.get("Total Sent", 0)),
            "Total Open": int(row.get("Total Open", 0)),
            "Total Clicks": int(row.get("Total clicks", 0)),
            "Goal 1 View Through Converted Users": int(row.get("Goal 1 View Through Converted Users", 0)),
            "Goal 1 Click Through Converted Users": int(row.get("Goal 1 Click Through Converted Users", 0)),
            "Goal 1 Click Through Conversion Events": int(row.get("Goal 1 Click Through Conversion Events", 0)),
            "Goal 1 Click Through Total Revenue": float(row.get("Goal 1 Click Through Total Revenue", 0.0))
        })
    return records


def process_whatsapp_channel(
    file_source: Union[str, BinaryIO, bytes],
    normalize_tags: bool = True
) -> List[Dict[str, Any]]:
    """Extract and aggregate WhatsApp channel metrics."""
    df = _read_df(file_source)
    metrics = [
        "Total Sent", "Total Delivered", "Total Read", "Total clicks",
        "Goal 1 View Through Converted Users",
        "Goal 1 Click Through Converted Users",
        "Goal 1 Click Through Conversion Events",
        "Goal 1 Click Through Total Revenue"
    ]
    existing = [c for c in metrics if c in df.columns]
    if existing:
        active = df[(df[existing] != 0).any(axis=1)].copy()
    else:
        active = df.copy()

    tag_col = "Tag Category: Uncategorized" if "Tag Category: Uncategorized" in active.columns else active.columns[3]
    active["campaign"] = active[tag_col].apply(normalize_campaign_tag if normalize_tags else lambda x: str(x).strip())

    grp = active.groupby("campaign")[existing].sum()
    records = []
    for camp, row in grp.iterrows():
        records.append({
            "campaign": camp,
            "Channel": "WhatsApp",
            "Total Sent": int(row.get("Total Sent", 0)),
            "Total Delivered": int(row.get("Total Delivered", 0)),
            "Total Read": int(row.get("Total Read", 0)),
            "Total Clicks": int(row.get("Total clicks", 0)),
            "Goal 1 View Through Converted Users": int(row.get("Goal 1 View Through Converted Users", 0)),
            "Goal 1 Click Through Converted Users": int(row.get("Goal 1 Click Through Converted Users", 0)),
            "Goal 1 Click Through Conversion Events": int(row.get("Goal 1 Click Through Conversion Events", 0)),
            "Goal 1 Click Through Total Revenue": float(row.get("Goal 1 Click Through Total Revenue", 0.0))
        })
    return records


def process_app_push_channel(
    file_source: Union[str, BinaryIO, bytes],
    normalize_tags: bool = True
) -> List[Dict[str, Any]]:
    """Extract and aggregate App Push (Android + iOS combined) channel metrics."""
    df = _read_df(file_source)
    app_cols = [
        "Android Sent", "Ios Sent",
        "Android Impressions", "Ios Impressions",
        "Android Clicks", "Ios Clicks",
        "Goal 1 View Through Converted Users Android", "Goal 1 View Through Converted Users Ios",
        "Goal 1 Click Through Converted Users Android", "Goal 1 Click Through Converted Users Ios",
        "Goal 1 Click Through Conversion Events Android", "Goal 1 Click Through Conversion Events Ios",
        "Goal 1 Click Through Total Revenue Android", "Goal 1 Click Through Total Revenue Ios"
    ]
    existing = [c for c in app_cols if c in df.columns]
    if existing:
        active = df[(df[existing] != 0).any(axis=1)].copy()
    else:
        active = df.copy()

    tag_col = "Tag Category: Uncategorized" if "Tag Category: Uncategorized" in active.columns else "Tag Category: Default"
    active["campaign"] = active[tag_col].apply(normalize_campaign_tag if normalize_tags else lambda x: str(x).strip())

    grp = active.groupby("campaign")[existing].sum()
    records = []
    for camp, row in grp.iterrows():
        records.append({
            "campaign": camp,
            "Channel": "App Push",
            "Total Sent": int(row.get("Android Sent", 0) + row.get("Ios Sent", 0)),
            "Total Impressions": int(row.get("Android Impressions", 0) + row.get("Ios Impressions", 0)),
            "Total Clicks": int(row.get("Android Clicks", 0) + row.get("Ios Clicks", 0)),
            "Goal 1 View Through Converted Users": int(
                row.get("Goal 1 View Through Converted Users Android", 0) +
                row.get("Goal 1 View Through Converted Users Ios", 0)
            ),
            "Goal 1 Click Through Converted Users": int(
                row.get("Goal 1 Click Through Converted Users Android", 0) +
                row.get("Goal 1 Click Through Converted Users Ios", 0)
            ),
            "Goal 1 Click Through Conversion Events": int(
                row.get("Goal 1 Click Through Conversion Events Android", 0) +
                row.get("Goal 1 Click Through Conversion Events Ios", 0)
            ),
            "Goal 1 Click Through Total Revenue": float(
                row.get("Goal 1 Click Through Total Revenue Android", 0.0) +
                row.get("Goal 1 Click Through Total Revenue Ios", 0.0)
            )
        })
    return records


def process_web_push_channel(
    file_source: Union[str, BinaryIO, bytes],
    normalize_tags: bool = True
) -> List[Dict[str, Any]]:
    """Extract and aggregate Web Push channel metrics."""
    df = _read_df(file_source)
    web_cols = [
        "Web Sent", "Web Impressions", "Web Clicks",
        "Goal 1 View Through Converted Users Web",
        "Goal 1 Click Through Converted Users Web",
        "Goal 1 Click Through Conversion Events Web",
        "Goal 1 Click Through Total Revenue Web"
    ]
    existing = [c for c in web_cols if c in df.columns]
    if existing:
        active = df[(df[existing] != 0).any(axis=1)].copy()
    else:
        active = df.copy()

    tag_col = "Tag Category: Uncategorized" if "Tag Category: Uncategorized" in active.columns else "Tag Category: Default"
    active["campaign"] = active[tag_col].apply(normalize_campaign_tag if normalize_tags else lambda x: str(x).strip())

    grp = active.groupby("campaign")[existing].sum()
    records = []
    for camp, row in grp.iterrows():
        records.append({
            "campaign": camp,
            "Channel": "Web",
            "Total Sent": int(row.get("Web Sent", 0)),
            "Total Delivered": int(row.get("Web Impressions", 0)),
            "Total Clicks": int(row.get("Web Clicks", 0)),
            "Goal 1 View Through Converted Users": int(row.get("Goal 1 View Through Converted Users Web", 0)),
            "Goal 1 Click Through Converted Users": int(row.get("Goal 1 Click Through Converted Users Web", 0)),
            "Goal 1 Click Through Conversion Events": int(row.get("Goal 1 Click Through Conversion Events Web", 0)),
            "Goal 1 Click Through Total Revenue": float(row.get("Goal 1 Click Through Total Revenue Web", 0.0))
        })
    return records


def process_overall_flows(
    file_source: Union[str, BinaryIO, bytes],
    normalize_tags: bool = True
) -> List[Dict[str, Any]]:
    """Extract and aggregate Omnichannel Journey / Flows metrics."""
    df = _read_df(file_source)
    over_cols = [
        "Trips Started Total users",
        "Trips engaged",
        "Goal 1 View Through Total Conversion users (CVR)",
        "View Through Revenue (Primary Goal)"
    ]
    existing = [c for c in over_cols if c in df.columns]
    if existing:
        active = df[(df[existing] != 0).any(axis=1)].copy()
    else:
        active = df.copy()

    tag_col = "Tag Category: Uncategorized" if "Tag Category: Uncategorized" in active.columns else "Tag Category: Default"
    active["campaign"] = active[tag_col].apply(normalize_campaign_tag if normalize_tags else lambda x: str(x).strip())

    grp = active.groupby("campaign")[existing].sum()
    records = []
    for camp, row in grp.iterrows():
        records.append({
            "campaign": camp,
            "Channel": "overAll",
            "Trips Started Total users": int(row.get("Trips Started Total users", 0)),
            "Trips engaged": int(row.get("Trips engaged", 0)),
            "Goal 1 View Through Total Conversion users (CVR)": int(row.get("Goal 1 View Through Total Conversion users (CVR)", 0)),
            "View Through Revenue (Primary Goal)": float(row.get("View Through Revenue (Primary Goal)", 0.0))
        })
    return records


def generate_consolidated_report(
    email_file: Optional[Union[str, BinaryIO, bytes]] = None,
    whatsapp_file: Optional[Union[str, BinaryIO, bytes]] = None,
    push_file: Optional[Union[str, BinaryIO, bytes]] = None,
    overall_file: Optional[Union[str, BinaryIO, bytes]] = None,
    normalize_tags: bool = True
) -> pd.DataFrame:
    """Run all channel pipelines and consolidate into unified campaign report."""
    records: List[Dict[str, Any]] = []

    if email_file is not None:
        records.extend(process_email_channel(email_file, normalize_tags=normalize_tags))

    if whatsapp_file is not None:
        records.extend(process_whatsapp_channel(whatsapp_file, normalize_tags=normalize_tags))

    if push_file is not None:
        records.extend(process_app_push_channel(push_file, normalize_tags=normalize_tags))
        records.extend(process_web_push_channel(push_file, normalize_tags=normalize_tags))

    if overall_file is not None:
        records.extend(process_overall_flows(overall_file, normalize_tags=normalize_tags))

    if not records:
        return pd.DataFrame()

    combined_df = pd.DataFrame(records)

    col_order = [
        "campaign",
        "Channel",
        "Total Sent",
        "Total Open",
        "Total Delivered",
        "Total Read",
        "Total Impressions",
        "Total Clicks",
        "Goal 1 View Through Converted Users",
        "Goal 1 Click Through Converted Users",
        "Goal 1 Click Through Conversion Events",
        "Goal 1 Click Through Total Revenue",
        "Trips Started Total users",
        "Trips engaged",
        "Goal 1 View Through Total Conversion users (CVR)",
        "View Through Revenue (Primary Goal)"
    ]

    for c in col_order:
        if c not in combined_df.columns:
            combined_df[c] = np.nan

    grouped = combined_df.groupby(["campaign", "Channel"], as_index=False).sum()
    grouped = grouped[col_order]
    return grouped


def format_dataframe_for_export(df: pd.DataFrame) -> pd.DataFrame:
    """Format DataFrame for executive Excel presentation with blanked non-applicable metrics."""
    export_df = df.copy()
    overall_cols = [
        "Trips Started Total users",
        "Trips engaged",
        "Goal 1 View Through Total Conversion users (CVR)",
        "View Through Revenue (Primary Goal)"
    ]
    channel_cols = [
        "Total Sent",
        "Total Open",
        "Total Delivered",
        "Total Read",
        "Total Impressions",
        "Total Clicks",
        "Goal 1 View Through Converted Users",
        "Goal 1 Click Through Converted Users",
        "Goal 1 Click Through Conversion Events",
        "Goal 1 Click Through Total Revenue"
    ]

    is_overall = export_df["Channel"] == "overAll"
    export_df.loc[is_overall, channel_cols] = np.nan
    export_df.loc[~is_overall, overall_cols] = np.nan
    return export_df


def export_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Export the formatted DataFrame to an in-memory styled Excel workbook."""
    output = io.BytesIO()
    formatted = format_dataframe_for_export(df)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        formatted.to_excel(writer, index=False, sheet_name="Campaign Performance")
        ws = writer.sheets["Campaign Performance"]

        # Basic styling
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9")
        )

        for col_num, col in enumerate(ws.columns, 1):
            cell = col[0]
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = cell.column_letter
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            for row_cell in col[1:]:
                row_cell.border = thin_border
                val = row_cell.value
                if isinstance(val, (int, float)):
                    if "Revenue" in str(col[0].value):
                        row_cell.number_format = "#,##0.00"
                    else:
                        row_cell.number_format = "#,##0"

    output.seek(0)
    return output.getvalue()


DEFAULT_JOURNEY_MAPPINGS = {
    "['AMD', 'OneWayAbandoned_V3']": "One Way Abandonment - AMD",
    "['AMD', 'RoundTrip_V3']": "Round Trip Abandonment - AMD",
    "['Booking_Failed_Nudge']": "Booking Failed Nudge",
    "['Checkout_Step1_Abandonment_V3']": "Checkout Step 1 Abandonment",
    "['Domestic', 'OneWayAbandoned_V3']": "One Way Abandonment - Domestic",
    "['Domestic', 'RoundTrip_V3']": "Round Trip Abandonment - Domestic",
    "['Flexi_Promortion']": "Flexi Fare Promotion",
    "['Flight_Selected_Abandoned']": "Flight Selected Abandonment",
    "['Marriott']": "Marriott Partnership",
    "[]": "Untagged Campaigns"
}

JOURNEY_OUT_COLS = [
    "Journey Name", "Channels",
    "Trip Started", "Trip Engaged",
    "Sent", "Delivered/Impressions", "Delivery/Impression %",
    "Open/Read", "Clicks", "CTR",
    "VT Conv Users", "VT Total Conv users", "VT Revenue",
    "CT Conv Users", "CT Conv Events", "CT Revenue",
    "Avg Order Value", "Cost", "ROI"
]


def transform_to_journey_tracker(
    df_raw: pd.DataFrame,
    mappings: Optional[Dict[str, str]] = None,
    email_rate: float = 0.007,
    wa_rate: float = 0.007
) -> Tuple[pd.DataFrame, List[str]]:
    """Transform consolidated campaign report into the Akasa Journey Tracker format."""
    if mappings is None:
        mappings = DEFAULT_JOURNEY_MAPPINGS

    df = df_raw.copy()
    df["campaign"] = df["campaign"].astype(str)
    df["Channel"] = df["Channel"].replace({"Web": "Web Push"})
    df["jname"] = df["campaign"].map(mappings).fillna(df["campaign"])

    df_ov = df[df["Channel"] == "overAll"]
    df_ch = df[df["Channel"] != "overAll"]

    overall = {}
    for jn, grp in df_ov.groupby("jname"):
        overall[jn] = {
            "ts": float(grp.get("Trips Started Total users", pd.Series([0])).sum()),
            "te": float(grp.get("Trips engaged", pd.Series([0])).sum()),
            "vtc": float(grp.get("Goal 1 View Through Total Conversion users (CVR)", pd.Series([0])).sum()),
            "vtr": float(grp.get("View Through Revenue (Primary Goal)", pd.Series([0])).sum()),
        }

    channel_order = ["Email", "App Push", "Web Push", "WhatsApp"]
    cost_channels = {"Email", "WhatsApp"}

    def ch_metrics(grp, ch):
        clicks = float(grp.get("Total Clicks", grp.get("Total clicks", pd.Series([0]))).sum())

        if ch == "App Push":
            delivered = float(grp.get("Total Impressions", pd.Series([0])).sum())
        elif ch in ("Web Push", "WhatsApp"):
            delivered = float(grp.get("Total Delivered", pd.Series([0])).sum())
        else:
            delivered = 0.0

        if ch == "Email":
            open_read = float(grp.get("Total Open", pd.Series([0])).sum())
        elif ch == "WhatsApp":
            open_read = float(grp.get("Total Read", pd.Series([0])).sum())
        elif ch == "App Push":
            open_read = "-"
        else:
            open_read = None

        return {
            "sent": float(grp.get("Total Sent", pd.Series([0])).sum()),
            "delivered": delivered,
            "open_read": open_read,
            "clicks": clicks,
            "vt_cu": float(grp.get("Goal 1 View Through Converted Users", pd.Series([0])).sum()),
            "ct_cu": float(grp.get("Goal 1 Click Through Converted Users", pd.Series([0])).sum()),
            "ct_ce": float(grp.get("Goal 1 Click Through Conversion Events", pd.Series([0])).sum()),
            "ct_tr": float(grp.get("Goal 1 Click Through Total Revenue", pd.Series([0])).sum()),
        }

    rows = []
    warnings = []
    seen = {}
    for jn in df_ch["jname"]:
        seen.setdefault(jn, True)

    for jn in seen:
        if not any(mappings.get(t) == jn for t in mappings) and jn not in mappings:
            warnings.append(f"Unmapped tag used as journey name: {jn}")

        jgrp = df_ch[df_ch["jname"] == jn]
        ov = overall.get(jn, {"ts": 0, "te": 0, "vtc": 0, "vtr": 0})
        first = True

        for ch in channel_order:
            cg = jgrp[jgrp["Channel"] == ch]
            if cg.empty:
                continue

            m = ch_metrics(cg, ch)
            if (m["sent"] == 0 and m["delivered"] == 0 and
                    m["clicks"] == 0 and m["vt_cu"] == 0 and m["ct_cu"] == 0):
                continue

            sent = int(m["sent"])
            delivered = int(m["delivered"])
            clicks = int(m["clicks"])
            vt_cu = int(m["vt_cu"])
            ct_cu = int(m["ct_cu"])
            ct_ce = int(m["ct_ce"])
            ct_tr = int(m["ct_tr"])
            del_pct = round(delivered / sent, 4) if sent else 0
            ctr = round(clicks / sent, 4) if sent else 0
            aov = round(ct_tr / ct_ce, 0) if ct_ce else 0

            rate = email_rate if ch == "Email" else (wa_rate if ch == "WhatsApp" else 0)
            cost = round(rate * sent, 2) if ch in cost_channels else None
            roi = round(ct_tr / cost, 2) if cost and cost > 0 else (None if ch not in cost_channels else -1)

            rows.append({
                "Journey Name": jn if first else "",
                "Channels": ch,
                "Trip Started": int(ov["ts"]) if first and ov["ts"] else None,
                "Trip Engaged": int(ov["te"]) if first and ov["te"] else None,
                "Sent": sent,
                "Delivered/Impressions": delivered,
                "Delivery/Impression %": del_pct,
                "Open/Read": m["open_read"],
                "Clicks": clicks,
                "CTR": ctr,
                "VT Conv Users": vt_cu,
                "VT Total Conv users": int(ov["vtc"]) if first and ov["vtc"] else None,
                "VT Revenue": int(ov["vtr"]) if first and ov["vtr"] else None,
                "CT Conv Users": ct_cu,
                "CT Conv Events": ct_ce,
                "CT Revenue": ct_tr,
                "Avg Order Value": int(aov),
                "Cost": cost,
                "ROI": roi,
                "_ch": ch,
                "_er": email_rate,
                "_wr": wa_rate
            })
            first = False

    return pd.DataFrame(rows), warnings


def export_journey_tracker_to_excel_bytes(
    rows_df: Union[pd.DataFrame, List[Dict[str, Any]]],
    email_rate: float = 0.007,
    wa_rate: float = 0.007
) -> bytes:
    """Export Journey Tracker DataFrame to styled Excel with formulas and formatting."""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Journey Report"

    col_widths = {
        "Journey Name": 36, "Channels": 12,
        "Trip Started": 13, "Trip Engaged": 13,
        "Sent": 10, "Delivered/Impressions": 21, "Delivery/Impression %": 21,
        "Open/Read": 10, "Clicks": 8, "CTR": 8,
        "VT Conv Users": 14, "VT Total Conv users": 19, "VT Revenue": 15,
        "CT Conv Users": 14, "CT Conv Events": 15, "CT Revenue": 13,
        "Avg Order Value": 16, "Cost": 10, "ROI": 8,
    }

    header_fill = PatternFill("solid", start_color="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=9)
    body_font = Font(name="Arial", size=9)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    odd_fill = PatternFill("solid", start_color="DCE6F1")
    even_fill = PatternFill("solid", start_color="FFFFFF")
    thin = Border(
        left=Side(style="thin", color="B8CCE4"),
        right=Side(style="thin", color="B8CCE4"),
        top=Side(style="thin", color="B8CCE4"),
        bottom=Side(style="thin", color="B8CCE4"),
    )

    for c, col in enumerate(JOURNEY_OUT_COLS, 1):
        cell = ws.cell(row=1, column=c, value=col)
        cell.font, cell.fill, cell.alignment, cell.border = (
            header_font, header_fill, center, thin
        )
        ws.column_dimensions[get_column_letter(c)].width = col_widths.get(col, 12)

    sent_let = get_column_letter(JOURNEY_OUT_COLS.index("Sent") + 1)
    ctrev_let = get_column_letter(JOURNEY_OUT_COLS.index("CT Revenue") + 1)
    cost_let = get_column_letter(JOURNEY_OUT_COLS.index("Cost") + 1)

    inr_cols = {"VT Revenue", "CT Revenue", "Avg Order Value", "Cost"}
    inr_fmt = "[$\u20b9-4009]#,##,##0;[$\u20b9-4009]-"
    pct_cols = {"Delivery/Impression %", "CTR"}
    cost_channels = {"Email", "WhatsApp"}

    records = rows_df.to_dict(orient="records") if isinstance(rows_df, pd.DataFrame) else rows_df

    for r, row in enumerate(records, 2):
        fill = odd_fill if r % 2 == 1 else even_fill
        ch = row.get("_ch", row.get("Channels"))
        rate = email_rate if ch == "Email" else wa_rate

        for c, col in enumerate(JOURNEY_OUT_COLS, 1):
            if col == "Cost":
                val = f"=ROUND({rate}*{sent_let}{r},0)" if ch in cost_channels else None
            elif col == "ROI":
                val = f"=IFERROR(ROUND({ctrev_let}{r}/{cost_let}{r},0),-1)" if ch in cost_channels else None
            else:
                val = row.get(col)
                if val == "" or pd.isna(val):
                    val = None

            cell = ws.cell(row=r, column=c, value=val)
            cell.font = body_font
            cell.fill = fill
            cell.border = thin
            cell.alignment = left if col in ("Journey Name", "Channels") else center

            if col in inr_cols and isinstance(val, (int, float)):
                cell.number_format = inr_fmt
            elif col in pct_cols and isinstance(val, (int, float)):
                cell.number_format = "0.00%"
            elif isinstance(val, (int, float)):
                cell.number_format = "#,##0"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
