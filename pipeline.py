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
