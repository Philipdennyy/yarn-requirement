import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Yarn Requirement & Purchase Analysis",
    layout="wide"
)


# ============================================================
# BASIC HELPERS
# ============================================================

def find_sheet_name(sheet_names, keyword):
    keyword = keyword.lower()

    for name in sheet_names:
        if keyword in name.lower():
            return name

    return None


def clean_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def safe_float(value, default=0.0):
    """Safely convert a value to float for calculations."""
    try:
        if pd.isna(value):
            return default
    except (TypeError, ValueError):
        pass

    try:
        if isinstance(value, str):
            value = value.replace(",", "").strip()
            if value == "":
                return default
        return float(value)
    except (TypeError, ValueError):
        return default


# ============================================================
# YARN NORMALIZATION
# ============================================================

def normalize_yarn_description(description):

    if pd.isna(description):
        return ""

    text = str(description).upper().strip()

    # BOBBIN / BOBIN treated as BB
    text = text.replace("BOBBIN", "BB")
    text = text.replace("BOBIN", "BB")

    # First numeric value only
    number_match = re.search(
        r"\d+(?:\.\d+)?",
        text
    )

    if number_match:

        number = float(number_match.group())

        if number.is_integer():
            number_text = str(int(number))
        else:
            number_text = str(number)

        text_before_number = text[
            :number_match.start()
        ]

    else:

        number_text = ""
        text_before_number = text

    words = re.findall(
        r"[A-Z]+",
        text_before_number
    )

    ignored_words = {
        "BB",
        "BOBIN",
        "BOBBIN",
        "DW",
        "NW"
    }

    valid_words = []

    for word in words:

        if word in ignored_words:
            continue

        if len(word) < 3:
            continue

        valid_words.append(word)

    yarn_name = " ".join(valid_words)

    if number_text:
        return f"{yarn_name} {number_text}".strip()

    return yarn_name.strip()


# ============================================================
# PO BASE NUMBER
# ============================================================

def normalize_po(po):

    """
    Treat:
        2000033576
        2000033576A
        2000033576B
        2000033576C

    as the same PO group.

    Only trailing alphabetic suffixes are removed.
    Original PO values remain unchanged in Excel.
    """

    if pd.isna(po):
        return ""

    text = str(po).strip()

    if not text or text.lower() == "nan":
        return ""

    # Remove trailing alphabetic suffix
    return re.sub(
        r"[A-Za-z]+$",
        "",
        text
    )


# ============================================================
# PREPARE FIRM PLAN
# ============================================================

def prepare_firm_plan(df):

    data = clean_columns(df)

    required = [
        "Description",
        "Expected Quantity",
        "Prod. Order No.",
        "PO"
    ]

    missing = [
        c for c in required
        if c not in data.columns
    ]

    if missing:
        raise ValueError(
            "Firm Plan is missing: "
            + ", ".join(missing)
        )

    # Standardize Firm Plan descriptions before grouping.
    # Remove spaces around hyphens and replace BOBBIN with BB.
    data["Description"] = (
        data["Description"]
        .astype(str)
        .str.strip()
        .str.replace(r"\s*-\s*", "-", regex=True)
        .str.replace(r"\bbobbin\b", "BB", regex=True, case=False)
    )

    data["_Original Row"] = range(len(data))

    data["_Automatic Key"] = (
        data["Description"]
        .apply(normalize_yarn_description)
    )

    data["Expected Quantity"] = pd.to_numeric(
        data["Expected Quantity"],
        errors="coerce"
    ).fillna(0)

    # Base PO used only for production allocation
    data["_Base PO"] = (
        data["PO"]
        .apply(normalize_po)
    )

    return data


# ============================================================
# PREPARE BIN
# ============================================================

def prepare_bin(df):

    data = clean_columns(df)

    required = [
        "Location Code",
        "Item Description",
        "Quantity",
        "Lot No"
    ]

    missing = [
        c for c in required
        if c not in data.columns
    ]

    if missing:
        raise ValueError(
            "Bin is missing: "
            + ", ".join(missing)
        )

    data["Location Code"] = (
        data["Location Code"]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    data = data[
        data["Location Code"].isin(
            ["STORES", "WIP"]
        )
    ].copy()

    data["_Automatic Key"] = (
        data["Item Description"]
        .apply(normalize_yarn_description)
    )

    data["Quantity"] = pd.to_numeric(
        data["Quantity"],
        errors="coerce"
    ).fillna(0)

    data["_Lot ID"] = (
        data["Lot No"]
        .astype(str)
        .str.strip()
    )

    empty_lot = (
        data["_Lot ID"].isna()
        |
        (data["_Lot ID"] == "")
        |
        (
            data["_Lot ID"].str.upper()
            == "NAN"
        )
    )

    data.loc[
        empty_lot,
        "_Lot ID"
    ] = (
        "BIN_ROW_"
        +
        data.loc[
            empty_lot
        ].index.astype(str)
    )

    return data


# ============================================================
# PREPARE PENDING PO
# ============================================================

def prepare_pending_po(df):

    data = clean_columns(df)

    required = [
        "Description",
        "Outstanding Quantity"
    ]

    missing = [
        c for c in required
        if c not in data.columns
    ]

    if missing:
        raise ValueError(
            "Pending PO is missing: "
            + ", ".join(missing)
        )

    data["_Automatic Key"] = (
        data["Description"]
        .apply(normalize_yarn_description)
    )

    data["Outstanding Quantity"] = pd.to_numeric(
        data["Outstanding Quantity"],
        errors="coerce"
    ).fillna(0)

    if "Document No." in data.columns:

        data["_Lot ID"] = (
            data["Document No."]
            .astype(str)
            .str.strip()
        )

    else:

        data["_Lot ID"] = ""

    empty_lot = (
        data["_Lot ID"].isna()
        |
        (data["_Lot ID"] == "")
        |
        (
            data["_Lot ID"].str.upper()
            == "NAN"
        )
    )

    data.loc[
        empty_lot,
        "_Lot ID"
    ] = (
        "PENDING_PO_ROW_"
        +
        data.loc[
            empty_lot
        ].index.astype(str)
    )

    return data


# ============================================================
# CACHED EXCEL PROCESSING
# ============================================================

@st.cache_data(show_spinner=False)
def process_workbook(file_bytes):

    excel = BytesIO(file_bytes)

    xl = pd.ExcelFile(excel)

    sheet_names = xl.sheet_names

    firm_sheet = find_sheet_name(
        sheet_names,
        "Firm Plan"
    )

    bin_sheet = find_sheet_name(
        sheet_names,
        "Bin"
    )

    pending_sheet = find_sheet_name(
        sheet_names,
        "Pending PO"
    )

    missing = []

    if not firm_sheet:
        missing.append("Firm Plan")

    if not bin_sheet:
        missing.append("Bin")

    if not pending_sheet:
        missing.append("Pending PO")

    if missing:
        raise ValueError(
            "Missing required sheet(s): "
            + ", ".join(missing)
        )

    firm_raw = pd.read_excel(
        excel,
        sheet_name=firm_sheet
    )

    bin_raw = pd.read_excel(
        excel,
        sheet_name=bin_sheet
    )

    pending_raw = pd.read_excel(
        excel,
        sheet_name=pending_sheet
    )

    firm_plan = prepare_firm_plan(
        firm_raw
    )

    bin_data = prepare_bin(
        bin_raw
    )

    pending_po = prepare_pending_po(
        pending_raw
    )

    return (
        firm_plan,
        bin_data,
        pending_po
    )


# ============================================================
# BUILD LOT LOOKUPS
# ============================================================

@st.cache_data(show_spinner=False)
def build_bin_lot_lookup(bin_data):

    lookup = {}

    grouped = (
        bin_data
        .groupby(
            [
                "_Automatic Key",
                "Location Code",
                "_Lot ID"
            ],
            sort=False
        )
        .agg(
            Quantity=(
                "Quantity",
                "sum"
            ),
            Description=(
                "Item Description",
                "first"
            )
        )
        .reset_index()
    )

    grouped = grouped.sort_values(
        [
            "_Automatic Key",
            "Location Code",
            "Quantity"
        ],
        ascending=[
            True,
            True,
            False
        ]
    )

    for _, row in grouped.iterrows():

        key = (
            row["_Automatic Key"],
            row["Location Code"]
        )

        lot = str(
            row["_Lot ID"]
        ).strip()

        option = {
            "lot": lot,
            "description": str(
                row["Description"]
            ).strip(),
            "quantity": float(
                row["Quantity"]
            )
        }

        lookup.setdefault(
            key,
            []
        ).append(option)

    return lookup


@st.cache_data(show_spinner=False)
def build_pending_lookup(pending_po):

    lookup = {}

    grouped = (
        pending_po
        .groupby(
            [
                "_Automatic Key",
                "_Lot ID"
            ],
            sort=False
        )
        .agg(
            Quantity=(
                "Outstanding Quantity",
                "sum"
            ),
            Description=(
                "Description",
                "first"
            )
        )
        .reset_index()
    )

    grouped = grouped.sort_values(
        [
            "_Automatic Key",
            "Quantity"
        ],
        ascending=[
            True,
            False
        ]
    )

    for _, row in grouped.iterrows():

        key = row[
            "_Automatic Key"
        ]

        lot = str(
            row["_Lot ID"]
        ).strip()

        lookup.setdefault(
            key,
            []
        ).append(
            {
                "lot": lot,
                "description": str(
                    row["Description"]
                ).strip(),
                "quantity": float(
                    row["Quantity"]
                )
            }
        )

    return lookup


# ============================================================
# MATCHING CHECKBOX STATE
# ============================================================

def update_matching_selection(description, category, lot, widget_key):
    selection = st.session_state.match_selections.setdefault(
        description,
        {"stores_lots": [], "wip_lots": [], "pending_lots": []}
    )
    selected = bool(st.session_state.get(widget_key, False))
    lots = selection.setdefault(category, [])
    if selected and lot not in lots:
        lots.append(lot)
    elif not selected and lot in lots:
        lots.remove(lot)


# ============================================================
# APPLY MATCHING
# ============================================================

def apply_matching(
    firm_plan,
    bin_data,
    pending_po,
    selections
):

    firm_plan = firm_plan.copy()
    bin_data = bin_data.copy()
    pending_po = pending_po.copy()

    firm_plan["_Matching Key"] = (
        firm_plan["Description"]
        .astype(str)
        .str.strip()
    )

    bin_data["_Matching Key"] = None
    pending_po["_Matching Key"] = None

    for description, selection in selections.items():

        automatic_key = (
            normalize_yarn_description(
                description
            )
        )

        stores_lots = selection.get(
            "stores_lots",
            []
        )

        wip_lots = selection.get(
            "wip_lots",
            []
        )

        pending_lots = selection.get(
            "pending_lots",
            []
        )

        # STORES

        if stores_lots:

            mask = (
                bin_data[
                    "_Automatic Key"
                ]
                == automatic_key
            ) & (
                bin_data[
                    "Location Code"
                ]
                == "STORES"
            ) & (
                bin_data[
                    "_Lot ID"
                ]
                .astype(str)
                .str.strip()
                .isin(
                    [
                        str(x).strip()
                        for x in stores_lots
                    ]
                )
            )

            bin_data.loc[
                mask,
                "_Matching Key"
            ] = description

        # WIP

        if wip_lots:

            mask = (
                bin_data[
                    "_Automatic Key"
                ]
                == automatic_key
            ) & (
                bin_data[
                    "Location Code"
                ]
                == "WIP"
            ) & (
                bin_data[
                    "_Lot ID"
                ]
                .astype(str)
                .str.strip()
                .isin(
                    [
                        str(x).strip()
                        for x in wip_lots
                    ]
                )
            )

            bin_data.loc[
                mask,
                "_Matching Key"
            ] = description

        # PENDING PO

        if pending_lots:

            mask = (
                pending_po[
                    "_Automatic Key"
                ]
                == automatic_key
            ) & (
                pending_po[
                    "_Lot ID"
                ]
                .astype(str)
                .str.strip()
                .isin(
                    [
                        str(x).strip()
                        for x in pending_lots
                    ]
                )
            )

            pending_po.loc[
                mask,
                "_Matching Key"
            ] = description

    return (
        firm_plan,
        bin_data,
        pending_po
    )


# ============================================================
# CREATE REPORT
# ============================================================

def create_report(
    firm_plan,
    bin_data,
    pending_po
):

    grouped = (
        firm_plan
        .groupby(
            "Description",
            sort=False,
            as_index=False
        )
        .agg(
            {
                "Expected Quantity": "sum",
                "_Matching Key": "first"
            }
        )
    )

    grouped = grouped.rename(
        columns={
            "Expected Quantity":
                "Firm Plan"
        }
    )

    # STORES

    stores = (
        bin_data[
            (
                bin_data[
                    "Location Code"
                ]
                == "STORES"
            )
            &
            bin_data[
                "_Matching Key"
            ].notna()
        ]
        .groupby(
            "_Matching Key"
        )["Quantity"]
        .sum()
        .to_dict()
    )

    # WIP

    wip = (
        bin_data[
            (
                bin_data[
                    "Location Code"
                ]
                == "WIP"
            )
            &
            bin_data[
                "_Matching Key"
            ].notna()
        ]
        .groupby(
            "_Matching Key"
        )["Quantity"]
        .sum()
        .to_dict()
    )

    # PURCHASE

    purchase = (
        pending_po[
            pending_po[
                "_Matching Key"
            ].notna()
        ]
        .groupby(
            "_Matching Key"
        )[
            "Outstanding Quantity"
        ]
        .sum()
        .to_dict()
    )

    rows = []

    for _, row in grouped.iterrows():

        description = row[
            "Description"
        ]

        key = row[
            "_Matching Key"
        ]

        firm_qty = float(
            row["Firm Plan"]
        )

        wastage = firm_qty * 0.20

        requirement = (
            firm_qty
            + wastage
        )

        stores_qty = float(
            stores.get(
                key,
                0
            )
        )

        wip_qty = float(
            wip.get(
                key,
                0
            )
        )

        purchase_qty = float(
            purchase.get(
                key,
                0
            )
        )

        total_stock = (
            stores_qty
            + wip_qty
        )

        rows.append(
            {
                "Description":
                    description,

                "_Matching Key":
                    key,

                "Firm Plan":
                    firm_qty,

                "Released":
                    None,

                "Wastage 20%":
                    wastage,

                "Requirement":
                    requirement,

                "Safety Stock":
                    0.0,

                "Gross Requirement":
                    requirement,

                "Stores":
                    stores_qty,

                "WIP":
                    wip_qty,

                "Production":
                    None,

                "Total Stock":
                    total_stock,

                "Diff":
                    requirement
                    - total_stock,

                "Purchase Qty":
                    purchase_qty,

                "Requirement PO":
                    requirement
                    - total_stock
                    - purchase_qty,

                "Manual Purchase Qty":
                    0.0
            }
        )

    return pd.DataFrame(rows)


# ============================================================
# PRODUCTION ALLOCATION
# ============================================================

def allocate_production(
    firm_plan,
    bin_data,
    pending_po,
    report_df
):
    """
    For each Description in the report:
    1. Total Stock = Stores + WIP.
    2. Filter Firm Plan by the exact Description.
    3. Group Firm Plan by base PO and sum Expected Quantity.
    4. Process grouped POs in original Firm Plan order.
    5. If the PO quantity fits in the remaining stock, weave it.
    6. The first PO that does not fit and every following PO go to Leftout.
    7. Remaining Quantity is the Expected Quantity sum in Leftout for that Description.
    """

    firm_plan = firm_plan.copy()
    report_df = report_df.copy()

    if "_Original Row" not in firm_plan.columns:
        firm_plan["_Original Row"] = range(len(firm_plan))

    if "_Base PO" not in firm_plan.columns:
        firm_plan["_Base PO"] = firm_plan["PO"].apply(normalize_po)

    # Preserve exact Firm Plan Description for the allocation logic.
    firm_plan["_Description Exact"] = firm_plan["Description"].astype(str).str.strip()

    woven_quantity = {}
    woven_orders = {}
    leftout_pairs = set()

    # Process descriptions in the same order as the report.
    for idx, report_row in report_df.iterrows():

        description = str(report_row.get("Description", "")).strip()

        if not description or description.lower() == "nan":
            continue

        total_stock = (
            safe_float(report_row.get("Stores"))
            + safe_float(report_row.get("WIP"))
        )

        purchase_qty = safe_float(report_row.get("Purchase Qty"))

        available_stock = max(
            total_stock + purchase_qty,
            0.0
        )

        description_fp = firm_plan[
            firm_plan["_Description Exact"] == description
        ].copy().sort_values("_Original Row")

        if description_fp.empty:
            report_df.at[idx, "Woven Quantity"] = 0.0
            report_df.at[idx, "Woven Count"] = 0
            report_df.at[idx, "Remaining Quantity"] = 0.0
            continue

        # Group by base PO while preserving first occurrence order.
        po_groups = []
        seen_pos = set()

        for _, fp_row in description_fp.iterrows():
            base_po = str(fp_row.get("_Base PO", "")).strip()
            if not base_po or base_po.lower() == "nan":
                continue

            if base_po not in seen_pos:
                seen_pos.add(base_po)
                po_groups.append(base_po)

        remaining_stock = available_stock
        woven_qty = 0.0
        woven_prod_orders = set()
        stopped = False

        for base_po in po_groups:

            po_df = description_fp[
                description_fp["_Base PO"] == base_po
            ].copy()

            po_qty = safe_float(po_df["Expected Quantity"].sum())

            if po_qty <= 0:
                continue

            if not stopped and po_qty <= remaining_stock + 1e-9:
                # Complete PO is woven.
                remaining_stock -= po_qty
                woven_qty += po_qty

                for prod in po_df["Prod. Order No."]:
                    if pd.isna(prod):
                        continue
                    prod = str(prod).strip()
                    if prod:
                        woven_prod_orders.add(prod)
            else:
                # First PO that cannot fit + all following POs = Leftout.
                stopped = True
                leftout_pairs.add((description, base_po))

        woven_quantity[description] = woven_qty
        woven_orders[description] = woven_prod_orders

    # --------------------------------------------------------
    # REPORT RESULTS FROM LEFT OUT
    # --------------------------------------------------------

    report_df["Woven Quantity"] = 0.0
    report_df["Woven Count"] = 0
    report_df["Remaining Quantity"] = 0.0

    for idx, report_row in report_df.iterrows():

        description = str(report_row.get("Description", "")).strip()

        leftout_mask = firm_plan.apply(
            lambda r: (
                r["_Description Exact"],
                r["_Base PO"]
            ) in leftout_pairs,
            axis=1
        )

        leftout_for_description = firm_plan[
            leftout_mask
            & (firm_plan["_Description Exact"] == description)
        ]

        leftout_qty = safe_float(
            leftout_for_description["Expected Quantity"].sum()
        )

        report_df.at[idx, "Woven Quantity"] = woven_quantity.get(description, 0.0)
        report_df.at[idx, "Woven Count"] = len(woven_orders.get(description, set()))
        report_df.at[idx, "Remaining Quantity"] = leftout_qty

    # --------------------------------------------------------
    # LEFT OUT SHEET DATA
    # --------------------------------------------------------

    leftout_mask = firm_plan.apply(
        lambda r: (
            r["_Description Exact"],
            r["_Base PO"]
        ) in leftout_pairs,
        axis=1
    )

    leftout_firm_plan = (
        firm_plan[leftout_mask]
        .copy()
        .sort_values("_Original Row")
    )

    # Remove helper columns from the Leftout output.
    leftout_output = leftout_firm_plan.drop(
        columns=["_Description Exact"],
        errors="ignore"
    )

    return report_df, leftout_output


# ============================================================
# FINAL CALCULATION
# ============================================================

def calculate_final_report(df, leftout_firm_plan=None):

    df = df.copy()

    if leftout_firm_plan is not None and not leftout_firm_plan.empty:
        leftout_counts = (leftout_firm_plan.assign(_Description_Count_Key=leftout_firm_plan["Description"].astype(str).str.strip()).groupby("_Description_Count_Key").size().to_dict())
        df["Remaining No. of Rolls/Orders"] = df["Description"].astype(str).str.strip().map(leftout_counts).fillna(0).astype(int)
    else:
        df["Remaining No. of Rolls/Orders"] = 0

    numeric_columns = [
        "Firm Plan",
        "Stores",
        "WIP",
        "Purchase Qty",
        "Safety Stock",
        "Manual Purchase Qty",
        "Woven Quantity",
        "Woven Count",
        "Remaining Quantity"
    ]

    for column in numeric_columns:

        if column in df.columns:

            df[column] = pd.to_numeric(
                df[column],
                errors="coerce"
            ).fillna(0)

    # --------------------------------------------------------
    # BASIC REQUIREMENT
    # --------------------------------------------------------

    df["Wastage 20%"] = (
        df["Firm Plan"]
        * 0.20
    )

    df["Requirement"] = (
        df["Firm Plan"]
        + df["Wastage 20%"]
    )

    df["Gross Requirement"] = (
        df["Requirement"]
        + df["Safety Stock"]
    )

    # --------------------------------------------------------
    # STOCK
    # --------------------------------------------------------

    df["Total Stock"] = (
        df["Stores"]
        + df["WIP"]
    )

    df["Diff"] = (
        df["Gross Requirement"]
        - df["Total Stock"]
    )

    df["Requirement PO"] = (
        df["Diff"]
        - df["Purchase Qty"]
    )

    # --------------------------------------------------------
    # NEW PURCHASE REQUIREMENT
    #
    # Remaining Quantity
    # + 20% of Remaining Quantity
    # + Manual Purchase Quantity
    # --------------------------------------------------------

    df[
        "New Purchase Requirement"
    ] = (
        df["Remaining Quantity"]
        +
        (
            df["Remaining Quantity"]
            * 0.20
        )
        +
        df["Manual Purchase Qty"]
    )

    # --------------------------------------------------------
    # REMARKS
    # --------------------------------------------------------

    remarks = []

    for _, row in df.iterrows():

        remarks.append(
            f"{int(row['Woven Count'])} "
            f"rolls/orders woven with available yarn. "
            f"Woven quantity = "
            f"{row['Woven Quantity']:.2f}. "
            f"Remaining quantity = "
            f"{row['Remaining Quantity']:.2f}. "
            f"Remaining no. of rolls/orders = "
            f"{int(row['Remaining No. of Rolls/Orders'])}. "
            f"New purchase requirement = "
            f"{row['New Purchase Requirement']:.2f}."
        )

    df["Remarks"] = remarks

    return df


# ============================================================
# OUTPUT EXCEL
# ============================================================

def create_output_workbook(
    file_bytes,
    report_df,
    pending_firm_plan
):

    wb = load_workbook(
        BytesIO(file_bytes)
    )

    # Update the original Firm Plan sheet as soon as the workbook is processed.
    firm_sheet_name = find_sheet_name(
        wb.sheetnames,
        "Firm Plan"
    )

    if firm_sheet_name:
        ws_firm = wb[firm_sheet_name]
        headers = {
            str(cell.value).strip().lower(): cell.column
            for cell in ws_firm[1]
            if cell.value is not None
        }
        desc_col = headers.get("description")

        if desc_col:
            for row in range(2, ws_firm.max_row + 1):
                cell = ws_firm.cell(row, desc_col)
                if isinstance(cell.value, str):
                    cell.value = re.sub(
                        r"\bbobbin\b",
                        "BB",
                        cell.value,
                        flags=re.IGNORECASE
                    )

    # --------------------------------------------------------
    # REQUIREMENT REPORT
    # --------------------------------------------------------

    if "Requirement Report" in wb.sheetnames:

        del wb[
            "Requirement Report"
        ]

    ws = wb.create_sheet(
        "Requirement Report"
    )

    columns = [
        "Description",
        "Firm Plan",
        "Released",
        "Wastage 20%",
        "Requirement",
        "Safety Stock",
        "Gross Requirement",
        "Stores",
        "WIP",
        "Production",
        "Total Stock",
        "Diff",
        "Purchase Qty",
        "Requirement PO",
        "New Purchase Requirement",
        "Remarks"
    ]

    for col_num, column in enumerate(
        columns,
        1
    ):

        ws.cell(
            1,
            col_num,
            column
        )

    for row_num, (_, row) in enumerate(
        report_df.iterrows(),
        2
    ):

        for col_num, column in enumerate(
            columns,
            1
        ):

            value = row.get(
                column
            )

            if pd.isna(value):
                value = None

            ws.cell(
                row_num,
                col_num,
                value
            )

    # --------------------------------------------------------
    # LEFT OUT
    # --------------------------------------------------------

    if "leftout" in wb.sheetnames:
        del wb["leftout"]

    ws_pending = wb.create_sheet("leftout")

    pending_output = (
        pending_firm_plan
        .drop(
            columns=[
                "_Original Row",
                "_Automatic Key",
                "_Matching Key",
                "_Base PO"
            ],
            errors="ignore"
        )
    )

    for col_num, column in enumerate(
        pending_output.columns,
        1
    ):

        ws_pending.cell(
            1,
            col_num,
            column
        )

    for row_num, (_, row) in enumerate(
        pending_output.iterrows(),
        2
    ):

        for col_num, column in enumerate(
            pending_output.columns,
            1
        ):

            value = row[column]

            if pd.isna(value):
                value = None

            ws_pending.cell(
                row_num,
                col_num,
                value
            )

    # --------------------------------------------------------
    # WIDTH
    # --------------------------------------------------------

    for sheet in [
        ws,
        ws_pending
    ]:

        for column_cells in sheet.columns:

            max_length = 0

            letter = (
                column_cells[
                    0
                ].column_letter
            )

            for cell in column_cells:

                if cell.value is not None:

                    max_length = max(
                        max_length,
                        len(
                            str(
                                cell.value
                            )
                        )
                    )

            sheet.column_dimensions[
                letter
            ].width = min(
                max_length + 2,
                50
            )

    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

    output = BytesIO()

    wb.save(output)

    output.seek(0)

    return output


# ============================================================
# SESSION STATE
# ============================================================

if "file_bytes" not in st.session_state:
    st.session_state.file_bytes = None

if "firm_plan" not in st.session_state:
    st.session_state.firm_plan = None

if "bin_data" not in st.session_state:
    st.session_state.bin_data = None

if "pending_po" not in st.session_state:
    st.session_state.pending_po = None

if "matching_confirmed" not in st.session_state:
    st.session_state.matching_confirmed = False

if "match_selections" not in st.session_state:
    st.session_state.match_selections = {}

if "report_df" not in st.session_state:
    st.session_state.report_df = None

if "pending_firm_plan" not in st.session_state:
    st.session_state.pending_firm_plan = None

if "calculated" not in st.session_state:
    st.session_state.calculated = False

if "yarn_search" not in st.session_state:
    st.session_state.yarn_search = ""


# ============================================================
# TITLE
# ============================================================

st.title(
    "🧶 Yarn Requirement & Purchase Analysis"
)


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "Upload Excel Workbook",
    type=["xlsx"]
)


if uploaded_file:

    file_bytes = uploaded_file.getvalue()

    # ========================================================
    # NEW FILE
    # ========================================================

    if (
        st.session_state.file_bytes
        != file_bytes
    ):

        st.session_state.file_bytes = file_bytes

        st.session_state.matching_confirmed = False

        st.session_state.match_selections = {}

        st.session_state.report_df = None

        st.session_state.pending_firm_plan = None

        st.session_state.calculated = False

        st.session_state.yarn_search = ""
        st.session_state.selected_yarn_description = None
        st.session_state.sidebar_yarn_search = ""

        with st.spinner(
            "Reading workbook..."
        ):

            try:

                (
                    firm_plan,
                    bin_data,
                    pending_po
                ) = process_workbook(
                    file_bytes
                )

                st.session_state.firm_plan = (
                    firm_plan
                )

                st.session_state.bin_data = (
                    bin_data
                )

                st.session_state.pending_po = (
                    pending_po
                )

            except Exception as e:

                st.error(
                    str(e)
                )

                st.stop()

    # ========================================================
    # LOAD DATA
    # ========================================================

    firm_plan = (
        st.session_state.firm_plan
    )

    bin_data = (
        st.session_state.bin_data
    )

    pending_po = (
        st.session_state.pending_po
    )

    # Build cached lookups
    bin_lookup = build_bin_lot_lookup(
        bin_data
    )

    pending_lookup = build_pending_lookup(
        pending_po
    )

    all_descriptions = (
        firm_plan[
            "Description"
        ]
        .drop_duplicates()
        .tolist()
    )

    # ========================================================
    # STEP 1
    # ========================================================

    st.subheader(
        "1️⃣ Yarn Matching"
    )

    st.info(
        "Use the sidebar to search/select a yarn, then select the required "
        "STORES/WIP lots and Pending POs, then "
        "save the matching."
    )

    # ========================================================
    # SIDEBAR YARN LIST
    # ========================================================

    # The sidebar is now the only yarn search/navigation control.
    descriptions_with_index = list(
        enumerate(all_descriptions)
    )

    # ========================================================
    # MATCHING UI — SIDEBAR YARN NAVIGATION
    # ========================================================

    if "selected_yarn_description" not in st.session_state:
        st.session_state.selected_yarn_description = (
            all_descriptions[0]
            if all_descriptions
            else None
        )

    # Keep the selected yarn valid after searching.
    visible_descriptions = [
        description
        for _, description in descriptions_with_index
    ]

    if (
        st.session_state.selected_yarn_description
        not in visible_descriptions
    ):
        st.session_state.selected_yarn_description = (
            visible_descriptions[0]
            if visible_descriptions
            else None
        )

    # --------------------------------------------------------
    # SIDEBAR
    # --------------------------------------------------------

    with st.sidebar:
        st.header("🧶 Yarn Matching")

        st.caption(
            f"{len(all_descriptions)} yarn description(s)"
        )

        sidebar_search = st.text_input(
            "🔎 Search Yarn",
            key="sidebar_yarn_search",
            placeholder="Search description..."
        )

        sidebar_search_text = (
            sidebar_search.strip().lower()
        )

        if sidebar_search_text:
            sidebar_descriptions = [
                description
                for description in all_descriptions
                if sidebar_search_text in str(description).lower()
            ]
        else:
            sidebar_descriptions = all_descriptions

        st.caption(
            f"Showing {len(sidebar_descriptions)} yarn(s)"
        )

        # Use buttons instead of rendering every yarn's matching
        # controls. Only the selected yarn's controls appear below.
        for description in sidebar_descriptions:

            saved = st.session_state.match_selections.get(
                description,
                {}
            )

            is_saved = bool(
                saved.get("_saved", False)
            )

            prefix = "✅ " if is_saved else "⚪ "

            if st.button(
                prefix + str(description),
                key=f"yarn_select_{all_descriptions.index(description)}",
                use_container_width=True,
                type=(
                    "primary"
                    if description
                    == st.session_state.selected_yarn_description
                    else "secondary"
                )
            ):
                st.session_state.selected_yarn_description = description
                st.rerun()

    # --------------------------------------------------------
    # CURRENT YARN MATCHING
    # --------------------------------------------------------

    selected_description = (
        st.session_state.selected_yarn_description
    )

    if selected_description:

        automatic_key = normalize_yarn_description(
            selected_description
        )

        stores_options = bin_lookup.get(
            (automatic_key, "STORES"),
            []
        )

        wip_options = bin_lookup.get(
            (automatic_key, "WIP"),
            []
        )

        pending_options = pending_lookup.get(
            automatic_key,
            []
        )

        saved = st.session_state.match_selections.get(
            selected_description,
            {}
        )

        saved_stores = saved.get(
            "stores_lots",
            []
        )

        saved_wip = saved.get(
            "wip_lots",
            []
        )

        saved_pending = saved.get(
            "pending_lots",
            []
        )

        st.markdown(
            f"### 🧶 {selected_description}"
        )

        st.caption(
            "Automatic matching key: "
            + automatic_key
        )

        # ----------------------------------------------------
        # CURRENT YARN MATCHING FORM
        # ----------------------------------------------------
        # All checkboxes are inside this form. Therefore clicking a
        # checkbox does NOT rerun the full Streamlit program.
        # The app reruns only when Save Matching is clicked.
        with st.form(
            f"matching_form_{all_descriptions.index(selected_description)}"
        ):

            # ------------------------------------------------
            # STORES
            # ------------------------------------------------

            st.markdown("#### 📦 STORES")

            current_stores = []

            if not stores_options:
                st.caption("No matching STORES lots.")
            else:
                for option_number, option in enumerate(stores_options):

                    lot = option["lot"]

                    selected = st.checkbox(
                        option["description"]
                        + " | Lot: "
                        + lot
                        + " | Qty: "
                        + f"{option['quantity']:,.2f}",
                        value=(
                            lot in saved_stores
                            or (
                                len(stores_options) == 1
                                and not saved.get("_saved", False)
                            )
                        ),
                        key=(
                            f"current_stores_"
                            f"{all_descriptions.index(selected_description)}_"
                            f"{option_number}"
                        )
                    )

                    if selected:
                        current_stores.append(lot)

                if current_stores:
                    stores_total = sum(
                        option["quantity"]
                        for option in stores_options
                        if option["lot"] in current_stores
                    )

                    st.caption(
                        f"Selected {len(current_stores)} lot(s) | "
                        f"Total: {stores_total:,.2f}"
                    )

            # ------------------------------------------------
            # WIP
            # ------------------------------------------------

            st.markdown("#### 🏭 WIP")

            current_wip = []

            if not wip_options:
                st.caption("No matching WIP lots.")
            else:
                for option_number, option in enumerate(wip_options):

                    lot = option["lot"]

                    selected = st.checkbox(
                        option["description"]
                        + " | Lot: "
                        + lot
                        + " | Qty: "
                        + f"{option['quantity']:,.2f}",
                        value=(
                            lot in saved_wip
                            or (
                                len(wip_options) == 1
                                and not saved.get("_saved", False)
                            )
                        ),
                        key=(
                            f"current_wip_"
                            f"{all_descriptions.index(selected_description)}_"
                            f"{option_number}"
                        )
                    )

                    if selected:
                        current_wip.append(lot)

                if current_wip:
                    wip_total = sum(
                        option["quantity"]
                        for option in wip_options
                        if option["lot"] in current_wip
                    )

                    st.caption(
                        f"Selected {len(current_wip)} lot(s) | "
                        f"Total: {wip_total:,.2f}"
                    )

            # ------------------------------------------------
            # PENDING PO
            # ------------------------------------------------

            st.markdown("#### 📋 Pending PO")

            current_pending = []

            if not pending_options:
                st.caption("No matching Pending PO.")
            else:
                for option_number, option in enumerate(pending_options):

                    lot = option["lot"]

                    selected = st.checkbox(
                        option["description"]
                        + " | Document/Lot: "
                        + lot
                        + " | Qty: "
                        + f"{option['quantity']:,.2f}",
                        value=(
                            lot in saved_pending
                            or (
                                len(pending_options) == 1
                                and not saved.get("_saved", False)
                            )
                        ),
                        key=(
                            f"current_pending_"
                            f"{all_descriptions.index(selected_description)}_"
                            f"{option_number}"
                        )
                    )

                    if selected:
                        current_pending.append(lot)

                if current_pending:
                    pending_total = sum(
                        option["quantity"]
                        for option in pending_options
                        if option["lot"] in current_pending
                    )

                    st.caption(
                        f"Selected {len(current_pending)} Pending PO(s) | "
                        f"Total: {pending_total:,.2f}"
                    )

            st.divider()

            save_current = st.form_submit_button(
                "💾 Save Matching",
                type="primary",
                use_container_width=True
            )

        # ------------------------------------------------
        # SAVE CURRENT YARN
        # ------------------------------------------------

        if save_current:

            st.session_state.match_selections[
                selected_description
            ] = {
                "stores_lots": current_stores,
                "wip_lots": current_wip,
                "pending_lots": current_pending,
                "_saved": True
            }

            st.success(
                f"✅ Matching saved for: "
                f"{selected_description}"
            )

            st.rerun()

    else:
        st.info(
            "Select a yarn description from the sidebar "
            "to start matching."
        )

    # ========================================================
    # MATCHING SUMMARY
    # ========================================================

    st.markdown(
        "### Matching Summary"
    )

    summary_rows = []

    for description in all_descriptions:

        selection = (
            st.session_state
            .match_selections
            .get(
                description,
                {}
            )
        )

        stores_lots = selection.get(
            "stores_lots",
            []
        )

        wip_lots = selection.get(
            "wip_lots",
            []
        )

        pending_lots = selection.get(
            "pending_lots",
            []
        )

        # STORES

        stores_options = bin_lookup.get(
            (
                normalize_yarn_description(
                    description
                ),
                "STORES"
            ),
            []
        )

        stores_values = []

        for option in stores_options:

            if option["lot"] in stores_lots:

                stores_values.append(
                    f"{option['lot']} "
                    f"({option['quantity']:,.2f})"
                )

        stores_text = (
            " | ".join(stores_values)
            if stores_values
            else "No match"
        )

        # WIP

        wip_options = bin_lookup.get(
            (
                normalize_yarn_description(
                    description
                ),
                "WIP"
            ),
            []
        )

        wip_values = []

        for option in wip_options:

            if option["lot"] in wip_lots:

                wip_values.append(
                    f"{option['lot']} "
                    f"({option['quantity']:,.2f})"
                )

        wip_text = (
            " | ".join(wip_values)
            if wip_values
            else "No match"
        )

        # PENDING

        pending_options = pending_lookup.get(
            normalize_yarn_description(
                description
            ),
            []
        )

        pending_values = []

        for option in pending_options:

            if option["lot"] in pending_lots:

                pending_values.append(
                    f"{option['lot']} "
                    f"({option['quantity']:,.2f})"
                )

        pending_text = (
            " | ".join(pending_values)
            if pending_values
            else "No match"
        )

        summary_rows.append(
            {
                "Firm Plan":
                    description,

                "STORES":
                    stores_text,

                "WIP":
                    wip_text,

                "Pending PO":
                    pending_text
            }
        )

    st.dataframe(
        pd.DataFrame(
            summary_rows
        ),
        use_container_width=True,
        hide_index=True
    )

    # ========================================================
    # CONFIRM MATCHING
    # ========================================================

    if st.button(
        "✅ Confirm Yarn Matching",
        type="primary",
        use_container_width=True
    ):

        (
            matched_firm,
            matched_bin,
            matched_pending
        ) = apply_matching(
            firm_plan,
            bin_data,
            pending_po,
            st.session_state.match_selections
        )

        st.session_state.firm_plan = (
            matched_firm
        )

        st.session_state.bin_data = (
            matched_bin
        )

        st.session_state.pending_po = (
            matched_pending
        )

        st.session_state.matching_confirmed = True

        st.session_state.report_df = None

        st.session_state.pending_firm_plan = None

        st.session_state.calculated = False

        st.success(
            "Yarn matching confirmed."
        )

        st.rerun()

    # ========================================================
    # WAIT
    # ========================================================

    if not st.session_state.matching_confirmed:

        st.warning(
            "Your checkbox selections are saved automatically. "
            "Click 'Confirm Yarn Matching' when finished."
        )

        st.stop()

    # ========================================================
    # STEP 2 — REQUIREMENT PREVIEW
    # ========================================================

    st.subheader(
        "2️⃣ Requirement Preview"
    )

    firm_plan = (
        st.session_state.firm_plan
    )

    bin_data = (
        st.session_state.bin_data
    )

    pending_po = (
        st.session_state.pending_po
    )

    if st.session_state.report_df is None:

        report_df = create_report(
            firm_plan,
            bin_data,
            pending_po
        )

        (
            report_df,
            pending_firm_plan
        ) = allocate_production(
            firm_plan,
            bin_data,
            pending_po,
            report_df
        )

        st.session_state.report_df = (
            report_df
        )

        st.session_state.pending_firm_plan = (
            pending_firm_plan
        )

    report_df = (
        st.session_state.report_df.copy()
    )

    preview_columns = [
        "Description",
        "Firm Plan",
        "Wastage 20%",
        "Requirement",
        "Safety Stock",
        "Stores",
        "WIP",
        "Purchase Qty",
        "Manual Purchase Qty"
    ]

    with st.form("requirement_form"):
        edited_preview = st.data_editor(
            report_df[
                preview_columns
            ].copy(),
            use_container_width=True,
            hide_index=True,

            disabled=[
                "Description",
                "Firm Plan",
                "Wastage 20%",
                "Requirement"
            ],

            column_config={

                "Firm Plan":
                    st.column_config.NumberColumn(
                        format="%.2f"
                    ),

                "Wastage 20%":
                    st.column_config.NumberColumn(
                        format="%.2f"
                    ),

                "Requirement":
                    st.column_config.NumberColumn(
                        format="%.2f"
                    ),

                "Safety Stock":
                    st.column_config.NumberColumn(
                        format="%.2f"
                    ),

                "Stores":
                    st.column_config.NumberColumn(
                        format="%.2f"
                    ),

                "WIP":
                    st.column_config.NumberColumn(
                        format="%.2f"
                    ),

                "Purchase Qty":
                    st.column_config.NumberColumn(
                        format="%.2f"
                    ),

                "Manual Purchase Qty":
                    st.column_config.NumberColumn(
                        format="%.2f"
                    )
            },

            key="requirement_editor"
        )

        calculate_clicked = st.form_submit_button(
            "🔄 Calculate",
            type="primary",
            use_container_width=True
        )

    # ========================================================
    # CALCULATE
    # ========================================================

    if calculate_clicked:

        working = (
            st.session_state
            .report_df
            .copy()
        )

        working["Stores"] = pd.to_numeric(
            edited_preview["Stores"],
            errors="coerce"
        ).fillna(0)

        working["WIP"] = pd.to_numeric(
            edited_preview["WIP"],
            errors="coerce"
        ).fillna(0)

        working["Safety Stock"] = pd.to_numeric(
            edited_preview["Safety Stock"],
            errors="coerce"
        ).fillna(0)

        working[
            "Manual Purchase Qty"
        ] = pd.to_numeric(
            edited_preview[
                "Manual Purchase Qty"
            ],
            errors="coerce"
        ).fillna(0)

        # Recalculate Total Stock before production allocation so
        # edited STORES / WIP / Safety Stock values are used.
        working["Total Stock"] = (
            working["Stores"]
            + working["WIP"]
        )

        # Re-run production allocation using the ORIGINAL Firm Plan
        # and the edited report as the stock source.
        (
            working,
            pending_firm_plan
        ) = allocate_production(
            st.session_state.firm_plan,
            st.session_state.bin_data,
            st.session_state.pending_po,
            working
        )

        final_df = calculate_final_report(
            working,
            pending_firm_plan
        )

        st.session_state.report_df = (
            final_df
        )

        st.session_state.pending_firm_plan = (
            pending_firm_plan
        )

        st.session_state.calculated = True

        st.rerun()

    # ========================================================
    # STEP 3 — FINAL REPORT
    # ========================================================

    if st.session_state.calculated:

        final_df = (
            st.session_state
            .report_df
            .copy()
        )

        st.subheader(
            "3️⃣ Final Requirement Report"
        )

        final_columns = [
            "Description",
            "Firm Plan",
            "Released",
            "Wastage 20%",
            "Requirement",
            "Safety Stock",
            "Gross Requirement",
            "Stores",
            "WIP",
            "Production",
            "Total Stock",
            "Diff",
            "Purchase Qty",
            "Requirement PO",
            "New Purchase Requirement",
            "Remarks"
        ]

        st.dataframe(
            final_df[
                final_columns
            ],
            use_container_width=True,
            hide_index=True
        )

        # ====================================================
        # STEP 4
        # ====================================================

        st.subheader(
            "4️⃣ Pending PO to be Woven"
        )

        pending_df = (
            st.session_state
            .pending_firm_plan
            .copy()
        )

        pending_display = (
            pending_df
            .drop(
                columns=[
                    "_Original Row",
                    "_Automatic Key",
                    "_Matching Key",
                    "_Base PO"
                ],
                errors="ignore"
            )
        )

        if len(pending_display) > 0:

            st.dataframe(
                pending_display,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.success(
                "No pending Firm Plan POs."
            )

        # ====================================================
        # DOWNLOAD
        # ====================================================

        output_file = create_output_workbook(
            file_bytes,
            final_df,
            pending_df
        )

        st.download_button(
            "📥 Download Updated Excel Workbook",
            data=output_file,
            file_name=(
                "Yarn_Requirement_Analysis.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True
        )
