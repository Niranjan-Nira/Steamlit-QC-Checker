import io
import os
from datetime import date, timedelta

import pandas as pd
import streamlit as st

# -----------------------------
# App Config
# -----------------------------
st.set_page_config(page_title="Excel QA Checker", page_icon="📊", layout="wide")

# Columns that need leading zeros preserved in Excel export
COLUMNS_WITH_LEADING_ZEROS = ["Business Id", "Local Code"]

# -----------------------------
# Helpers
# -----------------------------
def sanitize_sheet_name(name: str) -> str:
    r"""Excel sheet names: max 31 chars, no \ / * [ ] : ?"""
    for ch in ['\\', '/', '*', '[', ']', ':', '?']:
        name = name.replace(ch, "_")
    return name[:31] if name else "Sheet"

def to_datetime_col(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")

def ensure_leading_zeros(df: pd.DataFrame, columns=COLUMNS_WITH_LEADING_ZEROS) -> pd.DataFrame:
    for col in columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.zfill(7)
    return df

def export_sections_to_excel(sections):
    """
    sections: list of dicts with keys: title, df (DataFrame)
    Returns: bytes of an .xlsx file
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sec in sections:
            title = sanitize_sheet_name(sec["title"])
            df = sec["df"].copy()

            # Preserve leading zeros for requested columns if present
            for col in COLUMNS_WITH_LEADING_ZEROS:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.zfill(7)

            df.to_excel(writer, sheet_name=title, index=False)
    output.seek(0)
    return output

# -----------------------------
# Core Processing
# -----------------------------
def process_excel(file_like) -> list[dict]:
    """
    Returns a list of sections: [{title, df, info, count}]
    """
    df = pd.read_excel(file_like, engine=None)  # let pandas pick the right engine
    df = ensure_leading_zeros(df)

    sections = []

    # Parse Verification Date
    if "Verification Date" in df.columns:
        df["Verification Date"] = to_datetime_col(df["Verification Date"])

        # Unparsed dates
        if df["Verification Date"].isna().any():
            unparsed_df = df[df["Verification Date"].isna()].copy()
            if not unparsed_df.empty:
                sections.append({
                    "title": "Unparsed Verification Dates",
                    "df": unparsed_df,
                    "info": "Some rows have invalid 'Verification Date' values that could not be parsed.",
                    "count": len(unparsed_df)
                })

        # Should be Today or Yesterday
        today = pd.to_datetime(date.today())
        yesterday = pd.to_datetime(date.today() - timedelta(days=1))

        # Normalize datetimes to date (removes time-of-day differences)
        norm = df["Verification Date"].dt.normalize()
        mask = norm.isin([today.normalize(), yesterday.normalize()])
        verificationdate_df = df[~mask.fillna(False)].copy()
        sections.append({
            "title": "Verification Date ",
            "df": verificationdate_df,
            "info": "Verification Date should be either today or yesterday.",
            "count": len(verificationdate_df)
        })
    else:
        sections.append({
            "title": "Verification Date (Column Missing)",
            "df": pd.DataFrame(),
            "info": "'Verification Date' column is missing.",
            "count": 0
        })

    # Bar/Nightclub, Dining, Caterers without food/BWL
    def missing_food_flags(frame):
        return ((frame["Food Type"].isnull()) |
                (frame["Beer"].isnull()) |
                (frame["Wine"].isnull()) |
                (frame["Liquor"].isnull()))

    required_cols = ["Local Trade Channel", "Food Type", "Beer", "Wine", "Liquor"]
    if all(col in df.columns for col in required_cols):
        bar = df[(df["Local Trade Channel"] == "Bar/Nightclub") & missing_food_flags(df)]
        cater = df[(df["Local Trade Channel"] == "Caterers") & missing_food_flags(df)]
        dining = df[(df["Local Trade Channel"] == "Dining") & missing_food_flags(df)]
        combined = pd.concat([bar, cater, dining], ignore_index=True)
        sections.append({
            "title": "Bar/Nightclub, Dining & Caterers Without Food",
            "df": combined,
            "info": "Bar/Nightclub, Dining & Caterers should always have the food type and B/W/L flags.",
            "count": len(combined)
        })

    # Wholesale Clubs / Mass Merchandise / Category killer / Fulfillment - one of IRT/MG missing
    need_cols = ["Local Trade Channel", "IRT Local Code", "MG Local Code"]
    if all(col in df.columns for col in need_cols):
        def one_missing(frame, channel):
            part1 = df[(df["Local Trade Channel"] == channel) &
                       (df["IRT Local Code"].notnull()) &
                       (df["MG Local Code"].isnull())]
            part2 = df[(df["Local Trade Channel"] == channel) &
                       (df["IRT Local Code"].isnull()) &
                       (df["MG Local Code"].notnull())]
            return pd.concat([part1, part2], ignore_index=True)

        combined_wc = one_missing(df, "Wholesale Clubs")
        combined_mm = one_missing(df, "Mass Merchandise Stores")
        combined_ck = one_missing(df, "Category killer")
        combined_f  = one_missing(df, "Fulfillment")
        final_combined = pd.concat([combined_wc, combined_mm, combined_ck, combined_f], ignore_index=True)

        sections.append({
            "title": "Without IRT And MG Wholesale Clubs, Category Killers, Mass Merchandise, Fulfillment",
            "df": final_combined,
            "info": "IRT or MG are missing for the below TDCs.",
            "count": len(final_combined)
        })

    # Pharmacy flag
    if all(col in df.columns for col in ["Local Trade Channel", "Pharmacy"]):
        pharmacy_df = df[(df["Local Trade Channel"] == "Drug Stores and Pharmacies") &
                         ((df["Pharmacy"] == "N") | (df["Pharmacy"].isnull()))]
        sections.append({
            "title": "Pharmacy Flag ",
            "df": pharmacy_df,
            "info": "Trade channel is Drug Stores and Pharmacies but Pharmacy flag is 'N' or blank.",
            "count": len(pharmacy_df)
        })

    # Banner Rule 1: capitalization (istitle)
    if "Name" in df.columns:
        banner1_df = df[df["Name"].fillna("").astype(str).str.istitle() == False]
        sections.append({
            "title": "Banner Rule 1 Exceptions",
            "df": banner1_df,
            "info": "Banner name should be in title case. Ex: 'new york' → 'New York'.",
            "count": len(banner1_df)
        })

        # Banner Rule 2/3 keyword filters
        banner2_df = df[df["Name"].fillna("").str.contains(
            r" Advertising| Accounting| Company| And| Cos| distribution| distributor| ent| inc| LLC| region| Warehouse",
            case=False, na=False)]
        banner3_df = df[df["Name"].fillna("").str.contains(" Co", case=False, na=False, regex=False)]
        bannercombo_df = pd.concat([banner2_df, banner3_df], ignore_index=True)
        sections.append({
            "title": "Banner Rule 2 Exceptions",
            "df": bannercombo_df,
            "info": "Banner has Advertising/Accounting/Company/And/Cos/Co/dist/distribution/distributor/ent/inc/LLC/region/Warehouse which are not allowed.",
            "count": len(bannercombo_df)
        })

    # Address Quality
    addr_cols = ["Address Quality", "Address"]
    if all(col in df.columns for col in addr_cols):
        address_df = df[(df["Address Quality"] == "Non Standardized") &
                        df["Address"].fillna("").str.contains("Highway|Street|Parkway|Route", case=False)]
        sections.append({
            "title": "Address ",
            "df": address_df,
            "info": "Address contains Highway/Parkway/Route and is Non Standardized.",
            "count": len(address_df)
        })

    # Wrong MG Exception
    if all(col in df.columns for col in ["Name", "MG Name"]):
        wrong_mg_df = df[(df["MG Name"].notna()) & (df["Name"] != df["MG Name"])]
        sections.append({
            "title": "Wrong MG Exception",
            "df": wrong_mg_df,
            "info": "Banner name does not match MG Name. Exception MG Name will not always match banner.",
            "count": len(wrong_mg_df)
        })

    # Gas Flag
    need_cols = ["Name", "Gas"]
    if all(col in df.columns for col in need_cols):
        gas_df = df[df["Name"].fillna("").str.contains("Gas|Fuel", case=False, na=False) &
                    ((df["Gas"] == "N") | (df["Gas"].isnull()))]
        sections.append({
            "title": "Gas Flag",
            "df": gas_df,
            "info": "Banner contains 'Gas'/'Fuel' but Gas flag is null/blank.",
            "count": len(gas_df)
        })

    # BWL blanks in Pet, Cannabis, Fulfillment
    need_cols = ["Local Trade Channel", "Beer", "Wine", "Liquor"]
    if all(col in df.columns for col in need_cols):
        bwl_blank = df[(df["Local Trade Channel"].isin(["Cannabis", "Pet", "Fulfillment"])) &
                       ((df["Beer"].isnull()) | (df["Wine"].isnull()) | (df["Liquor"].isnull()))]
        sections.append({
            "title": "BWL Blanks in Pet, Cannabis & Fulfillment",
            "df": bwl_blank,
            "info": "Pet, Cannabis & Fulfillment should not have blanks in B/W/L.",
            "count": len(bwl_blank)
        })

    # Exception codes 93Z / 98Z / 94Z
    if "Exception Code" in df.columns:
        if "Status" in df.columns and "Verification Source" in df.columns:
            exception93_df = df[(df["Status"].isin(["Open, Operating", "Closed", "Future Opening", "Inactive/Not Verified"])) &
                                (df["Exception Code"] == "777793Z")]
            exception93rule2_df = df[(df["Exception Code"] == "777793Z") &
                                     (df["Verification Source"] != "Attempted Contact Failed")]
            combined_93 = pd.concat([exception93_df, exception93rule2_df], ignore_index=True)
            sections.append({
                "title": "Exception code 93Z ",
                "df": combined_93,
                "info": "93Z is for 'Unverifiable'. Status should not be Open/Closed/Future Opening and VSS should be 'Attempted Contact Failed'.",
                "count": len(combined_93)
            })

        exception98_df = df[df["Exception Code"].eq("777798Z") &
                            (df["Status"].isin(["Unverifiable", "Future Opening"]) if "Status" in df.columns else True)]
        sections.append({
            "title": "Exception code 98Z ",
            "df": exception98_df,
            "info": "Store status 'Unverifiable' or 'Future Opening' with 98Z.",
            "count": len(exception98_df)
        })

        if "MG Name" in df.columns:
            exception94_df = df[df["MG Name"].fillna("").str.contains("/EM") & (df["Exception Code"] != "777794Z")]
            sections.append({
                "title": "Exception code 94Z ",
                "df": exception94_df,
                "info": "94Z is for business having both syndicated and non-syndicated.",
                "count": len(exception94_df)
            })

    # Syndicate FO
    if all(col in df.columns for col in ["Local Trade Channel", "Status"]):
        syndicate_fo = df[(df["Local Trade Channel"].isin(["Unknown Retailers", "Unknown On-Premise"])) &
                          (df["Status"] == "Future Opening")]
        sections.append({
            "title": "FO for Non-Syndicate ",
            "df": syndicate_fo,
            "info": "Unknown Retailers/Unknown On-Premise with Future Opening status.",
            "count": len(syndicate_fo)
        })

    # Attempted Contact Failed for Open, Operating
    if all(col in df.columns for col in ["Status", "Verification Source"]):
        attempt_cnt = df[(df["Status"] == "Open, Operating") &
                         (df["Verification Source"] == "Attempted Contact Failed")]
        sections.append({
            "title": "Attempt Contact Failed for OP ",
            "df": attempt_cnt,
            "info": "Open Operating should not have 'Attempted Contact Failed'.",
            "count": len(attempt_cnt)
        })

    # Unverifiable with wrong trade channel
    need_cols = ["Status", "Local Trade Channel", "Local Sub Channel", "Sec Trade Channel", "Sec Sub Channel"]
    if all(col in df.columns for col in need_cols):
        unverifiable_df = df[(df["Status"] == "Unverifiable") & (
            (df["Local Trade Channel"] != "Unknown Retailers") |
            (df["Local Sub Channel"] != "Retail Other") |
            (df["Sec Trade Channel"] != "Miscellaneous Retail") |
            (df["Sec Sub Channel"] != "Other Retail"))]
        sections.append({
            "title": "Unverifiable With Wrong Trade Channel ",
            "df": unverifiable_df,
            "info": "Expected: Unknown Retailers / Retail Other / Miscellaneous Retail / Other Retail.",
            "count": len(unverifiable_df)
        })

    # Null phone for Open, Operating (exclude Transportation & Fulfillment)
    need_cols = ["Status", "Phone", "Local Trade Channel"]
    if all(col in df.columns for col in need_cols):
        nullph_df = df[(df["Status"] == "Open, Operating") &
                       (df["Phone"].isnull()) &
                       (~df["Local Trade Channel"].isin(["Transportation", "Fulfillment"]))]
        sections.append({
            "title": "Null Phone For OP ",
            "df": nullph_df,
            "info": "Null phone for Open Operating (excluding Transportation/Fulfillment).",
            "count": len(nullph_df)
        })

    # State BWL Grid
    need_cols = ["State/Province", "Beer", "Wine", "Liquor", "Local Trade Channel"]
    if all(col in df.columns for col in need_cols):
        stategrid = df[(df["State/Province"].isin(["New York", "South Carolina"])) &
                       (df["Beer"] == "Y") & (df["Wine"] == "Y") & (df["Liquor"] == "Y") &
                       (df["Local Trade Channel"] == "Liquor, Wine and Beer Stores")]
        sections.append({
            "title": "State BWL Grid for New York and South Carolina ",
            "df": stategrid,
            "info": "Liquor shops in NY/SC cannot sell B/W/L in a single TD.",
            "count": len(stategrid)
        })

    # Unusual VSS (exclude common sources)
    if "Verification Source" in df.columns:
        exclude_sources = [
            "Web Sites, Other",
            "Telephone, Direct",
            "Telephone, Indirect",
            "Attempted Contact Failed",
            "Web Lookup",
        ]
        unusual_vss = df[~df["Verification Source"].isin(exclude_sources)]
        sections.append({
            "title": "Unusual VSS",
            "df": unusual_vss,
            "info": "Verification sources not commonly used.",
            "count": len(unusual_vss)
        })

    # Pet supplier rules
    need_cols = ["Local Trade Channel", "IRT Local Code", "Grocery Supplier Number"]
    if all(col in df.columns for col in need_cols):
        petsupplier_df = df[(df["Local Trade Channel"] == "Pet") &
                            (df["IRT Local Code"].notnull()) &
                            (df["Grocery Supplier Number"].isnull())]
        sections.append({
            "title": "Pet Supplier",
            "df": petsupplier_df,
            "info": "Pet with IRT but missing Grocery Supplier. Pet can have only Grocery Supplier.",
            "count": len(petsupplier_df)
        })

    # Pet chain
    need_cols = ["Name", "IRT Local Code", "MG Local Code", "Grocery Supplier Number"]
    if all(col in df.columns for col in need_cols):
        petchain_df = df[df["Name"].fillna("").str.contains("Petco|PetSmart|Unleashed By Petco", case=False, na=False) &
                         ((df["IRT Local Code"].isnull()) |
                          (df["MG Local Code"].isnull()) |
                          (df["Grocery Supplier Number"].isnull()))]
        sections.append({
            "title": "Pet Chain ",
            "df": petchain_df,
            "info": "Petco/PetSmart/Unleashed By Petco missing IRT or MG or Supplier.",
            "count": len(petchain_df)
        })

    # Closed stores
    if "Status" in df.columns:
        closed_df = df[df["Status"] == "Closed"]
        sections.append({
            "title": "Closed Store",
            "df": closed_df,
            "info": "All stores marked as Closed status.",
            "count": len(closed_df)
        })

    # Chain with supplier but without IRT (Grocery/Convenience)
    need_cols = ["Local Trade Channel", "Grocery Supplier Number", "IRT Local Code"]
    if all(col in df.columns for col in need_cols):
        chain_irt = df[(df["Local Trade Channel"].isin(["Grocery Stores", "Convenience Stores"])) &
                       (df["Grocery Supplier Number"].notnull()) &
                       (df["IRT Local Code"].isnull())]
        sections.append({
            "title": "Grocery & Convenience Stores Without IRT",
            "df": chain_irt,
            "info": "Grocery/Convenience have Supplier without IRT.",
            "count": len(chain_irt)
        })

    # Grocery/Mass/Convenience with IRT or MG but missing various suppliers (and exclude gas stations for Convenience)
    need_cols = ["Local Trade Channel", "IRT Local Code", "MG Local Code",
                 "Grocery Supplier Number", "Confection Supplier Number", "GM Supplier Number",
                 "Frozen Supplier Number", "HBC Supplier Number", "Local Sub Channel"]
    if all(col in df.columns for col in need_cols):
        grocery_df = df[(df["Local Trade Channel"] == "Grocery Stores") &
                        ((df["IRT Local Code"].notnull()) | (df["MG Local Code"].notnull())) &
                        ((df["Grocery Supplier Number"].isnull()) |
                         (df["Confection Supplier Number"].isnull()) |
                         (df["GM Supplier Number"].isnull()) |
                         (df["Frozen Supplier Number"].isnull()) |
                         (df["HBC Supplier Number"].isnull()))]

        mass_df = df[(df["Local Trade Channel"] == "Mass Merchandise Stores") &
                     ((df["IRT Local Code"].notnull()) | (df["MG Local Code"].notnull())) &
                     ((df["Grocery Supplier Number"].isnull()) |
                      (df["Confection Supplier Number"].isnull()) |
                      (df["GM Supplier Number"].isnull()) |
                      (df["HBC Supplier Number"].isnull()))]

        conven_df = df[(df["Local Trade Channel"] == "Convenience Stores") &
                       (df["Local Sub Channel"] != "Gasoline Stations") &
                       ((df["IRT Local Code"].notnull()) | (df["MG Local Code"].notnull())) &
                       ((df["Grocery Supplier Number"].isnull()) |
                        (df["Confection Supplier Number"].isnull()))]

        combined_stores = pd.concat([grocery_df, mass_df, conven_df], ignore_index=True)
        sections.append({
            "title": "Grocery, Mass Merchandise & Convenience Stores Without Supplier",
            "df": combined_stores,
            "info": "These stores have IRT/MG but are missing one or more required suppliers (gas stations excluded for Convenience).",
            "count": len(combined_stores)
        })

    # IRT present but MG missing for Fulfillment/Grocery/Convenience
    need_cols = ["Local Trade Channel", "IRT Local Code", "MG Local Code"]
    if all(col in df.columns for col in need_cols):
        irtnomg = df[df["Local Trade Channel"].isin(["Fulfillment", "Grocery Stores", "Convenience Stores"]) &
                     (df["IRT Local Code"].notnull()) &
                     (df["MG Local Code"].isnull())]
        sections.append({
            "title": "Grocery, Convenience & Fulfillment Stores having IRT Without MG",
            "df": irtnomg,
            "info": "Has IRT but missing MG.",
            "count": len(irtnomg)
        })

    # Inactive / Special Event / Special Projects rule
    need_cols = ["Local Trade Channel", "Local Sub Channel", "Verification Source", "Status"]
    if all(col in df.columns for col in need_cols):
        inactive_df = df[(df["Local Trade Channel"] == "Unknown On-Premise") &
                         (df["Local Sub Channel"] == "Special Event") &
                         (df["Verification Source"] != "Special Projects") &
                         (df["Status"] != "Inactive/Not Verified")]
        sections.append({
            "title": "Inactive status",
            "df": inactive_df,
            "info": "Unknown On-Premise + Special Event must have VSS='Special Projects' and Status='Inactive/Not Verified'.",
            "count": len(inactive_df)
        })

    # ChainIRT_html already covered above as "Grocery & Convenience Stores Without IRT"

    # Nice-to-have: sort sections by count (desc) to surface biggest issues first
    sections_sorted = sorted(sections, key=lambda s: s["count"], reverse=True)
    return sections_sorted

# -----------------------------
# UI
# -----------------------------
st.title("📊 Data Quality Checker — Streamlit")
st.write("Upload your Excel file to run all quality checks. View results by section and export everything to a multi-sheet Excel file.")

uploaded_file = st.file_uploader("Upload an Excel file (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("Processing…"):
        sections = process_excel(uploaded_file)

    # Summary
    total_rows = sum(s["count"] for s in sections)
    st.success(f"Done! Found **{total_rows}** total flagged rows across **{len(sections)}** sections.")

    # Export button
    xlsx_bytes = export_sections_to_excel(sections)
    st.download_button(
        label="⬇️ Download all results as Excel",
        data=xlsx_bytes,
        file_name="processed_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # Display sections
    st.markdown("### Results by Section")
    # Show the most populated sections first
    for sec in sections:
        title = sec["title"].strip()
        count = sec["count"]
        info = sec.get("info", "")
        df = sec["df"]

        with st.expander(f"{title} — {count} row(s)"):
            if info:
                st.caption(info)
            if df is not None and not df.empty:
                st.dataframe(df, use_container_width=True, hide_index=True)
            else:
                st.info("No rows in this section.")

else:
    st.info("👆 Start by uploading an Excel file (.xlsx/.xls).")

# Footer
st.caption("Converted from Flask to Streamlit • Preserves leading zeros on export • Click a section to inspect rows.")
