import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Amazon Transaction Report", page_icon="📦", layout="wide")

st.title("📦 Amazon Transaction Report Generator")
st.markdown("Upload your files below to generate **Orders Pivot**, **Refunds Pivot**, **Returns Analysis**, and **Brand-wise Reverse Logistics** reports.")

# ─── File Uploaders ───────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📁 Upload Files")
    transaction_file = st.file_uploader("1. Unified Transaction Report (CSV)", type=["csv"])
    returns_file     = st.file_uploader("2. Returns Report (CSV)", type=["csv"])
    pm_file          = st.file_uploader("3. PM Master File (XLSX)", type=["xlsx"])
    st.markdown("---")
    st.info("All three files are required to generate the full report.")

def to_excel_bytes(dfs: dict) -> bytes:
    """Write multiple DataFrames into a single Excel file (one sheet each)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buf.getvalue()

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")

# ─── Processing ───────────────────────────────────────────────────────────────
if transaction_file and returns_file and pm_file:
    with st.spinner("Processing files…"):

        # 1. Load Transaction CSV
        df = pd.read_csv(
            transaction_file,
            skiprows=11,
            thousands=",",
            low_memory=False,
        )

        # 2. Orders
        df_orders = df[df["type"] == "Order"].copy()
        df_orders = df_orders[df_orders["product sales"] != 0]
        df_orders = df_orders.sort_values("order id")
        df_orders["order id"] = df_orders["order id"].astype(str)
        df_orders["Sku"]      = df_orders["Sku"].astype(str)
        df_orders["Con"]      = df_orders["order id"] + df_orders["Sku"]

        pivot_orders = pd.pivot_table(
            df_orders,
            index="Con",
            values=["quantity", "total"],
            aggfunc="sum",
            margins=True,
            margins_name="Grand Total",
        ).reset_index()
        pivot_orders.columns.name = None
        pivot_orders = pivot_orders.sort_index()
        pivot_orders["total"] = pivot_orders["total"].round(2)

        # 3. Refunds
        df_refunds = df[df["type"] == "Refund"].copy()
        df_refunds = df_refunds[df_refunds["product sales"] != 0]
        df_refunds = df_refunds.sort_values("order id")
        df_refunds["order id"] = df_refunds["order id"].astype(str)
        df_refunds["Sku"]      = df_refunds["Sku"].astype(str)
        df_refunds["Con"]      = df_refunds["order id"] + df_refunds["Sku"]

        pivot_refunds = pd.pivot_table(
            df_refunds,
            index="Con",
            values=["quantity", "total"],
            aggfunc="sum",
            margins=True,
            margins_name="Grand Total",
        ).reset_index()
        pivot_refunds.columns.name = None
        pivot_refunds = pivot_refunds.sort_index()
        pivot_refunds["total"] = pivot_refunds["total"].round(2)

        # 4. Returns
        df_returns = pd.read_csv(returns_file, low_memory=False)
        df_returns["order-id"] = df_returns["order-id"].astype(str)
        df_returns["sku"]      = df_returns["sku"].astype(str)
        df_returns["Con"]      = df_returns["order-id"] + df_returns["sku"]
        df_returns = df_returns.sort_values("Con")

        pivot_returns = pd.pivot_table(
            df_returns,
            index=["return-date","Con","order-id","sku","asin","fnsku","product-name","fulfillment-center-id","detailed-disposition"],
            values="quantity",
            aggfunc="sum",
        ).reset_index()
        pivot_returns.columns.name = None

        # Lookup: Order Payment
        pivot_orders["Con"] = pivot_orders["Con"].astype(str)
        lookup_orders = pivot_orders[["Con", "total"]].rename(columns={"total": "Order Payment"})
        pivot_returns = pivot_returns.merge(lookup_orders, on="Con", how="left")

        # Lookup: Refund Payment
        pivot_refunds["Con"] = pivot_refunds["Con"].astype(str)
        pivot_returns["Con"] = pivot_returns["Con"].astype(str)
        lookup_refunds = pivot_refunds[["Con", "total"]].rename(columns={"total": "Refund Payment"})
        pivot_returns = pivot_returns.merge(lookup_refunds, on="Con", how="left")

        # Initial Reverse Logistic Charges (before NaN handling)
        pivot_returns["Reverse Logistic Charges"] = (
            pivot_returns["Order Payment"] + pivot_returns["Refund Payment"]
        )

        # Step 1: Where Order Payment is NaN → set BOTH Order Payment AND Refund Payment to 0
        pivot_returns.loc[
            pivot_returns["Order Payment"].isna(),
            ["Order Payment", "Refund Payment"]
        ] = 0

        # Step 2: Where Refund Payment is NaN → set BOTH Refund Payment AND Order Payment to 0
        pivot_returns.loc[
            pivot_returns["Refund Payment"].isna(),
            ["Refund Payment", "Order Payment"]
        ] = 0

        # Ensure numeric and recalculate Reverse Logistic Charges
        pivot_returns["Order Payment"]  = pd.to_numeric(pivot_returns["Order Payment"],  errors="coerce").fillna(0)
        pivot_returns["Refund Payment"] = pd.to_numeric(pivot_returns["Refund Payment"], errors="coerce").fillna(0)
        pivot_returns["Reverse Logistic Charges"] = (
            pivot_returns["Order Payment"] + pivot_returns["Refund Payment"]
        )

        # Scale positive charges by 25%
        mask = pivot_returns["Reverse Logistic Charges"] > 0
        pivot_returns.loc[mask, "Reverse Logistic Charges"] = (
            pivot_returns.loc[mask, "Reverse Logistic Charges"] * 0.25
        )

        # PM lookup
        df_pm = pd.read_excel(pm_file)
        df_pm["ASIN"] = df_pm["ASIN"].astype(str).str.strip()
        pivot_returns["asin"] = pivot_returns["asin"].astype(str).str.strip()
        lookup_pm = df_pm[["ASIN", "Brand", "Brand Manager"]].copy()
        pivot_returns = pivot_returns.merge(lookup_pm, left_on="asin", right_on="ASIN", how="left")
        pivot_returns.drop(columns=["ASIN"], inplace=True, errors="ignore")

        # Format date
        pivot_returns["return-date"] = pd.to_datetime(pivot_returns["return-date"], errors="coerce").dt.date

        # Brand Reverse Logistic Pivot
        pivot_brand = pd.pivot_table(
            pivot_returns,
            index="Brand",
            values="Reverse Logistic Charges",
            aggfunc="sum",
            margins=True,
            margins_name="Grand Total",
        ).reset_index()
        pivot_brand.columns.name = None

    st.success("✅ Reports generated successfully!")

    # ─── Display Tabs ─────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs(["📋 Orders Pivot", "↩️ Refunds Pivot", "🔄 Returns Analysis", "🏷️ Brand Reverse Logistics"])

    with tab1:
        st.subheader(f"Orders Pivot  —  {len(pivot_orders):,} rows")
        st.dataframe(pivot_orders, use_container_width=True, height=400)
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("⬇️ Download CSV", to_csv_bytes(pivot_orders),
                               "pivot_orders.csv", "text/csv", key="dl_orders_csv")
        with col2:
            st.download_button("⬇️ Download Excel", to_excel_bytes({"Orders Pivot": pivot_orders}),
                               "pivot_orders.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_orders_xlsx")

    with tab2:
        st.subheader(f"Refunds Pivot  —  {len(pivot_refunds):,} rows")
        st.dataframe(pivot_refunds, use_container_width=True, height=400)
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("⬇️ Download CSV", to_csv_bytes(pivot_refunds),
                               "pivot_refunds.csv", "text/csv", key="dl_refunds_csv")
        with col2:
            st.download_button("⬇️ Download Excel", to_excel_bytes({"Refunds Pivot": pivot_refunds}),
                               "pivot_refunds.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_refunds_xlsx")

    with tab3:
        st.subheader(f"Returns Analysis  —  {len(pivot_returns):,} rows")

        # Quick filter
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            brands = ["All"] + sorted(pivot_returns["Brand"].dropna().unique().tolist())
            sel_brand = st.selectbox("Filter by Brand", brands)
        with col_f2:
            disps = ["All"] + sorted(pivot_returns["detailed-disposition"].dropna().unique().tolist())
            sel_disp = st.selectbox("Filter by Disposition", disps)

        df_view = pivot_returns.copy()
        if sel_brand != "All":
            df_view = df_view[df_view["Brand"] == sel_brand]
        if sel_disp != "All":
            df_view = df_view[df_view["detailed-disposition"] == sel_disp]

        st.dataframe(df_view, use_container_width=True, height=400)
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("⬇️ Download CSV", to_csv_bytes(pivot_returns),
                               "returns_analysis.csv", "text/csv", key="dl_ret_csv")
        with col2:
            st.download_button("⬇️ Download Excel", to_excel_bytes({"Returns Analysis": pivot_returns}),
                               "returns_analysis.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_ret_xlsx")

    with tab4:
        st.subheader("Brand-wise Reverse Logistics Charges")
        st.dataframe(pivot_brand, use_container_width=True, height=500)
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("⬇️ Download CSV", to_csv_bytes(pivot_brand),
                               "brand_reverse_logistics.csv", "text/csv", key="dl_brand_csv")
        with col2:
            st.download_button("⬇️ Download Excel", to_excel_bytes({"Brand Reverse Logistics": pivot_brand}),
                               "brand_reverse_logistics.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_brand_xlsx")

    # ── Download ALL in one Excel ──────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📥 Download All Reports in One Excel File")
    all_excel = to_excel_bytes({
        "Orders Pivot":           pivot_orders,
        "Refunds Pivot":          pivot_refunds,
        "Returns Analysis":       pivot_returns,
        "Brand Reverse Logistics": pivot_brand,
    })
    st.download_button(
        "⬇️ Download All Reports (Excel)",
        all_excel,
        "all_amazon_reports.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

else:
    st.info("👈 Please upload all three files in the sidebar to begin.")
    st.markdown("""
    ### Required Files
    | # | File | Description |
    |---|------|-------------|
    | 1 | **Unified Transaction CSV** | Amazon Unified Transaction Report (header rows will be skipped automatically) |
    | 2 | **Returns CSV** | Amazon FBA Returns Report |
    | 3 | **PM Master XLSX** | Product Master file with ASIN, Brand, and Brand Manager columns |

    ### Generated Reports
    - **Orders Pivot** — Quantity & Total per Order+SKU combination  
    - **Refunds Pivot** — Quantity & Total for refunded items  
    - **Returns Analysis** — Returns with Order Payment, Refund Payment, Reverse Logistic Charges, Brand & Brand Manager  
    - **Brand Reverse Logistics** — Brand-wise sum of Reverse Logistic Charges  
    """)