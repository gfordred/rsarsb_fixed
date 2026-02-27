# app.py
import streamlit as st
import pandas as pd
import altair as alt
from datetime import date

# Your existing bond engine + rate lookup (kept the same as your original app)
from rsarsb_book_value import calculate_bond_metrics, get_rsa_rsb_rate


def fixed_rate_bond_calculator():
    st.header("Fixed Rate Bond Calculator")
    st.write("Analyse a single **Fixed Rate** bond (no portfolio, no auth, no database).")

    # ---- Inputs (sidebar) ----
    st.sidebar.header("Bond Parameters")

    start_date = st.sidebar.date_input(
        "Investment Start Date",
        value=date(2023, 10, 20),
        key="start_date",
    )

    principal = st.sidebar.number_input(
        "Principal Investment",
        min_value=1000.0,
        value=1_000_000.0,
        step=1000.0,
        format="%.2f",
        key="principal",
    )

    term = st.sidebar.selectbox(
        "Term (in years)",
        options=[2, 3, 5],
        index=1,
        key="term",
    )

    interest_payment_type = st.sidebar.selectbox(
        "Interest Payment Type",
        options=["semi_annual", "monthly", "reinvest"],
        index=0,
        format_func=lambda x: x.replace("_", " ").title(),
        key="interest_payment_type",
    )

    # ---- Calculate ----
    if st.sidebar.button("Calculate Bond Metrics", type="primary"):
        bond = pd.Series(
            {
                "start_date": start_date,
                "term": term,
                "interest_payment_type": interest_payment_type,
                "principal": principal,
                "unique_reference": "Calculator",
                "bond_type": "Fixed Rate",
            }
        )

        display_fixed_rate_results(bond)


def display_fixed_rate_results(bond: pd.Series):
    st.subheader("Results")

    # --- Rate headline metric ---
    try:
        rate = get_rsa_rsb_rate(bond["start_date"], bond["term"])
        st.metric(
            label=f"{bond['term']}-Year Fixed-Rate Yield",
            value=f"{rate:.4%}",
            help=f"Rate for investments starting in {bond['start_date'].strftime('%B %Y')}.",
        )
    except (ValueError, FileNotFoundError) as e:
        rate = None
        st.warning(f"Could not retrieve rate: {e}")

    # --- Core engine outputs ---
    try:
        daily_metrics_df, cash_flows_df = calculate_bond_metrics(bond)
    except Exception as e:
        st.error(f"Bond metric calculation failed: {e}")
        return

    if daily_metrics_df.empty and cash_flows_df.empty:
        st.info("No data returned from calculate_bond_metrics for this bond.")
        return

    # Ensure datetime columns
    if not cash_flows_df.empty and "Date" in cash_flows_df.columns:
        cash_flows_df = cash_flows_df.copy()
        cash_flows_df["Date"] = pd.to_datetime(cash_flows_df["Date"])

    if not daily_metrics_df.empty and "Date" in daily_metrics_df.columns:
        daily_metrics_df = daily_metrics_df.copy()
        daily_metrics_df["Date"] = pd.to_datetime(daily_metrics_df["Date"])

    # ---- Tabs ----
    tab_names = ["Cash Flow Schedule", "Daily Metrics Table", "Charts"]
    if bond["interest_payment_type"] == "reinvest":
        tab_names.insert(1, "Reinvestment Schedule")

    tabs = st.tabs(tab_names)

    cash_flow_tab = tabs[0]
    reinvest_tab = tabs[1] if bond["interest_payment_type"] == "reinvest" else None
    daily_metrics_tab = tabs[2] if bond["interest_payment_type"] == "reinvest" else tabs[1]
    charts_tab = tabs[3] if bond["interest_payment_type"] == "reinvest" else tabs[2]

    # ---- Reinvestment Schedule (only for reinvest) ----
    if reinvest_tab is not None:
        with reinvest_tab:
            st.subheader("Reinvestment Capitalisation Schedule")

            if cash_flows_df.empty:
                st.info("No cash flow schedule available.")
            else:
                capitalisation_events = cash_flows_df[cash_flows_df["Type"] == "Capitalisation"].copy()

                if capitalisation_events.empty:
                    st.info("No capitalisation events found.")
                elif daily_metrics_df.empty:
                    st.info("Daily metrics missing, cannot build reinvestment schedule.")
                else:
                    capitalisation_events.rename(columns={"Cash_Flow": "Interest_Capitalised"}, inplace=True)

                    merged_df = pd.merge_asof(
                        capitalisation_events.sort_values("Date"),
                        daily_metrics_df[["Date", "Principal_Balance"]].sort_values("Date"),
                        on="Date",
                        direction="backward",
                    ).rename(columns={"Principal_Balance": "Principal_After"})

                    merged_df["Prev_Date"] = merged_df["Date"] - pd.Timedelta(days=1)

                    final_merged_df = pd.merge(
                        merged_df,
                        daily_metrics_df[["Date", "Principal_Balance"]],
                        left_on="Prev_Date",
                        right_on="Date",
                        how="left",
                    ).rename(columns={"Principal_Balance": "Principal_Before"})

                    summary_df = final_merged_df[["Date_x", "Principal_Before", "Interest_Capitalised", "Principal_After"]].rename(
                        columns={"Date_x": "Date"}
                    )

                    st.dataframe(
                        summary_df,
                        use_container_width=True,
                        column_config={
                            "Date": st.column_config.DateColumn("Date", format="YYYY-MM-DD"),
                            "Principal_Before": st.column_config.NumberColumn("Principal Before", format="R %.2f"),
                            "Interest_Capitalised": st.column_config.NumberColumn("Interest Capitalised", format="R %.2f"),
                            "Principal_After": st.column_config.NumberColumn("Principal After", format="R %.2f"),
                        },
                        hide_index=True,
                    )

    # ---- Cash Flow Schedule ----
    with cash_flow_tab:
        st.subheader("Cash Flow Schedule")
        if cash_flows_df.empty:
            st.info("No cash flows to display.")
        else:
            # Try to format commonly-used columns
            column_config = {}
            if "Date" in cash_flows_df.columns:
                column_config["Date"] = st.column_config.DateColumn("Date", format="YYYY-MM-DD")
            if "Cash_Flow" in cash_flows_df.columns:
                column_config["Cash_Flow"] = st.column_config.NumberColumn("Cash Flow", format="R %.2f")

            st.dataframe(
                cash_flows_df,
                use_container_width=True,
                column_config=column_config,
                hide_index=True,
            )

    # ---- Daily Metrics ----
    with daily_metrics_tab:
        st.subheader("Daily Metrics")
        if daily_metrics_df.empty:
            st.info("No daily metrics to display.")
        else:
            metric_cols = [
                "Principal_Balance",
                "Coupon_Cash_Flow",
                "Principal_Cash_Flow",
                "Accrued_Interest",
                "Book_Value",
                "Total_Coupons_Paid",
                "Total_Coupons_Capitalised",
            ]
            column_config = {}
            if "Date" in daily_metrics_df.columns:
                column_config["Date"] = st.column_config.DateColumn("Date", format="YYYY-MM-DD")
            for c in metric_cols:
                if c in daily_metrics_df.columns:
                    column_config[c] = st.column_config.NumberColumn(c.replace("_", " "), format="R %.2f")

            st.dataframe(
                daily_metrics_df,
                use_container_width=True,
                column_config=column_config,
                hide_index=True,
            )

    # ---- Charts ----
    with charts_tab:
        st.subheader("Charts")

        if daily_metrics_df.empty:
            st.info("No daily metrics to chart.")
            return

        # Book Value vs Principal
        st.markdown("**Book Value vs Principal Over Time**")
        chart_df = daily_metrics_df.copy()
        # If the last row is maturity terminal-only, you can drop it like in your original logic:
        if len(chart_df) > 1:
            chart_df = chart_df.iloc[:-1].copy()

        melt_cols = [c for c in ["Book_Value", "Principal_Balance"] if c in chart_df.columns]
        if "Date" in chart_df.columns and melt_cols:
            plot_data = chart_df.melt(
                id_vars=["Date"],
                value_vars=melt_cols,
                var_name="Metric",
                value_name="Value",
            )
            plot_data["Metric"] = plot_data["Metric"].replace({"Principal_Balance": "Principal", "Book_Value": "Book Value"})

            chart = (
                alt.Chart(plot_data)
                .mark_line()
                .encode(
                    x=alt.X("Date:T", title="Date"),
                    y=alt.Y("Value:Q", title="Value (ZAR)"),
                    color="Metric:N",
                    tooltip=[
                        alt.Tooltip("Date:T", title="Date"),
                        alt.Tooltip("Metric:N", title="Metric"),
                        alt.Tooltip("Value:Q", title="Value", format=",.2f"),
                    ],
                )
                .interactive()
            )
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("Missing required columns to plot Book Value vs Principal.")

        # Accrued Interest
        st.markdown("**Accrued Interest Over Time**")
        if "Date" in daily_metrics_df.columns and "Accrued_Interest" in daily_metrics_df.columns:
            st.line_chart(daily_metrics_df.set_index("Date")["Accrued_Interest"])
        else:
            st.info("Missing required columns to plot Accrued Interest.")


def main():
    st.set_page_config(page_title="Fixed Rate Bond Calculator", layout="wide")
    st.title("Fixed Rate Bond Calculator (Single Bond)")

    # Optional: light global styling (keep minimal)
    st.markdown(
        """
        <style>
            h1 { margin-bottom: 0.25rem; }
            .stMetric { padding: 0.5rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    fixed_rate_bond_calculator()


if __name__ == "__main__":
    main()