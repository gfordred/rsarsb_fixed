import streamlit as st
import pandas as pd
import altair as alt
from datetime import date
import os
import io
import base64
from matplotlib import pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.patches as mpatches

from rsarsb_book_value import calculate_bond_metrics, get_rsa_rsb_rate
from pdf_export import create_bond_pdf_report, get_download_link


def load_rates_data():
    """Load the RSA RSB rates from Excel file."""
    try:
        rates_file = "rsarsb_rates.xlsx"
        if os.path.exists(rates_file):
            df = pd.read_excel(rates_file)
            df['RSB Rate Publish Date'] = pd.to_datetime(df['RSB Rate Publish Date'])
            return df
        else:
            return None
    except Exception as e:
        st.error(f"Error loading rates file: {e}")
        return None


def overview_tab():
    """Display overview with latest rates and visualizations."""
    st.markdown("<h2 style='font-weight: bold; font-size: 32px;'>Market Overview</h2>", unsafe_allow_html=True)
    
    rates_df = load_rates_data()
    
    if rates_df is None or rates_df.empty:
        st.warning("Unable to load rates data from rsarsb_rates.xlsx")
        return
    
    latest_rates = rates_df.iloc[-1]
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            label="2-Year RSA RSB Rate",
            value=f"{latest_rates['RSARSB2']:.4f}%" if pd.notna(latest_rates['RSARSB2']) else "N/A",
            help=f"As of {latest_rates['RSB Rate Publish Date'].strftime('%Y-%m-%d')}"
        )
    
    with col2:
        st.metric(
            label="3-Year RSA RSB Rate",
            value=f"{latest_rates['RSARSB3']:.4f}%" if pd.notna(latest_rates['RSARSB3']) else "N/A",
            help=f"As of {latest_rates['RSB Rate Publish Date'].strftime('%Y-%m-%d')}"
        )
    
    with col3:
        st.metric(
            label="5-Year RSA RSB Rate",
            value=f"{latest_rates['RSARSB5']:.4f}%" if pd.notna(latest_rates['RSARSB5']) else "N/A",
            help=f"As of {latest_rates['RSB Rate Publish Date'].strftime('%Y-%m-%d')}"
        )
    
    st.markdown("---")
    
    st.markdown("<h3 style='font-weight: bold; font-size: 24px;'>Latest Rates Summary</h3>", unsafe_allow_html=True)
    
    # Table in full width
    summary_data = {
        "Term": ["2-Year", "3-Year", "5-Year"],
        "RSA RSB Rate": [
            f"{latest_rates['RSARSB2']:.4f}%" if pd.notna(latest_rates['RSARSB2']) else "N/A",
            f"{latest_rates['RSARSB3']:.4f}%" if pd.notna(latest_rates['RSARSB3']) else "N/A",
            f"{latest_rates['RSARSB5']:.4f}%" if pd.notna(latest_rates['RSARSB5']) else "N/A"
        ],
        "Government Bond Yield": [
            f"{latest_rates['GTZAR2']:.4f}%" if pd.notna(latest_rates['GTZAR2']) else "N/A",
            f"{latest_rates['GTZAR3']:.4f}%" if pd.notna(latest_rates['GTZAR3']) else "N/A",
            f"{latest_rates['GTZAR5']:.4f}%" if pd.notna(latest_rates['GTZAR5']) else "N/A"
        ],
        "Inflation-Linked 3Y": [
            f"{latest_rates['RSAILRSB3']:.4f}%" if pd.notna(latest_rates['RSAILRSB3']) else "N/A",
            f"{latest_rates['RSAILRSB3']:.4f}%" if pd.notna(latest_rates['RSAILRSB3']) else "N/A",
            f"{latest_rates['RSAILRSB5']:.4f}%" if pd.notna(latest_rates['RSAILRSB5']) else "N/A"
        ]
    }
    
    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df, width='stretch', hide_index=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Yield curve chart in its own full-width row
    latest_date = latest_rates['RSB Rate Publish Date']
    st.markdown(f"<h3 style='font-weight: bold; font-size: 20px;'>RSA RSB Yield Curves (as of {latest_date.strftime('%Y-%m-%d')})</h3>", unsafe_allow_html=True)
    
    # Get historical yield curves
    current_date = latest_rates['RSB Rate Publish Date']
    
    # Find rates for 1, 3, 6, 12 months ago
    historical_periods = {
        'Current': 0,
        '1M ago': 1,
        '3M ago': 3,
        '6M ago': 6,
        '12M ago': 12
    }
    
    all_curves = []
    
    for period_name, months_ago in historical_periods.items():
        target_date = current_date - pd.DateOffset(months=months_ago)
        
        # Find closest date in rates_df
        rates_df_sorted = rates_df.sort_values('RSB Rate Publish Date')
        closest_idx = (rates_df_sorted['RSB Rate Publish Date'] - target_date).abs().idxmin()
        period_rates = rates_df_sorted.loc[closest_idx]
        
        for term, term_label in [(2, '2Y'), (3, '3Y'), (5, '5Y')]:
            rate_col = f'RSARSB{term}'
            if pd.notna(period_rates[rate_col]):
                all_curves.append({
                    'Term': term,
                    'Yield': period_rates[rate_col],
                    'Period': period_name
                })
    
    yield_curve_df = pd.DataFrame(all_curves)
    
    if not yield_curve_df.empty:
        y_min = yield_curve_df['Yield'].min() * 0.95
        y_max = yield_curve_df['Yield'].max() * 1.05
        
        # Enhanced color scale with better visibility and distinction
        color_scale = alt.Scale(
            domain=['Current', '1M ago', '3M ago', '6M ago', '12M ago'],
            range=['#FF9500', '#00D4FF', '#00FF88', '#FFD700', '#FF1493']
        )
        
        yield_curve_chart = (
            alt.Chart(yield_curve_df)
            .mark_line(point=alt.OverlayMarkDef(size=120, filled=True), strokeWidth=3.5)
            .encode(
                x=alt.X('Term:Q', title='Term (Years)', scale=alt.Scale(domain=[1.5, 5.5]), 
                       axis=alt.Axis(values=[2, 3, 5], labelFontSize=13, titleFontSize=15, titleFontWeight='bold', labelColor='#E8E8E8', titleColor='#E8E8E8')),
                y=alt.Y('Yield:Q', title='Yield (%)', scale=alt.Scale(domain=[y_min, y_max]), 
                       axis=alt.Axis(labelFontSize=13, titleFontSize=15, titleFontWeight='bold', labelColor='#E8E8E8', titleColor='#E8E8E8')),
                color=alt.Color('Period:N', scale=color_scale, 
                               legend=alt.Legend(title='Period', titleFontSize=14, labelFontSize=12, 
                                               titleFontWeight='bold', orient='right', 
                                               labelColor='#E8E8E8', titleColor='#FF9500')),
                tooltip=[
                    alt.Tooltip('Period:N', title='Period'),
                    alt.Tooltip('Term:Q', title='Term (Years)'),
                    alt.Tooltip('Yield:Q', title='Yield (%)', format='.4f')
                ]
            )
            .properties(height=400)
            .configure_view(strokeWidth=0)
            .configure_axis(gridColor='#2C2C2C', domainColor='#5C5C5C')
            .interactive()
        )
        
        st.altair_chart(yield_curve_chart, width='stretch')
    else:
        st.info("Insufficient data for historical yield curves")
    
    st.markdown("---")
    st.markdown("<h3 style='font-weight: bold; font-size: 24px;'>Historical Rate Trends</h3>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        latest_date_str = latest_rates['RSB Rate Publish Date'].strftime('%Y-%m-%d')
        st.markdown(f"<p style='font-weight: bold; font-size: 18px;'>RSA RSB Rates Over Time (Latest: {latest_date_str})</p>", unsafe_allow_html=True)
        
        chart_data = rates_df[['RSB Rate Publish Date', 'RSARSB2', 'RSARSB3', 'RSARSB5']].copy()
        chart_data = chart_data.dropna(subset=['RSARSB2', 'RSARSB3', 'RSARSB5'])
        
        melted_data = chart_data.melt(
            id_vars=['RSB Rate Publish Date'],
            value_vars=['RSARSB2', 'RSARSB3', 'RSARSB5'],
            var_name='Term',
            value_name='Rate'
        )
        
        melted_data['Term'] = melted_data['Term'].map({
            'RSARSB2': '2-Year',
            'RSARSB3': '3-Year',
            'RSARSB5': '5-Year'
        })
        
        y_min = melted_data['Rate'].min() * 0.95
        y_max = melted_data['Rate'].max() * 1.05
        
        chart = alt.Chart(melted_data).mark_line(point=False, strokeWidth=2.5).encode(
            x=alt.X('RSB Rate Publish Date:T', title='Date', axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold')),
            y=alt.Y('Rate:Q', title='Rate (%)', scale=alt.Scale(domain=[y_min, y_max]), axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold')),
            color=alt.Color('Term:N', 
                legend=alt.Legend(title="Term", titleFontSize=14, labelFontSize=12, titleFontWeight='bold'),
                scale=alt.Scale(range=['#FF9500', '#00A0E3', '#E8E8E8'])
            ),
            tooltip=[
                alt.Tooltip('RSB Rate Publish Date:T', title='Date', format='%Y-%m-%d'),
                alt.Tooltip('Term:N', title='Term'),
                alt.Tooltip('Rate:Q', title='Rate', format='.4f')
            ]
        ).properties(height=400).configure_view(
            strokeWidth=0
        ).configure_axis(
            gridColor='#2C2C2C',
            domainColor='#5C5C5C'
        ).interactive()
        
        st.altair_chart(chart, width='stretch')
    
    with col2:
        st.markdown(f"<p style='font-weight: bold; font-size: 18px;'>Government Bond Yields vs RSA RSB Rates (5-Year, Latest: {latest_date_str})</p>", unsafe_allow_html=True)
        
        comparison_data = rates_df[['RSB Rate Publish Date', 'GTZAR5', 'RSARSB5']].copy()
        comparison_data = comparison_data.dropna(subset=['GTZAR5', 'RSARSB5'])
        comparison_data['Spread'] = comparison_data['RSARSB5'] - comparison_data['GTZAR5']
        
        melted_comparison = comparison_data.melt(
            id_vars=['RSB Rate Publish Date', 'Spread'],
            value_vars=['GTZAR5', 'RSARSB5'],
            var_name='Instrument',
            value_name='Rate'
        )
        
        melted_comparison['Instrument'] = melted_comparison['Instrument'].map({
            'GTZAR5': 'Gov Bond 5Y',
            'RSARSB5': 'RSA RSB 5Y'
        })
        
        rate_min = melted_comparison['Rate'].min() * 0.95
        rate_max = melted_comparison['Rate'].max() * 1.05
        spread_min = comparison_data['Spread'].min() * 0.9
        spread_max = comparison_data['Spread'].max() * 1.1
        
        base = alt.Chart(comparison_data).encode(
            x=alt.X('RSB Rate Publish Date:T', title='Date', axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold'))
        )
        
        rates_chart = alt.Chart(melted_comparison).mark_line(point=False, strokeWidth=2.5).encode(
            x=alt.X('RSB Rate Publish Date:T', title='Date', axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold')),
            y=alt.Y('Rate:Q', title='Rate (%)', scale=alt.Scale(domain=[rate_min, rate_max]), axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold')),
            color=alt.Color('Instrument:N', 
                legend=alt.Legend(title="Instrument", titleFontSize=14, labelFontSize=12, titleFontWeight='bold'),
                scale=alt.Scale(range=['#00A0E3', '#E8E8E8'])
            ),
            tooltip=[
                alt.Tooltip('RSB Rate Publish Date:T', title='Date', format='%Y-%m-%d'),
                alt.Tooltip('Instrument:N', title='Instrument'),
                alt.Tooltip('Rate:Q', title='Rate', format='.4f')
            ]
        )
        
        spread_chart = base.mark_line(point=False, strokeDash=[5, 5], strokeWidth=2.5, color='#FF9500').encode(
            y=alt.Y('Spread:Q', title='Spread (%)', scale=alt.Scale(domain=[spread_min, spread_max]), axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold')),
            tooltip=[
                alt.Tooltip('RSB Rate Publish Date:T', title='Date', format='%Y-%m-%d'),
                alt.Tooltip('Spread:Q', title='Spread', format='.4f')
            ]
        )
        
        comparison_chart = alt.layer(
            rates_chart,
            spread_chart
        ).resolve_scale(
            y='independent'
        ).properties(height=400).configure_view(
            strokeWidth=0
        ).configure_axis(
            gridColor='#2C2C2C',
            domainColor='#5C5C5C'
        ).interactive()
        
        st.altair_chart(comparison_chart, width='stretch')
    
    st.markdown("---")
    
    recent_data = rates_df.tail(12)
    twelve_months_ago = recent_data.iloc[0]['RSB Rate Publish Date'] if not recent_data.empty else latest_rates['RSB Rate Publish Date']
    st.markdown(f"<h3 style='font-weight: bold; font-size: 24px;'>Rate Statistics (Last 12 Months: {twelve_months_ago.strftime('%Y-%m-%d')} to {latest_date_str})</h3>", unsafe_allow_html=True)
    
    stats_data = {
        "Metric": ["Current", "Average", "Minimum", "Maximum", "Std Dev"],
        "2-Year": [
            f"{latest_rates['RSARSB2']:.4f}%" if pd.notna(latest_rates['RSARSB2']) else "N/A",
            f"{recent_data['RSARSB2'].mean():.4f}%" if recent_data['RSARSB2'].notna().any() else "N/A",
            f"{recent_data['RSARSB2'].min():.4f}%" if recent_data['RSARSB2'].notna().any() else "N/A",
            f"{recent_data['RSARSB2'].max():.4f}%" if recent_data['RSARSB2'].notna().any() else "N/A",
            f"{recent_data['RSARSB2'].std():.4f}%" if recent_data['RSARSB2'].notna().any() else "N/A"
        ],
        "3-Year": [
            f"{latest_rates['RSARSB3']:.4f}%" if pd.notna(latest_rates['RSARSB3']) else "N/A",
            f"{recent_data['RSARSB3'].mean():.4f}%" if recent_data['RSARSB3'].notna().any() else "N/A",
            f"{recent_data['RSARSB3'].min():.4f}%" if recent_data['RSARSB3'].notna().any() else "N/A",
            f"{recent_data['RSARSB3'].max():.4f}%" if recent_data['RSARSB3'].notna().any() else "N/A",
            f"{recent_data['RSARSB3'].std():.4f}%" if recent_data['RSARSB3'].notna().any() else "N/A"
        ],
        "5-Year": [
            f"{latest_rates['RSARSB5']:.4f}%" if pd.notna(latest_rates['RSARSB5']) else "N/A",
            f"{recent_data['RSARSB5'].mean():.4f}%" if recent_data['RSARSB5'].notna().any() else "N/A",
            f"{recent_data['RSARSB5'].min():.4f}%" if recent_data['RSARSB5'].notna().any() else "N/A",
            f"{recent_data['RSARSB5'].max():.4f}%" if recent_data['RSARSB5'].notna().any() else "N/A",
            f"{recent_data['RSARSB5'].std():.4f}%" if recent_data['RSARSB5'].notna().any() else "N/A"
        ]
    }
    
    stats_df = pd.DataFrame(stats_data)
    st.dataframe(stats_df, width='stretch', hide_index=True)


def calculator_tab():
    """Display bond calculator with inputs and results."""
    st.subheader("Bond Parameters")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        start_date = st.date_input(
            "Investment Start Date",
            value=date(2026, 2, 20),
            key="start_date",
        )
    
    with col2:
        principal = st.number_input(
            "Principal Investment (ZAR)",
            min_value=1000.0,
            value=1_000_000.0,
            step=1000.0,
            format="%.2f",
            key="principal",
        )
    
    with col3:
        term = st.selectbox(
            "Term (in years)",
            options=[2, 3, 5],
            index=1,
            key="term",
        )
    
    with col4:
        interest_payment_type = st.selectbox(
            "Interest Payment Type",
            options=["semi_annual", "monthly", "reinvest"],
            index=2,
            format_func=lambda x: x.replace("_", " ").title(),
            key="interest_payment_type",
        )
    
    st.markdown("---")
    
    with st.expander("📖 **Calculation Methodology & Examples**", expanded=False):
        st.markdown("""
        <div style='color: #E8E8E8; font-size: 14px;'>
        
        ### Bond Calculation Logic
        
        The RSA Retail Savings Bond calculator uses a **day-count convention** to accurately compute interest accrual and cash flows over the bond's lifetime.
        
        #### **Step 1: Rate Determination**
        The fixed interest rate is determined based on:
        - **Investment Start Date**: The month when the investment begins
        - **Term**: 2, 3, or 5 years
        - **Rate Source**: Historical RSA RSB rates from the `rsarsb_rates.xlsx` file
        
        #### **Step 2: Interest Accrual**
        Interest accrues **daily** using the formula:
        
        ```
        Daily Interest = Principal Balance × (Annual Rate / 365)
        ```
        
        #### **Step 3: Payment Frequency & Interest Calculation**
        
        **Semi-Annual Payments:**
        - Interest paid every **6 months** (typically March 31 and September 30)
        - **Full period interest** = Principal × (Annual Rate × 1/2)
        - For a full 6-month period: Interest = Principal × Rate × 0.5
        - **Same-period investment rule**: If investment is made in a period that also has a payment date, the interest for that partial period is **carried forward** to the next payment
        - Example: Investment on Aug 15, payment date Aug 31 → Aug 15-31 interest paid on Feb 28
        - **March investments**: Short stub from investment date to Mar 31 is added to the Sep 30 payment
        - **September investments**: Short stub from investment date to Sep 30 is added to the Mar 31 payment
        - Example: Investment on Mar 15 → Mar 15-31 interest (16 days) paid on Sep 30 along with Apr 1-Sep 30 interest
        - Principal returned at maturity
        - Cash flows occur on payment dates
        
        **Monthly Payments:**
        - Interest paid every **month**
        - **Full month interest** = Principal × (Annual Rate × 1/12)
        - For a full month: Interest = Principal × Rate / 12
        - **Same-month investment rule**: If investment is made in a month that also has a payment date at month-end, the interest for that partial month is **carried forward** to the next payment period
        - Example: Investment on Feb 15, payment date Feb 28 → Feb 15-28 interest paid on Mar 31
        - Principal returned at maturity
        - More frequent cash flows
        
        **Reinvestment (Capitalisation):**
        - Interest **added to principal** periodically (semi-annually or monthly)
        - Compounding effect increases returns
        - Principal + accumulated interest returned at maturity
        
        #### **Step 4: Stub Period Rules**
        
        **Short Stub Period:**
        - Occurs when investment starts mid-period
        - Interest calculated for **actual days** until next payment date
        - Formula: Interest = Principal × Rate × (Days / 365)
        - Payment occurs in the **following period** on the regular payment date
        
        **Long Stub Period:**
        - Occurs when a period extends beyond normal payment frequency
        - Interest accrues for the extended period
        - Payment made at the end of the long stub
        
        **Payment Timing Rules:**
        - Interest accrues from investment start date
        - First payment may be a **short stub** if starting mid-period
        - **Critical**: If investment date is in the same month/period as a payment date, that partial period's interest is **deferred** to the following payment
        - This prevents same-day or very short-period interest payments
        - Subsequent payments follow regular schedule (monthly or semi-annual)
        - Interest for a stub period is paid on the **next scheduled payment date**
        - Final payment includes any remaining accrued interest plus principal
        
        #### **Step 5: Book Value Calculation**
        
        ```
        Book Value = Principal Balance + Accrued Interest
        ```
        
        The book value represents the total value of the investment at any point in time.
        
        ---
        
        ### **Example Calculation**
        
        **Scenario:**
        - Principal: **R 1,000,000.00**
        - Term: **3 years**
        - Rate: **7.7500%** p.a.
        - Payment Type: **Reinvest**
        - Start Date: **2023-10-20**
        
        **Year 1 (First 6 months):**
        - Days: 183 days
        - Interest Accrued: R 1,000,000 × 7.75% × (183/365) = **R 38,835.62**
        - New Principal (after capitalisation): R 1,000,000 + R 38,835.62 = **R 1,038,835.62**
        
        **Year 2 (Next 6 months):**
        - Principal: R 1,038,835.62
        - Interest Accrued: R 1,038,835.62 × 7.75% × (182/365) = **R 40,186.44**
        - New Principal: R 1,038,835.62 + R 40,186.44 = **R 1,079,022.06**
        
        **At Maturity (3 years):**
        - Final Book Value: **~R 1,254,370.58**
        - Total Interest Earned: **R 254,370.58**
        - Effective Annual Return: **~7.85%** (due to compounding)
        
        ---
        
        ### **Key Metrics Calculated**
        
        1. **Principal Balance**: Outstanding principal at any date
        2. **Accrued Interest**: Interest accumulated but not yet paid/capitalised
        3. **Book Value**: Total investment value (Principal + Accrued Interest)
        4. **Cash Flows**: Payments received (interest and/or principal)
        5. **Total Return**: Final value minus initial investment
        6. **Effective Yield**: Annualised return including compounding effects
        
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    col_calc, col_export = st.columns([3, 1])
    
    with col_calc:
        calculate_button = st.button("Calculate Bond Metrics", type="primary")
    
    with col_export:
        # Enable PDF export if calculations exist in session state
        pdf_enabled = 'bond_results' in st.session_state
        export_pdf_button = st.button("📄 Export to PDF", disabled=not pdf_enabled, 
                                      help="Calculate bond metrics first to enable PDF export" if not pdf_enabled else "Download PDF report")
    
    if calculate_button:
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
        # Store bond parameters in session state for PDF export
        st.session_state['bond_params'] = {
            'principal': principal,
            'term': term,
            'start_date': start_date,
            'interest_payment_type': interest_payment_type
        }
        display_fixed_rate_results(bond)
    elif export_pdf_button and 'bond_results' in st.session_state:
        # Generate and download PDF
        try:
            results = st.session_state['bond_results']
            pdf_buffer = create_bond_pdf_report(
                st.session_state['bond_params'],
                results['rate'],
                results['daily_metrics_df'],
                results['cash_flows_df'],
                results['performance_metrics']
            )
            
            # Create download link
            st.markdown(get_download_link(pdf_buffer, 
                f"RSA_Bond_Analysis_{st.session_state['bond_params']['start_date'].strftime('%Y%m%d')}.pdf"),
                unsafe_allow_html=True)
            st.success("✅ PDF report generated successfully! Click the link above to download.")
        except Exception as e:
            st.error(f"Error generating PDF: {e}")
    else:
        st.info("Configure bond parameters above and click 'Calculate Bond Metrics' to see results")


def display_fixed_rate_results(bond: pd.Series):
    """Display bond calculation results professionally."""
    st.subheader("Bond Analysis Results")

    try:
        rate = get_rsa_rsb_rate(bond["start_date"], bond["term"])
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                label="Fixed Rate",
                value=f"{rate:.4%}",
                help=f"Rate for {bond['term']}-year term starting {bond['start_date'].strftime('%B %Y')}"
            )
        
        with col2:
            st.metric(
                label="Principal",
                value=f"R {bond['principal']:,.2f}",
                help="Initial investment amount"
            )
        
        with col3:
            st.metric(
                label="Term",
                value=f"{bond['term']} Years",
                help="Investment duration"
            )
        
        with col4:
            payment_type_display = bond['interest_payment_type'].replace("_", " ").title()
            st.metric(
                label="Payment Type",
                value=payment_type_display,
                help="Interest payment frequency"
            )
            
    except (ValueError, FileNotFoundError) as e:
        st.warning(f"Could not retrieve rate: {e}")
        rate = None

    try:
        daily_metrics_df, cash_flows_df = calculate_bond_metrics(bond)
    except Exception as e:
        st.error(f"Bond metric calculation failed: {e}")
        return

    if daily_metrics_df.empty and cash_flows_df.empty:
        st.info("No data returned from calculate_bond_metrics for this bond.")
        return

    if not cash_flows_df.empty and "Date" in cash_flows_df.columns:
        cash_flows_df = cash_flows_df.copy()
        cash_flows_df["Date"] = pd.to_datetime(cash_flows_df["Date"])

    if not daily_metrics_df.empty and "Date" in daily_metrics_df.columns:
        daily_metrics_df = daily_metrics_df.copy()
        daily_metrics_df["Date"] = pd.to_datetime(daily_metrics_df["Date"])

    st.markdown("---")
    
    if not daily_metrics_df.empty:
        initial_principal = bond['principal']
        payment_type = bond['interest_payment_type']
        
        # For bonds that pay out interest (monthly, semi_annual), sum all cash flows
        # For reinvest bonds, use book value
        if payment_type in ['monthly', 'semi_annual']:
            # Sum all cash flows received (excluding the initial principal investment)
            total_cash_received = 0
            principal_repayment = 0
            interest_payments = 0
            
            if not cash_flows_df.empty and 'Cash_Flow' in cash_flows_df.columns and 'Type' in cash_flows_df.columns:
                # Sum all coupon payments (positive cash flows)
                interest_payments = cash_flows_df[cash_flows_df['Type'] == 'Coupon']['Cash_Flow'].sum()
                # Get principal repayment (positive cash flow)
                principal_repayment = cash_flows_df[cash_flows_df['Type'] == 'Principal Repayment']['Cash_Flow'].sum()
                # Total cash received = interest + principal repayment (exclude negative principal investment)
                total_cash_received = interest_payments + principal_repayment
            
            # For payment bonds, final value is just the interest earned (principal is returned)
            # Total return should be based on interest only since principal is returned
            final_value = interest_payments
            total_interest = interest_payments
            
        else:  # reinvest
            # Find the row with maximum book value (at or just before maturity)
            max_book_value_idx = daily_metrics_df['Book_Value'].idxmax()
            final_metrics = daily_metrics_df.loc[max_book_value_idx]
            
            # Get total cash flows to calculate actual return
            total_cash_received = 0
            if not cash_flows_df.empty and 'Cash_Flow' in cash_flows_df.columns:
                total_cash_received = cash_flows_df['Cash_Flow'].sum()
            
            # Calculate final value: either book value or total cash received
            final_value = max(final_metrics.get('Book_Value', 0), total_cash_received)
            total_interest = final_metrics.get('Total_Coupons_Paid', 0) + final_metrics.get('Total_Coupons_Capitalised', 0)
        
        st.markdown("<h3 style='font-weight: bold; font-size: 24px;'>Performance Metrics</h3>", unsafe_allow_html=True)
        
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            if payment_type in ['monthly', 'semi_annual']:
                st.metric(
                    label="Total Cash Received",
                    value=f"R {total_cash_received:,.2f}",
                    help="Total cash received (principal repayment + all interest payments)"
                )
            else:
                st.metric(
                    label="Final Value at Maturity",
                    value=f"R {final_value:,.2f}",
                    help="Total value at maturity (principal + capitalised interest)"
                )
        
        with col2:
            st.metric(
                label="Total Interest Earned",
                value=f"R {total_interest:,.2f}",
                help="Total interest earned over bond lifetime"
            )
        
        with col3:
            # For payment bonds, return is interest only (principal is returned)
            # For reinvest bonds, return is final value minus initial principal
            if payment_type in ['monthly', 'semi_annual']:
                total_return = total_interest
                return_pct = (total_return / initial_principal) * 100 if initial_principal > 0 else 0
            else:
                total_return = final_value - initial_principal
                return_pct = (total_return / initial_principal) * 100 if initial_principal > 0 else 0
            
            st.metric(
                label="Total Return",
                value=f"{return_pct:.2f}%",
                delta=f"R {total_return:,.2f}",
                help="Total return as percentage of initial investment"
            )
        
        with col4:
            years = bond['term']
            if years > 0 and initial_principal > 0:
                # For payment bonds, calculate based on interest earned
                # For reinvest bonds, calculate based on final value
                if payment_type in ['monthly', 'semi_annual']:
                    # Annualized return based on interest payments
                    effective_annual_return = (total_interest / initial_principal / years) * 100
                else:
                    # Annualized return with compounding
                    if final_value > 0:
                        effective_annual_return = ((final_value / initial_principal) ** (1/years) - 1) * 100
                    else:
                        effective_annual_return = 0
                
                st.metric(
                    label="Effective Annual Return",
                    value=f"{effective_annual_return:.4f}%",
                    help="Annualized return including compounding"
                )
            else:
                st.metric(
                    label="Effective Annual Return",
                    value="N/A",
                    help="Annualized return including compounding"
                )
        
        with col5:
            if rate and years > 0:
                # For payment bonds, compare interest earned to simple interest
                # For reinvest, compare to simple interest compounded
                if payment_type in ['monthly', 'semi_annual']:
                    # Simple interest on principal
                    simple_interest = initial_principal * rate * years
                    benefit = total_interest - simple_interest
                    st.metric(
                        label="vs Simple Interest",
                        value=f"R {benefit:,.2f}",
                        help="Difference vs simple interest (should be ~0 for payment bonds)"
                    )
                else:
                    # For reinvest, compare final value to simple interest total
                    simple_interest_total = initial_principal * (1 + rate * years)
                    benefit = final_value - simple_interest_total
                    st.metric(
                        label="Compounding Benefit",
                        value=f"R {benefit:,.2f}",
                        help="Additional return from compounding vs simple interest"
                    )
            else:
                st.metric(
                    label="Compounding Benefit",
                    value="N/A",
                    help="Additional return from compounding vs simple interest"
                )
        
        # Store results in session state for PDF export
        st.session_state['bond_results'] = {
            'rate': rate,
            'daily_metrics_df': daily_metrics_df,
            'cash_flows_df': cash_flows_df,
            'performance_metrics': {
                'total_cash_received': total_cash_received if payment_type in ['monthly', 'semi_annual'] else final_value,
                'total_interest': total_interest,
                'return_pct': return_pct,
                'effective_annual_return': effective_annual_return if 'effective_annual_return' in locals() else 0
            }
        }

    st.markdown("---")
    
    result_tabs = ["Cash Flow Schedule", "Daily Metrics", "Visualizations"]
    if bond["interest_payment_type"] == "reinvest":
        result_tabs.insert(1, "Reinvestment Schedule")

    tabs = st.tabs(result_tabs)

    cash_flow_tab = tabs[0]
    reinvest_tab = tabs[1] if bond["interest_payment_type"] == "reinvest" else None
    daily_metrics_tab = tabs[2] if bond["interest_payment_type"] == "reinvest" else tabs[1]
    charts_tab = tabs[3] if bond["interest_payment_type"] == "reinvest" else tabs[2]

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
                        width='stretch',
                        column_config={
                            "Date": st.column_config.DateColumn("Date", format="YYYY-MM-DD"),
                            "Principal_Before": st.column_config.NumberColumn("Principal Before", format="R %.2f"),
                            "Interest_Capitalised": st.column_config.NumberColumn("Interest Capitalised", format="R %.2f"),
                            "Principal_After": st.column_config.NumberColumn("Principal After", format="R %.2f"),
                        },
                        hide_index=True,
                    )

    with cash_flow_tab:
        st.subheader("Cash Flow Schedule")
        if cash_flows_df.empty:
            st.info("No cash flows to display.")
        else:
            column_config = {}
            if "Date" in cash_flows_df.columns:
                column_config["Date"] = st.column_config.DateColumn("Date", format="YYYY-MM-DD")
            if "Cash_Flow" in cash_flows_df.columns:
                column_config["Cash_Flow"] = st.column_config.NumberColumn("Cash Flow", format="R %.2f")

            st.dataframe(
                cash_flows_df,
                width='stretch',
                column_config=column_config,
                hide_index=True,
            )

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
                width='stretch',
                column_config=column_config,
                hide_index=True,
            )

    with charts_tab:
        st.subheader("Visualizations")

        if daily_metrics_df.empty:
            st.info("No daily metrics to chart.")
            return

        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("<p style='font-weight: bold; font-size: 16px;'>Book Value Growth Over Time</p>", unsafe_allow_html=True)
            chart_df = daily_metrics_df.copy()
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

                y_min = plot_data['Value'].min() * 0.98
                y_max = plot_data['Value'].max() * 1.02

                chart = (
                    alt.Chart(plot_data)
                    .mark_line(strokeWidth=2.5)
                    .encode(
                        x=alt.X("Date:T", title="Date", axis=alt.Axis(labelFontSize=11, titleFontSize=13, titleFontWeight='bold')),
                        y=alt.Y("Value:Q", title="Value (ZAR)", scale=alt.Scale(domain=[y_min, y_max]), axis=alt.Axis(labelFontSize=11, titleFontSize=13, titleFontWeight='bold')),
                        color=alt.Color("Metric:N", scale=alt.Scale(range=['#FF9500', '#00A0E3']), legend=alt.Legend(titleFontSize=13, labelFontSize=11, titleFontWeight='bold')),
                        tooltip=[
                            alt.Tooltip("Date:T", title="Date", format="%Y-%m-%d"),
                            alt.Tooltip("Metric:N", title="Metric"),
                            alt.Tooltip("Value:Q", title="Value", format=",.2f"),
                        ],
                    )
                    .properties(height=350)
                    .configure_view(strokeWidth=0)
                    .configure_axis(gridColor='#2C2C2C', domainColor='#5C5C5C')
                    .interactive()
                )
                st.altair_chart(chart, width='stretch')
            else:
                st.info("Missing required columns to plot Book Value vs Principal.")

        with col2:
            st.markdown("<p style='font-weight: bold; font-size: 16px;'>Accrued Interest Over Time</p>", unsafe_allow_html=True)
            if "Date" in daily_metrics_df.columns and "Accrued_Interest" in daily_metrics_df.columns:
                accrued_df = daily_metrics_df[["Date", "Accrued_Interest"]].copy()
                
                y_min = accrued_df['Accrued_Interest'].min() * 0.95 if accrued_df['Accrued_Interest'].min() > 0 else 0
                y_max = accrued_df['Accrued_Interest'].max() * 1.05
                
                accrued_chart = (
                    alt.Chart(accrued_df)
                    .mark_area(line={'color': '#00A0E3', 'strokeWidth': 2.5}, color=alt.Gradient(
                        gradient='linear',
                        stops=[alt.GradientStop(color='#00A0E3', offset=0),
                               alt.GradientStop(color='#003D5C', offset=1)],
                        x1=1, x2=1, y1=1, y2=0
                    ))
                    .encode(
                        x=alt.X("Date:T", title="Date", axis=alt.Axis(labelFontSize=11, titleFontSize=13, titleFontWeight='bold')),
                        y=alt.Y("Accrued_Interest:Q", title="Accrued Interest (ZAR)", scale=alt.Scale(domain=[y_min, y_max]), axis=alt.Axis(labelFontSize=11, titleFontSize=13, titleFontWeight='bold')),
                        tooltip=[
                            alt.Tooltip("Date:T", title="Date", format="%Y-%m-%d"),
                            alt.Tooltip("Accrued_Interest:Q", title="Accrued Interest", format=",.2f"),
                        ],
                    )
                    .properties(height=350)
                    .configure_view(strokeWidth=0)
                    .configure_axis(gridColor='#2C2C2C', domainColor='#5C5C5C')
                    .interactive()
                )
                st.altair_chart(accrued_chart, width='stretch')
            else:
                st.info("Missing required columns to plot Accrued Interest.")
        
        st.markdown("---")
        
        st.markdown("<p style='font-weight: bold; font-size: 16px;'>Cumulative Interest Breakdown</p>", unsafe_allow_html=True)
        if "Total_Coupons_Paid" in daily_metrics_df.columns or "Total_Coupons_Capitalised" in daily_metrics_df.columns:
            cumulative_df = daily_metrics_df[["Date"]].copy()
            
            if "Total_Coupons_Paid" in daily_metrics_df.columns:
                cumulative_df["Interest Paid"] = daily_metrics_df["Total_Coupons_Paid"]
            if "Total_Coupons_Capitalised" in daily_metrics_df.columns:
                cumulative_df["Interest Capitalised"] = daily_metrics_df["Total_Coupons_Capitalised"]
            
            melted_cumulative = cumulative_df.melt(id_vars=["Date"], var_name="Type", value_name="Amount")
            
            cumulative_chart = (
                alt.Chart(melted_cumulative)
                .mark_line(strokeWidth=2.5)
                .encode(
                    x=alt.X("Date:T", title="Date", axis=alt.Axis(labelFontSize=11, titleFontSize=13, titleFontWeight='bold')),
                    y=alt.Y("Amount:Q", title="Cumulative Interest (ZAR)", axis=alt.Axis(labelFontSize=11, titleFontSize=13, titleFontWeight='bold')),
                    color=alt.Color("Type:N", scale=alt.Scale(range=['#FF9500', '#00A0E3']), legend=alt.Legend(titleFontSize=13, labelFontSize=11, titleFontWeight='bold')),
                    tooltip=[
                        alt.Tooltip("Date:T", title="Date", format="%Y-%m-%d"),
                        alt.Tooltip("Type:N", title="Type"),
                        alt.Tooltip("Amount:Q", title="Amount", format=",.2f"),
                    ],
                )
                .properties(height=350)
                .configure_view(strokeWidth=0)
                .configure_axis(gridColor='#2C2C2C', domainColor='#5C5C5C')
                .interactive()
            )
            st.altair_chart(cumulative_chart, width='stretch')


def main():
    st.set_page_config(page_title="Fixed Rate RSA Retail Bond Calculator", layout="wide")
    
    st.markdown(
        """
        <style>
            /* Bloomberg-inspired theme */
            @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap');
            
            html, body, [class*="css"] {
                font-family: 'Roboto', sans-serif;
            }
            
            /* Main background */
            .stApp {
                background-color: #0A0A0A;
            }
            
            /* Headers */
            h1, h2, h3, h4, h5, h6 {
                color: #E8E8E8 !important;
                font-weight: bold !important;
            }
            
            h1 {
                font-size: 36px !important;
                margin-bottom: 1rem !important;
            }
            
            /* Metrics */
            .stMetric {
                background-color: #1A1A1A;
                padding: 1rem;
                border-radius: 4px;
                border-left: 3px solid #FF9500;
            }
            
            .stMetric label {
                color: #B8B8B8 !important;
                font-size: 14px !important;
                font-weight: bold !important;
            }
            
            .stMetric [data-testid="stMetricValue"] {
                color: #E8E8E8 !important;
                font-size: 28px !important;
                font-weight: bold !important;
            }
            
            /* Dataframes */
            .stDataFrame {
                background-color: #1A1A1A;
            }
            
            /* Tables */
            table {
                background-color: #1A1A1A !important;
                color: #E8E8E8 !important;
            }
            
            thead tr th {
                background-color: #2A2A2A !important;
                color: #FF9500 !important;
                font-weight: bold !important;
                font-size: 14px !important;
                border-bottom: 2px solid #FF9500 !important;
            }
            
            tbody tr td {
                color: #E8E8E8 !important;
                font-size: 13px !important;
                border-bottom: 1px solid #2A2A2A !important;
            }
            
            tbody tr:hover {
                background-color: #252525 !important;
            }
            
            /* Tabs */
            .stTabs [data-baseweb="tab-list"] {
                gap: 8px;
                background-color: #1A1A1A;
                padding: 0.5rem;
                border-radius: 4px;
            }
            
            .stTabs [data-baseweb="tab"] {
                background-color: #2A2A2A;
                color: #B8B8B8;
                font-weight: bold;
                font-size: 16px;
                border-radius: 4px;
                padding: 0.75rem 1.5rem;
            }
            
            .stTabs [aria-selected="true"] {
                background-color: #FF9500;
                color: #0A0A0A;
            }
            
            /* Input fields */
            .stDateInput, .stNumberInput, .stSelectbox {
                background-color: #1A1A1A;
            }
            
            input, select {
                background-color: #2A2A2A !important;
                color: #E8E8E8 !important;
                border: 1px solid #3A3A3A !important;
                font-size: 14px !important;
            }
            
            label {
                color: #B8B8B8 !important;
                font-weight: bold !important;
                font-size: 13px !important;
            }
            
            /* Buttons */
            .stButton > button {
                background-color: #FF9500;
                color: #0A0A0A;
                font-weight: bold;
                font-size: 16px;
                border: none;
                padding: 0.75rem 2rem;
                border-radius: 4px;
            }
            
            .stButton > button:hover {
                background-color: #E68600;
            }
            
            /* Text */
            p, span, div {
                color: #E8E8E8;
            }
            
            /* Dividers */
            hr {
                border-color: #2A2A2A !important;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )
    
    st.markdown("<h1 style='font-weight: bold; font-size: 36px; color: #E8E8E8;'>Fixed Rate RSA Retail Bond Calculator</h1>", unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["📊 Overview", "🧮 Calculator"])
    
    with tab1:
        overview_tab()
    
    with tab2:
        calculator_tab()


if __name__ == "__main__":
    main()
