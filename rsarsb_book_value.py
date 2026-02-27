import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta
from dateutil.rrule import rrule, MONTHLY, FR


def get_rsa_rsb_rate(start_date: date, term: int) -> float:
    """
    Looks up the RSA Retail Savings Bond rate from 'rsarsb_rates.xlsx'.

    The rate is determined by the start date's month and year. The function finds
    the corresponding rate published for that month.

    Args:
        start_date: The start date of the bond investment.
        term: The term of the bond in years (2, 3, or 5).

    Returns:
        The fixed-rate yield for the given term and start date.
    """
    # Define the mapping from term to the relevant column name in the Excel file
    term_to_column = {
        2: 'RSARSB2',
        3: 'RSARSB3',
        5: 'RSARSB5'
    }

    if term not in term_to_column:
        raise ValueError("Invalid term. Term must be 2, 3, or 5 years.")

    column_name = term_to_column[term]

    # Load the rates from the Excel file
    try:
        # Load the rates from the first sheet of the Excel file by default
        rates_df = pd.read_excel('rsarsb_rates.xlsx')
    except FileNotFoundError:
        raise FileNotFoundError("The file 'rsarsb_rates.xlsx' was not found.")
    except Exception as e:
        raise Exception(f"An error occurred while reading the Excel file: {e}")

    # Ensure the date column is in datetime format
    rates_df['RSB Rate Publish Date'] = pd.to_datetime(rates_df['RSB Rate Publish Date'])

    # Find the rate for the given start_date's month and year
    lookup_month = start_date.month
    lookup_year = start_date.year

    # Filter the DataFrame for the matching month and year
    rate_row = rates_df[
        (rates_df['RSB Rate Publish Date'].dt.year == lookup_year) &
        (rates_df['RSB Rate Publish Date'].dt.month == lookup_month)
    ]

    if rate_row.empty:
        raise ValueError(f"No rate found for {start_date.strftime('%B %Y')}.")

    # Extract the rate from the specified column
    rate = rate_row.iloc[0][column_name]

    # Adjust for percentage by dividing by 100
    return rate / 100.0


def _calculate_fixed_rate_metrics(
    start_date: date,
    term: int,
    interest_payment_type: str,
    principal: float,
    unique_reference: str,
    rate: float
):
    """
    Calculates daily bond metrics and cash flows for a fixed-income bond.

    Args:
        start_date: The start date of the bond investment.
        term: The term of the bond in years (2, 3, or 5).
        interest_payment_type: 'semi_annual', 'monthly', or 'reinvest'.
        principal: The principal investment amount.
        unique_reference: A unique identifier for the bond.

    Returns:
        A tuple containing:
        - pd.DataFrame: Daily metrics from start to maturity.
        - pd.DataFrame: All expected cash flows.
    """
    maturity_date = start_date + relativedelta(years=term)

    # --- 1. Generate Cash Flow Events (Refactored for Compounding) ---
    cash_flows = [{'Date': start_date, 'Cash_Flow': -principal, 'Type': 'Principal Investment'}]
    current_principal = principal

    # Determine coupon dates
    if interest_payment_type in ['semi_annual', 'monthly', 'reinvest']:
        if interest_payment_type in ['semi_annual', 'reinvest']:
            # Find all potential coupon dates (Mar 31, Sep 30)
            all_coupon_dates = []
            for y in range(start_date.year, maturity_date.year + 2):
                all_coupon_dates.extend([date(y, 3, 31), date(y, 9, 30)])
            all_coupon_dates.sort()
            coupon_dates = [d for d in all_coupon_dates if start_date < d <= maturity_date]
        else: # monthly
            coupon_dates = [d.date() for d in rrule(MONTHLY, dtstart=start_date, until=maturity_date, bymonthday=-1)]
            coupon_dates = [d for d in coupon_dates if d > start_date]

        if not coupon_dates or coupon_dates[-1] < maturity_date:
            coupon_dates.append(maturity_date)

        # Check if first coupon date is in same month/period as start date
        # If so, skip it and defer interest to next payment
        deferred_interest = 0
        skip_first_payment = False
        
        if coupon_dates and interest_payment_type in ['monthly', 'semi_annual']:
            first_coupon = coupon_dates[0]
            
            if interest_payment_type == 'monthly':
                # Same month if year and month match
                if first_coupon.year == start_date.year and first_coupon.month == start_date.month:
                    skip_first_payment = True
                    # Calculate deferred interest for this period
                    days_in_period = (first_coupon - start_date).days
                    deferred_interest = current_principal * rate * days_in_period / 365
            elif interest_payment_type == 'semi_annual':
                # Same period if both dates fall in same semi-annual period
                # Mar 31 period: Oct 1 - Mar 31
                # Sep 30 period: Apr 1 - Sep 30
                start_period = 'Mar' if 1 <= start_date.month <= 3 or start_date.month >= 10 else 'Sep'
                first_coupon_period = 'Mar' if first_coupon.month == 3 else 'Sep'
                
                # Check if in same semi-annual period
                if start_period == first_coupon_period:
                    if start_period == 'Mar' and start_date.month >= 10:
                        # Oct-Dec of previous year to Mar of current year
                        if first_coupon.year == start_date.year or (first_coupon.year == start_date.year + 1 and first_coupon.month == 3):
                            skip_first_payment = True
                            days_in_period = (first_coupon - start_date).days
                            deferred_interest = current_principal * rate * days_in_period / 365
                    elif start_period == 'Sep' and 4 <= start_date.month <= 9:
                        # Apr-Sep of same year
                        if first_coupon.year == start_date.year and first_coupon.month == 9:
                            skip_first_payment = True
                            days_in_period = (first_coupon - start_date).days
                            deferred_interest = current_principal * rate * days_in_period / 365

        last_coupon_dt = start_date
        for idx, coupon_dt in enumerate(coupon_dates):
            # Skip first payment if in same month/period
            if idx == 0 and skip_first_payment:
                last_coupon_dt = coupon_dt
                continue
                
            days_in_period = (coupon_dt - last_coupon_dt).days
            coupon_amount = 0

            if interest_payment_type in ['semi_annual', 'reinvest']:
                # Full semi-annual period
                if 180 <= days_in_period <= 185:
                    coupon_amount = current_principal * rate / 2
                # Stub period
                else:
                    coupon_amount = current_principal * rate * days_in_period / 365
            elif interest_payment_type == 'monthly':
                # Full month period
                if 28 <= days_in_period <= 31:
                    coupon_amount = current_principal * rate / 12
                # Stub period
                else:
                    coupon_amount = current_principal * rate * days_in_period / 365

            # Add deferred interest to second payment (first actual payment)
            if idx == 1 and skip_first_payment and deferred_interest > 0:
                coupon_amount += deferred_interest

            # For reinvestment bonds, the coupon is not a cash flow but a capitalisation event.
            # We'll record it with a different type to distinguish it from actual cash flows.
            event_type = 'Capitalisation' if interest_payment_type == 'reinvest' else 'Coupon'
            cash_flows.append({'Date': coupon_dt, 'Cash_Flow': coupon_amount, 'Type': event_type})
            
            if interest_payment_type == 'reinvest':
                current_principal += coupon_amount
            
            last_coupon_dt = coupon_dt

    # Add final principal repayment
    if interest_payment_type == 'reinvest':
        # For reinvestment bonds, split the final payout into principal and capitalized interest
        total_capitalized_interest = current_principal - principal
        cash_flows.append({'Date': maturity_date, 'Cash_Flow': principal, 'Type': 'Principal Repayment'})
        if total_capitalized_interest > 0:
            # The final payout of accumulated interest is treated as a standard coupon payment for cash flow purposes.
            cash_flows.append({'Date': maturity_date, 'Cash_Flow': total_capitalized_interest, 'Type': 'Coupon'})
    else:
        # For other bonds, just repay the original principal
        cash_flows.append({'Date': maturity_date, 'Cash_Flow': principal, 'Type': 'Principal Repayment'})

    cash_flows_df = pd.DataFrame(cash_flows).sort_values(by='Date').reset_index(drop=True)

    # --- 2. Generate Daily Metrics ---
    dates = pd.date_range(start_date, maturity_date, freq='D')
    metrics = pd.DataFrame(index=dates)
    metrics['Principal_Balance'] = principal
    metrics['Coupon_Cash_Flow'] = 0.0
    metrics['Principal_Cash_Flow'] = 0.0
    metrics['Total_Coupons_Paid'] = 0.0
    metrics['Total_Coupons_Capitalised'] = 0.0

    # Populate cash flows on their respective dates, ensuring index type is preserved
    maturity_ts = pd.to_datetime(maturity_date)
    for _, row in cash_flows_df.iterrows():
        ts = pd.to_datetime(row['Date'])
        if ts > maturity_ts: continue
        if row['Type'] == 'Coupon':
            metrics.loc[ts, 'Coupon_Cash_Flow'] += row['Cash_Flow']
        elif row['Type'] == 'Principal Repayment':
            metrics.loc[ts, 'Principal_Cash_Flow'] = row['Cash_Flow']

    # --- 3. Generate Daily Metrics from Corrected Cash Flows ---

    # Set principal balance based on bond type
    if interest_payment_type == 'reinvest':
        # For reinvestment bonds, the principal balance increases with each capitalisation event.
        capital_events = cash_flows_df[cash_flows_df['Type'].isin(['Principal Investment', 'Capitalisation'])].copy()
        capital_events['Date'] = pd.to_datetime(capital_events['Date'])
        capital_events_sum = capital_events.groupby('Date')['Cash_Flow'].sum()
        metrics['Principal_Balance'] = capital_events_sum.abs().cumsum().reindex(metrics.index, method='ffill')
        metrics['Total_Coupons_Capitalised'] = metrics['Principal_Balance'] - principal
    else:
        # For other bonds, the principal balance is constant.
        metrics['Principal_Balance'] = principal

    # Total coupons paid is the cumulative sum of coupon cash flows.
    metrics['Total_Coupons_Paid'] = metrics['Coupon_Cash_Flow'].cumsum()

    # Calculate daily accrued interest using the *previous day's* principal to handle compounding correctly.
    metrics['Prev_Day_Principal'] = metrics['Principal_Balance'].shift(1).fillna(principal)
    metrics['Accrued_Interest'] = 0.0
    last_coupon_date = start_date
    coupon_dates_set = set(pd.to_datetime(cash_flows_df[cash_flows_df['Type'].isin(['Coupon', 'Capitalisation'])]['Date']).dt.date)

    for ts, row in metrics.iterrows():
        if ts.date() in coupon_dates_set:
            last_coupon_date = ts.date()
        days_since_last_coupon = (ts.date() - last_coupon_date).days
        prev_day_principal = row['Prev_Day_Principal']
        accrual = prev_day_principal * rate * days_since_last_coupon / 365
        metrics.loc[ts, 'Accrued_Interest'] = accrual

    # Book value is the compounding principal plus the interest accrued since the last capitalisation.
    metrics['Book_Value'] = metrics['Principal_Balance'] + metrics['Accrued_Interest']

    # Final day adjustments
    metrics.loc[pd.to_datetime(maturity_date), 'Principal_Balance'] = 0
    metrics.loc[pd.to_datetime(maturity_date), 'Accrued_Interest'] = 0
    metrics.loc[pd.to_datetime(maturity_date), 'Book_Value'] = 0

    # Reorder and format
    daily_metrics_df = metrics[[
        'Principal_Balance', 'Coupon_Cash_Flow', 'Principal_Cash_Flow', 
        'Accrued_Interest', 'Book_Value', 'Total_Coupons_Paid', 'Total_Coupons_Capitalised'
    ]].reset_index().rename(columns={'index': 'Date'})

    return daily_metrics_df, cash_flows_df

def _calculate_inflation_linked_metrics(start_date: date, term: int, principal: float, unique_reference: str):
    """Placeholder for inflation-linked bond calculations."""
    # TODO: Implement logic for inflation-linked bonds.
    # This would involve fetching inflation data (e.g., CPI) and adjusting cash flows accordingly.
    print(f"Note: Calculation for inflation-linked bond '{unique_reference}' is not yet implemented.")
    return pd.DataFrame(), pd.DataFrame()

def _calculate_top_up_metrics(start_date: date, term: int, principal: float, unique_reference: str):
    """Placeholder for top-up bond calculations."""
    # TODO: Implement logic for top-up bonds.
    # This would involve handling additional principal injections over the life of the bond.
    print(f"Note: Calculation for top-up bond '{unique_reference}' is not yet implemented.")
    return pd.DataFrame(), pd.DataFrame()

def calculate_bond_metrics(
    bond: pd.Series
):
    """Dispatcher function to calculate metrics based on bond type."""
    bond_type = bond.get('bond_type', 'Fixed Rate')
    
    if bond_type == 'Fixed Rate':
        try:
            rate = get_rsa_rsb_rate(bond['start_date'], bond['term'])
            return _calculate_fixed_rate_metrics(
                start_date=bond['start_date'],
                term=bond['term'],
                interest_payment_type=bond['interest_payment_type'],
                principal=bond['principal'],
                unique_reference=bond['unique_reference'],
                rate=rate
            )
        except (ValueError, FileNotFoundError) as e:
            print(f"Could not calculate fixed-rate bond '{bond['unique_reference']}': {e}")
            return pd.DataFrame(), pd.DataFrame()

    elif bond_type == 'Inflation Linked':
        return _calculate_inflation_linked_metrics(
            start_date=bond['start_date'],
            term=bond['term'],
            principal=bond['principal'],
            unique_reference=bond['unique_reference']
        )

    elif bond_type == 'Top-Up':
        return _calculate_top_up_metrics(
            start_date=bond['start_date'],
            term=bond['term'],
            principal=bond['principal'],
            unique_reference=bond['unique_reference']
        )

    else:
        print(f"Warning: Unknown bond type '{bond_type}' for '{bond['unique_reference']}'.")
        return pd.DataFrame(), pd.DataFrame()


def run_and_print_scenario(start_date, term, interest_payment_type, principal, rate):
    """Helper function to run a scenario and print its cash flow schedule."""
    bond_ref = f"BOND_{term}YR_{interest_payment_type.upper()}"
    print("-" * 60)
    print(f"SCENARIO: {term}-Year Bond, {interest_payment_type.replace('_', ' ').title()} Payments")
    print(f"Start Date: {start_date}, Principal: {principal:,.2f}")
    print("-" * 60)

    # Create a mock bond series for the helper function
    mock_bond = pd.Series({
        'start_date': start_date,
        'term': term,
        'interest_payment_type': interest_payment_type,
        'principal': principal,
        'unique_reference': bond_ref,
        'bond_type': 'Fixed Rate' # Assuming fixed rate for the test runner
    })
    _, cash_flows_df = calculate_bond_metrics(mock_bond)

    print("\n----- Cash Flow Schedule -----")
    print(cash_flows_df.to_string())
    print("\n\n")

if __name__ == '__main__':
    # --- Define Base Parameters for Scenarios ---
    base_principal = 1000000.00
    base_start_date = date(2023, 10, 20)
    base_term = 3 # Using a 3-year term for clearer examples

    # --- Run Scenarios for Each Payment Type ---

    # --- Get Rate for Scenarios ---
    try:
        scenario_rate = get_rsa_rsb_rate(base_start_date, base_term)
        print(f"Using rate {scenario_rate:.4f} for all scenarios with start date {base_start_date}\n")

        # --- Run Scenarios for Each Payment Type ---

        # 1. Semi-Annual Payments
        run_and_print_scenario(
            start_date=base_start_date,
            term=base_term,
            interest_payment_type='semi_annual',
            principal=base_principal,
            rate=scenario_rate
        )

        # 2. Monthly Payments
        run_and_print_scenario(
            start_date=base_start_date,
            term=base_term,
            interest_payment_type='monthly',
            principal=base_principal,
            rate=scenario_rate
        )

        # 3. Reinvest (Capitalised Semi-Annual)
        run_and_print_scenario(
            start_date=base_start_date,
            term=base_term,
            interest_payment_type='reinvest',
            principal=base_principal,
            rate=scenario_rate
        )
    except (ValueError, FileNotFoundError) as e:
        print(f"Could not run scenarios: {e}")
