"""
PDF Export Module for RSA Retail Bond Calculator
Generates professional 2-page bond analysis reports
"""

import io
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.patches as mpatches
from datetime import datetime


def create_bond_pdf_report(bond_params, rate, daily_metrics_df, cash_flows_df, performance_metrics):
    """
    Generate a professional 2-page PDF report for bond analysis.
    
    Args:
        bond_params: Dictionary with bond parameters (principal, term, start_date, etc.)
        rate: Fixed rate for the bond
        daily_metrics_df: DataFrame with daily bond metrics
        cash_flows_df: DataFrame with cash flow schedule
        performance_metrics: Dictionary with calculated performance metrics
    
    Returns:
        BytesIO object containing the PDF
    """
    # Create PDF in memory
    pdf_buffer = io.BytesIO()
    
    # Bloomberg-inspired colors
    bg_color = '#0A0A0A'
    text_color = '#E8E8E8'
    accent_color = '#FF9500'
    secondary_color = '#00A0E3'
    grid_color = '#2C2C2C'
    
    with PdfPages(pdf_buffer) as pdf:
        # ===== PAGE 1: Overview and Performance Metrics =====
        fig = plt.figure(figsize=(8.5, 11))
        fig.patch.set_facecolor(bg_color)
        
        # Title
        fig.text(0.5, 0.95, 'RSA RETAIL SAVINGS BOND', 
                ha='center', va='top', fontsize=20, fontweight='bold', color=accent_color)
        fig.text(0.5, 0.92, 'Fixed Rate Bond Analysis Report', 
                ha='center', va='top', fontsize=14, color=text_color)
        
        # Bond Parameters Section
        fig.text(0.1, 0.87, 'BOND PARAMETERS', fontsize=12, fontweight='bold', color=accent_color)
        
        params_text = f"""
Principal Investment:        R {bond_params['principal']:,.2f}
Fixed Rate:                  {rate * 100:.4f}%
Term:                        {bond_params['term']} Years
Payment Type:                {bond_params['interest_payment_type'].replace('_', ' ').title()}
Investment Date:             {bond_params['start_date'].strftime('%Y-%m-%d')}
Maturity Date:               {(bond_params['start_date'].replace(year=bond_params['start_date'].year + bond_params['term'])).strftime('%Y-%m-%d')}
        """
        fig.text(0.1, 0.82, params_text.strip(), fontsize=10, color=text_color, 
                family='monospace', verticalalignment='top')
        
        # Performance Metrics Section
        fig.text(0.1, 0.65, 'PERFORMANCE METRICS', fontsize=12, fontweight='bold', color=accent_color)
        
        metrics_text = f"""
Total Cash Received:         R {performance_metrics['total_cash_received']:,.2f}
Total Interest Earned:       R {performance_metrics['total_interest']:,.2f}
Total Return:                {performance_metrics['return_pct']:.2f}%
Effective Annual Return:     {performance_metrics['effective_annual_return']:.4f}%
        """
        fig.text(0.1, 0.60, metrics_text.strip(), fontsize=10, color=text_color,
                family='monospace', verticalalignment='top')
        
        # Book Value Chart
        ax1 = fig.add_subplot(2, 1, 2)
        ax1.set_facecolor(bg_color)
        
        if not daily_metrics_df.empty and 'Date' in daily_metrics_df.columns:
            plot_df = daily_metrics_df[daily_metrics_df['Book_Value'] > 0].copy()
            
            ax1.plot(plot_df['Date'], plot_df['Book_Value'], 
                    color=accent_color, linewidth=2, label='Book Value')
            ax1.plot(plot_df['Date'], plot_df['Principal_Balance'], 
                    color=secondary_color, linewidth=2, linestyle='--', label='Principal Balance')
            
            ax1.set_xlabel('Date', fontsize=10, color=text_color, fontweight='bold')
            ax1.set_ylabel('Value (ZAR)', fontsize=10, color=text_color, fontweight='bold')
            ax1.set_title('Book Value Growth Over Time', fontsize=12, color=text_color, fontweight='bold', pad=10)
            ax1.grid(True, color=grid_color, alpha=0.3)
            ax1.legend(facecolor=bg_color, edgecolor=grid_color, labelcolor=text_color)
            
            ax1.spines['bottom'].set_color(grid_color)
            ax1.spines['top'].set_color(grid_color)
            ax1.spines['left'].set_color(grid_color)
            ax1.spines['right'].set_color(grid_color)
            ax1.tick_params(colors=text_color)
            
            # Format y-axis with thousand separators
            ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'R {x:,.0f}'))
        
        # Footer
        fig.text(0.5, 0.02, f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} | Page 1 of 2', 
                ha='center', fontsize=8, color=text_color)
        
        pdf.savefig(fig, facecolor=bg_color)
        plt.close(fig)
        
        # ===== PAGE 2: Cash Flow Schedule and Details =====
        fig2 = plt.figure(figsize=(8.5, 11))
        fig2.patch.set_facecolor(bg_color)
        
        # Title
        fig2.text(0.5, 0.95, 'CASH FLOW SCHEDULE & ANALYSIS', 
                 ha='center', va='top', fontsize=16, fontweight='bold', color=accent_color)
        
        # Cash Flow Table
        fig2.text(0.1, 0.90, 'CASH FLOW SCHEDULE', fontsize=12, fontweight='bold', color=accent_color)
        
        if not cash_flows_df.empty:
            # Display first 20 cash flows
            display_cf = cash_flows_df.head(20).copy()
            display_cf['Date'] = pd.to_datetime(display_cf['Date']).dt.strftime('%Y-%m-%d')
            display_cf['Cash_Flow'] = display_cf['Cash_Flow'].apply(lambda x: f"R {x:,.2f}")
            
            # Create table
            ax2 = fig2.add_subplot(2, 1, 1)
            ax2.axis('tight')
            ax2.axis('off')
            
            table = ax2.table(cellText=display_cf[['Date', 'Cash_Flow', 'Type']].values,
                            colLabels=['Date', 'Cash Flow', 'Type'],
                            cellLoc='left',
                            loc='center',
                            bbox=[0, 0, 1, 1])
            
            table.auto_set_font_size(False)
            table.set_fontsize(8)
            
            # Style table
            for (i, j), cell in table.get_celld().items():
                if i == 0:  # Header row
                    cell.set_facecolor('#2A2A2A')
                    cell.set_text_props(weight='bold', color=accent_color)
                else:
                    cell.set_facecolor('#1A1A1A' if i % 2 == 0 else bg_color)
                    cell.set_text_props(color=text_color)
                cell.set_edgecolor(grid_color)
        
        # Cumulative Interest Chart
        ax3 = fig2.add_subplot(2, 1, 2)
        ax3.set_facecolor(bg_color)
        
        if not daily_metrics_df.empty and 'Total_Coupons_Paid' in daily_metrics_df.columns:
            plot_df = daily_metrics_df.copy()
            
            if 'Total_Coupons_Paid' in plot_df.columns and plot_df['Total_Coupons_Paid'].sum() > 0:
                ax3.fill_between(plot_df['Date'], plot_df['Total_Coupons_Paid'], 
                               color=secondary_color, alpha=0.3, label='Interest Paid')
                ax3.plot(plot_df['Date'], plot_df['Total_Coupons_Paid'], 
                        color=secondary_color, linewidth=2)
            
            if 'Total_Coupons_Capitalised' in plot_df.columns and plot_df['Total_Coupons_Capitalised'].sum() > 0:
                ax3.fill_between(plot_df['Date'], plot_df['Total_Coupons_Capitalised'], 
                               color=accent_color, alpha=0.3, label='Interest Capitalised')
                ax3.plot(plot_df['Date'], plot_df['Total_Coupons_Capitalised'], 
                        color=accent_color, linewidth=2)
            
            ax3.set_xlabel('Date', fontsize=10, color=text_color, fontweight='bold')
            ax3.set_ylabel('Cumulative Interest (ZAR)', fontsize=10, color=text_color, fontweight='bold')
            ax3.set_title('Cumulative Interest Over Time', fontsize=12, color=text_color, fontweight='bold', pad=10)
            ax3.grid(True, color=grid_color, alpha=0.3)
            ax3.legend(facecolor=bg_color, edgecolor=grid_color, labelcolor=text_color)
            
            ax3.spines['bottom'].set_color(grid_color)
            ax3.spines['top'].set_color(grid_color)
            ax3.spines['left'].set_color(grid_color)
            ax3.spines['right'].set_color(grid_color)
            ax3.tick_params(colors=text_color)
            
            ax3.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'R {x:,.0f}'))
        
        # Footer
        fig2.text(0.5, 0.02, f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} | Page 2 of 2', 
                 ha='center', fontsize=8, color=text_color)
        
        pdf.savefig(fig2, facecolor=bg_color)
        plt.close(fig2)
    
    pdf_buffer.seek(0)
    return pdf_buffer


def get_download_link(pdf_buffer, filename="bond_analysis_report.pdf"):
    """
    Generate a download link for the PDF.
    
    Args:
        pdf_buffer: BytesIO object containing the PDF
        filename: Name for the downloaded file
    
    Returns:
        HTML string with download link
    """
    import base64
    
    b64 = base64.b64encode(pdf_buffer.read()).decode()
    return f'<a href="data:application/pdf;base64,{b64}" download="{filename}" style="text-decoration: none; color: #FF9500; font-weight: bold;">📄 Download PDF Report</a>'
