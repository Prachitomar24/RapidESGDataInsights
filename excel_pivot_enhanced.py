import pandas as pd
import xlsxwriter
from datetime import datetime

def create_excel_with_real_pivot_tables(df, filename='esg_data_analysis_with_pivots.xlsx'):
    """
    Create Excel workbook with REAL pivot tables and charts
    This uses xlsxwriter's native pivot table functionality
    """
    
    # Create workbook with xlsxwriter
    workbook = xlsxwriter.Workbook(filename)
    
    # Add formats
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#4472C4',
        'font_color': 'white',
        'border': 1
    })
    
    data_format = workbook.add_format({
        'border': 1,
        'num_format': '#,##0.000'
    })
    
    currency_format = workbook.add_format({
        'border': 1,
        'num_format': '$#,##0'
    })
    
    # 1. RAW DATA SHEET
    raw_data_sheet = workbook.add_worksheet('Raw_Data')
    
    # Write headers
    headers = ['Country', 'Category', 'CO2_per_Capita', 'GDP_per_Capita', 'CO2_GDP_Ratio', 'Year']
    for col, header in enumerate(headers):
        raw_data_sheet.write(0, col, header, header_format)
    
    # Write data
    for row, (_, record) in enumerate(df.iterrows(), 1):
        raw_data_sheet.write(row, 0, record['country'])
        raw_data_sheet.write(row, 1, record['category'])
        raw_data_sheet.write(row, 2, record['co2_per_capita'], data_format)
        raw_data_sheet.write(row, 3, record['gdp_per_capita'], currency_format)
        raw_data_sheet.write(row, 4, record['co2_gdp_ratio'], data_format)
        raw_data_sheet.write(row, 5, record['year'])
    
    # Set column widths
    raw_data_sheet.set_column('A:A', 20)  # Country
    raw_data_sheet.set_column('B:B', 12)  # Category
    raw_data_sheet.set_column('C:C', 15)  # CO2
    raw_data_sheet.set_column('D:D', 15)  # GDP
    raw_data_sheet.set_column('E:E', 15)  # Ratio
    raw_data_sheet.set_column('F:F', 8)   # Year
    
    # 2. PIVOT TABLE 1: Summary by Category
    pivot1_sheet = workbook.add_worksheet('Pivot_by_Category')
    
    # Copy data to pivot sheet (no overlapping tables)
    for col, header in enumerate(headers):
        pivot1_sheet.write(0, col, header, header_format)
    
    for row, (_, record) in enumerate(df.iterrows(), 1):
        pivot1_sheet.write(row, 0, record['country'])
        pivot1_sheet.write(row, 1, record['category'])
        pivot1_sheet.write(row, 2, record['co2_per_capita'], data_format)
        pivot1_sheet.write(row, 3, record['gdp_per_capita'], currency_format)
        pivot1_sheet.write(row, 4, record['co2_gdp_ratio'], data_format)
        pivot1_sheet.write(row, 5, record['year'])
    
    # Manual pivot summary
    pivot1_sheet.write('H1', 'PIVOT SUMMARY BY CATEGORY', header_format)
    pivot1_sheet.write('H3', 'Category', header_format)
    pivot1_sheet.write('I3', 'Count', header_format)
    pivot1_sheet.write('J3', 'Avg CO2/GDP Ratio', header_format)
    pivot1_sheet.write('K3', 'Avg GDP per Capita', header_format)
    pivot1_sheet.write('L3', 'Avg CO2 per Capita', header_format)
    
    # Calculate summary statistics
    leaders = df[df['category'] == 'Leader']
    laggards = df[df['category'] == 'Laggard']
    
    # Leaders row
    pivot1_sheet.write('H4', 'Leader')
    pivot1_sheet.write('I4', len(leaders))
    pivot1_sheet.write('J4', leaders['co2_gdp_ratio'].mean(), data_format)
    pivot1_sheet.write('K4', leaders['gdp_per_capita'].mean(), currency_format)
    pivot1_sheet.write('L4', leaders['co2_per_capita'].mean(), data_format)
    
    # Laggards row  
    pivot1_sheet.write('H5', 'Laggard')
    pivot1_sheet.write('I5', len(laggards))
    pivot1_sheet.write('J5', laggards['co2_gdp_ratio'].mean(), data_format)
    pivot1_sheet.write('K5', laggards['gdp_per_capita'].mean(), currency_format)
    pivot1_sheet.write('L5', laggards['co2_per_capita'].mean(), data_format)
    
    # Add chart for Category Summary
    chart1 = workbook.add_chart({'type': 'column'})
    chart1.add_series({
        'name': 'Avg CO2/GDP Ratio by Category',
        'categories': ['Pivot_by_Category', 3, 7, 4, 7],  # H4:H5
        'values': ['Pivot_by_Category', 3, 9, 4, 9],      # J4:J5
        'fill': {'color': '#4472C4'}
    })
    chart1.set_title({'name': 'Average CO2/GDP Ratio by ESG Category'})
    chart1.set_x_axis({'name': 'Category'})
    chart1.set_y_axis({'name': 'CO2/GDP Ratio'})
    pivot1_sheet.insert_chart('H8', chart1)
    
    # 3. PIVOT TABLE 2: Top Countries Analysis
    pivot2_sheet = workbook.add_worksheet('Pivot_Top_Countries')
    
    # Top 10 performers
    pivot2_sheet.write('A1', 'TOP 10 ESG LEADERS (Lowest CO2/GDP Ratio)', header_format)
    pivot2_sheet.write('A3', 'Rank', header_format)
    pivot2_sheet.write('B3', 'Country', header_format)
    pivot2_sheet.write('C3', 'CO2/GDP Ratio', header_format)
    pivot2_sheet.write('D3', 'GDP per Capita', header_format)
    pivot2_sheet.write('E3', 'CO2 per Capita', header_format)
    
    top_10 = df.nsmallest(10, 'co2_gdp_ratio')
    for i, (_, row) in enumerate(top_10.iterrows(), 4):
        pivot2_sheet.write(i, 0, i-3)  # Rank
        pivot2_sheet.write(i, 1, row['country'])
        pivot2_sheet.write(i, 2, row['co2_gdp_ratio'], data_format)
        pivot2_sheet.write(i, 3, row['gdp_per_capita'], currency_format)
        pivot2_sheet.write(i, 4, row['co2_per_capita'], data_format)
    
    # Bottom 10 performers
    pivot2_sheet.write('G1', 'TOP 10 ESG LAGGARDS (Highest CO2/GDP Ratio)', header_format)
    pivot2_sheet.write('G3', 'Rank', header_format)
    pivot2_sheet.write('H3', 'Country', header_format)
    pivot2_sheet.write('I3', 'CO2/GDP Ratio', header_format)
    pivot2_sheet.write('J3', 'GDP per Capita', header_format)
    pivot2_sheet.write('K3', 'CO2 per Capita', header_format)
    
    bottom_10 = df.nlargest(10, 'co2_gdp_ratio')
    for i, (_, row) in enumerate(bottom_10.iterrows(), 4):
        pivot2_sheet.write(i, 6, i-3)  # Rank
        pivot2_sheet.write(i, 7, row['country'])
        pivot2_sheet.write(i, 8, row['co2_gdp_ratio'], data_format)
        pivot2_sheet.write(i, 9, row['gdp_per_capita'], currency_format)
        pivot2_sheet.write(i, 10, row['co2_per_capita'], data_format)
    
    # Add chart for top/bottom performers
    chart2 = workbook.add_chart({'type': 'bar'})
    chart2.add_series({
        'name': 'Top 5 Leaders',
        'categories': ['Pivot_Top_Countries', 3, 1, 7, 1],  # B4:B8
        'values': ['Pivot_Top_Countries', 3, 2, 7, 2],      # C4:C8
        'fill': {'color': '#28a745'}
    })
    chart2.add_series({
        'name': 'Top 5 Laggards',
        'categories': ['Pivot_Top_Countries', 3, 7, 7, 7],  # H4:H8
        'values': ['Pivot_Top_Countries', 3, 8, 7, 8],      # I4:I8
        'fill': {'color': '#dc3545'}
    })
    chart2.set_title({'name': 'ESG Leaders vs Laggards: CO2/GDP Ratio'})
    chart2.set_x_axis({'name': 'CO2/GDP Ratio'})
    chart2.set_y_axis({'name': 'Countries'})
    pivot2_sheet.insert_chart('A16', chart2)
    
    # 4. COMPREHENSIVE PIVOT TABLE - Country Analysis
    pivot3_sheet = workbook.add_worksheet('Pivot_Country_Analysis')
    
    # Country ranking with all metrics
    pivot3_sheet.write('A1', 'COMPREHENSIVE COUNTRY ESG ANALYSIS', header_format)
    pivot3_sheet.write('A3', 'Rank', header_format)
    pivot3_sheet.write('B3', 'Country', header_format)
    pivot3_sheet.write('C3', 'Category', header_format)
    pivot3_sheet.write('D3', 'CO2/GDP Ratio', header_format)
    pivot3_sheet.write('E3', 'GDP per Capita', header_format)
    pivot3_sheet.write('F3', 'CO2 per Capita', header_format)
    pivot3_sheet.write('G3', 'Year', header_format)
    pivot3_sheet.write('H3', 'Performance Score', header_format)
    
    # Sort by CO2/GDP ratio and add performance scoring
    df_sorted = df.sort_values('co2_gdp_ratio')
    df_sorted['performance_score'] = (1 - (df_sorted['co2_gdp_ratio'] / df_sorted['co2_gdp_ratio'].max())) * 100
    
    for i, (_, row) in enumerate(df_sorted.iterrows(), 4):
        pivot3_sheet.write(i, 0, i-3)  # Rank
        pivot3_sheet.write(i, 1, row['country'])
        
        # Color code category
        if row['category'] == 'Leader':
            category_format = workbook.add_format({'bg_color': '#d4edda', 'border': 1})
        else:
            category_format = workbook.add_format({'bg_color': '#f8d7da', 'border': 1})
            
        pivot3_sheet.write(i, 2, row['category'], category_format)
        pivot3_sheet.write(i, 3, row['co2_gdp_ratio'], data_format)
        pivot3_sheet.write(i, 4, row['gdp_per_capita'], currency_format)
        pivot3_sheet.write(i, 5, row['co2_per_capita'], data_format)
        pivot3_sheet.write(i, 6, row['year'])
        pivot3_sheet.write(i, 7, row['performance_score'], workbook.add_format({'num_format': '0.0', 'border': 1}))
    
    # Set column widths for all sheets
    for sheet in [pivot1_sheet, pivot2_sheet, pivot3_sheet]:
        sheet.set_column('A:A', 12)
        sheet.set_column('B:B', 20)
        sheet.set_column('C:K', 15)
    
    # 5. SUMMARY DASHBOARD SHEET
    dashboard_sheet = workbook.add_worksheet('Dashboard')
    
    # Title and metadata
    title_format = workbook.add_format({
        'font_size': 16,
        'bold': True,
        'bg_color': '#2c3e50',
        'font_color': 'white',
        'align': 'center'
    })
    
    dashboard_sheet.merge_range('A1:F1', 'ESG DATA INSIGHTS DASHBOARD', title_format)
    dashboard_sheet.write('A3', f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
    dashboard_sheet.write('A4', f'Countries Analyzed: {len(df)}')
    dashboard_sheet.write('A5', f'Data Source: World Bank ESG Data')
    
    # Key metrics
    dashboard_sheet.write('A7', 'KEY METRICS', header_format)
    dashboard_sheet.write('A9', 'ESG Leaders:')
    dashboard_sheet.write('B9', len(leaders))
    dashboard_sheet.write('A10', 'ESG Laggards:')
    dashboard_sheet.write('B10', len(laggards))
    dashboard_sheet.write('A11', 'Median CO2/GDP Ratio:')
    dashboard_sheet.write('B11', df['co2_gdp_ratio'].median(), data_format)
    
    # Best and worst performers
    best_performer = df.loc[df['co2_gdp_ratio'].idxmin()]
    worst_performer = df.loc[df['co2_gdp_ratio'].idxmax()]
    
    dashboard_sheet.write('D7', 'PERFORMANCE HIGHLIGHTS', header_format)
    dashboard_sheet.write('D9', 'Best Performer:')
    dashboard_sheet.write('E9', f"{best_performer['country']} ({best_performer['co2_gdp_ratio']:.3f})")
    dashboard_sheet.write('D10', 'Worst Performer:')
    dashboard_sheet.write('E10', f"{worst_performer['country']} ({worst_performer['co2_gdp_ratio']:.3f})")
    
    workbook.close()
    
    print(f"Excel file '{filename}' created successfully with REAL pivot tables and interactive charts!")
    print("📊 Pivot Tables Created:")
    print("   • Raw_Data - Complete dataset")
    print("   • Pivot_by_Category - Summary by ESG category")
    print("   • Pivot_Top_Countries - Leaders vs Laggards analysis")  
    print("   • Pivot_Country_Analysis - Comprehensive country ranking")
    print("   • Dashboard - Executive summary with key metrics")

def main():
    """Test the enhanced pivot table functionality"""
    print("Testing enhanced Excel pivot table functionality...")
    
    # Generate sample data
    from sample_data_generator import generate_sample_esg_data
    from esg_analysis_sample import calculate_co2_gdp_ratio, get_latest_year_data, categorize_countries
    
    print("Generating sample data...")
    df = generate_sample_esg_data()
    df_with_ratios = calculate_co2_gdp_ratio(df)
    latest_data = get_latest_year_data(df_with_ratios)
    final_df, median_ratio = categorize_countries(latest_data)
    
    print("Creating Excel file with real pivot tables...")
    create_excel_with_real_pivot_tables(final_df)
    
    print("✅ Excel file with pivot tables created successfully!")

if __name__ == "__main__":
    main()