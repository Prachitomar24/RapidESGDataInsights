import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from data_processor import WorldBankDataProcessor
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import xlsxwriter

def create_excel_with_pivot_charts(df, filename='esg_data_analysis.xlsx'):
    """
    Create Excel workbook with data and pivot charts
    """
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Write main data
        df.to_excel(writer, sheet_name='Raw_Data', index=False)
        
        # Create summary data for pivot chart
        summary_df = df.groupby(['country', 'category']).agg({
            'co2_per_capita': 'mean',
            'gdp_per_capita': 'mean',
            'co2_gdp_ratio': 'mean'
        }).reset_index()
        
        summary_df.to_excel(writer, sheet_name='Summary_Data', index=False)
        
        # Create charts worksheet
        chart_worksheet = workbook.add_worksheet('Charts')
        
        # Chart 1: CO2 vs GDP scatter plot
        chart1 = workbook.add_chart({'type': 'scatter'})
        chart1.add_series({
            'name': 'Countries',
            'categories': ['Raw_Data', 1, 1, len(df), 1],  # Country names
            'values': ['Raw_Data', 1, 4, len(df), 4],      # GDP per capita
            'y2_values': ['Raw_Data', 1, 3, len(df), 3],   # CO2 per capita
        })
        chart1.set_title({'name': 'CO2 Emissions vs GDP per Capita'})
        chart1.set_x_axis({'name': 'GDP per Capita (USD)'})
        chart1.set_y_axis({'name': 'CO2 per Capita (metric tons)'})
        chart_worksheet.insert_chart('A2', chart1)
        
        # Chart 2: CO2/GDP Ratio by Category
        leaders_avg = df[df['category'] == 'Leader']['co2_gdp_ratio'].mean()
        laggards_avg = df[df['category'] == 'Laggard']['co2_gdp_ratio'].mean()
        
        chart2 = workbook.add_chart({'type': 'column'})
        chart2.add_series({
            'name': 'Average CO2/GDP Ratio',
            'categories': ['Leaders', 'Laggards'],
            'values': [leaders_avg, laggards_avg],
        })
        chart2.set_title({'name': 'Average CO2/GDP Ratio: Leaders vs Laggards'})
        chart2.set_x_axis({'name': 'Category'})
        chart2.set_y_axis({'name': 'CO2/GDP Ratio'})
        chart_worksheet.insert_chart('A18', chart2)
        
        print(f"Excel file '{filename}' created successfully with pivot charts!")

def create_visualizations(df):
    """
    Create and save various visualizations
    """
    # Set style
    plt.style.use('seaborn-v0_8')
    sns.set_palette("husl")
    
    # 1. CO2 vs GDP Scatter Plot
    plt.figure(figsize=(12, 8))
    colors = {'Leader': 'green', 'Laggard': 'red'}
    for category in df['category'].unique():
        data = df[df['category'] == category]
        plt.scatter(data['gdp_per_capita'], data['co2_per_capita'], 
                   c=colors[category], label=category, alpha=0.7, s=100)
    
    plt.xlabel('GDP per Capita (USD)', fontsize=12)
    plt.ylabel('CO2 per Capita (metric tons)', fontsize=12)
    plt.title('CO2 Emissions vs GDP per Capita by Country Category', fontsize=14, fontweight='bold')
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig('visualizations/co2_vs_gdp_scatter.png', dpi=300, bbox_inches='tight')
    plt.close()
    
    # 2. Top 10 and Bottom 10 CO2/GDP Ratios
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
    
    # Top 10 (worst performers)
    top_10 = df.nlargest(10, 'co2_gdp_ratio')
    ax1.barh(range(len(top_10)), top_10['co2_gdp_ratio'], color='red', alpha=0.7)
    ax1.set_yticks(range(len(top_10)))
    ax1.set_yticklabels(top_10['country'])
    ax1.set_xlabel('CO2/GDP Ratio')
    ax1.set_title('Top 10 Highest CO2/GDP Ratios (Laggards)', fontweight='bold')
    ax1.grid(True, alpha=0.3, axis='x')
    
    # Bottom 10 (best performers)
    bottom_10 = df.nsmallest(10, 'co2_gdp_ratio')
    ax2.barh(range(len(bottom_10)), bottom_10['co2_gdp_ratio'], color='green', alpha=0.7)
    ax2.set_yticks(range(len(bottom_10)))
    ax2.set_yticklabels(bottom_10['country'])
    ax2.set_xlabel('CO2/GDP Ratio')
    ax2.set_title('Top 10 Lowest CO2/GDP Ratios (Leaders)', fontweight='bold')
    ax2.grid(True, alpha=0.3, axis='x')
    
    plt.tight_layout()
    plt.savefig('visualizations/top_bottom_performers.png', dpi=300, bbox_inches='tight')
    plt.close()
    
    # 3. Category Distribution
    plt.figure(figsize=(8, 6))
    category_counts = df['category'].value_counts()
    colors = ['green', 'red']
    plt.pie(category_counts.values, labels=category_counts.index, autopct='%1.1f%%', 
            colors=colors, startangle=90)
    plt.title('Distribution of ESG Leaders vs Laggards', fontsize=14, fontweight='bold')
    plt.axis('equal')
    plt.tight_layout()
    plt.savefig('visualizations/category_distribution.png', dpi=300, bbox_inches='tight')
    plt.close()
    
    print("Visualizations saved to 'visualizations/' folder!")

def generate_brief(df, median_ratio):
    """
    Generate one-page executive brief
    """
    leaders = df[df['category'] == 'Leader']
    laggards = df[df['category'] == 'Laggard']
    
    brief = f"""
ESG DATA INSIGHTS: CO2/GDP ANALYSIS EXECUTIVE BRIEF
{'='*60}

ANALYSIS OVERVIEW
• Dataset: {len(df)} countries analyzed using World Bank ESG data
• Metric: CO2 emissions per GDP ratio (CO2 per capita / GDP per capita * 1000)
• Median threshold: {median_ratio:.2f}
• Analysis date: {datetime.now().strftime('%B %Y')}

KEY FINDINGS
• ESG Leaders: {len(leaders)} countries ({len(leaders)/len(df)*100:.1f}%)
• ESG Laggards: {len(laggards)} countries ({len(laggards)/len(df)*100:.1f}%)

TOP 5 ESG LEADERS (Lowest CO2/GDP Ratio):
{'-'*40}"""
    
    top_leaders = leaders.nsmallest(5, 'co2_gdp_ratio')
    for i, row in top_leaders.iterrows():
        brief += f"\n{row['country']:15} | Ratio: {row['co2_gdp_ratio']:.2f}"
    
    brief += f"\n\nTOP 5 ESG LAGGARDS (Highest CO2/GDP Ratio):\n{'-'*40}"
    
    top_laggards = laggards.nlargest(5, 'co2_gdp_ratio')
    for i, row in top_laggards.iterrows():
        brief += f"\n{row['country']:15} | Ratio: {row['co2_gdp_ratio']:.2f}"
    
    brief += f"""

STATISTICAL SUMMARY
• Average CO2/GDP ratio (Leaders): {leaders['co2_gdp_ratio'].mean():.2f}
• Average CO2/GDP ratio (Laggards): {laggards['co2_gdp_ratio'].mean():.2f}
• Largest gap between categories: {laggards['co2_gdp_ratio'].mean() - leaders['co2_gdp_ratio'].mean():.2f}

RECOMMENDATIONS
1. Leaders should maintain current sustainable practices and share best practices
2. Laggards should implement carbon reduction strategies and green technology adoption
3. Focus on energy efficiency improvements and renewable energy transition
4. Regular monitoring and benchmarking against leader countries

METHODOLOGY
Data sourced from World Bank Open Data API covering CO2 emissions per capita
and GDP per capita. Countries categorized based on median CO2/GDP ratio.
Analysis automated using Python for reproducible insights.

Generated by: Rapid ESG Data Insights Tool
Contact: ESG Analytics Team
"""
    
    with open('esg_brief.txt', 'w') as f:
        f.write(brief)
    
    print("Executive brief saved as 'esg_brief.txt'!")
    return brief

def main():
    """
    Main analysis function
    """
    print("🌍 Starting ESG Data Analysis...")
    
    # Initialize data processor
    processor = WorldBankDataProcessor()
    
    # Get data
    print("📊 Fetching CO2 emissions data...")
    co2_raw = processor.get_co2_emissions_data()
    if not co2_raw:
        print("❌ No CO2 data received")
        return
    co2_df = processor.process_data_to_dataframe(co2_raw, 'co2_per_capita')
    
    print("📊 Fetching GDP per capita data...")
    gdp_raw = processor.get_gdp_per_capita_data()
    if not gdp_raw:
        print("❌ No GDP data received")
        return
    gdp_df = processor.process_data_to_dataframe(gdp_raw, 'gdp_per_capita')
    
    print(f"CO2 data shape: {co2_df.shape}")
    print(f"GDP data shape: {gdp_df.shape}")
    print("CO2 columns:", co2_df.columns.tolist())
    print("GDP columns:", gdp_df.columns.tolist())
    
    # Combine and process
    print("🔄 Processing and combining datasets...")
    combined_df = processor.calculate_co2_gdp_ratio(co2_df, gdp_df)
    latest_data = processor.get_latest_year_data(combined_df)
    final_df, median_ratio = processor.categorize_countries(latest_data)
    
    print(f"✅ Analysis complete! Processed {len(final_df)} countries.")
    print(f"📈 Median CO2/GDP ratio: {median_ratio:.2f}")
    
    # Create outputs
    print("📊 Creating Excel workbook with pivot charts...")
    create_excel_with_pivot_charts(final_df)
    
    print("📈 Generating visualizations...")
    create_visualizations(final_df)
    
    print("📝 Creating executive brief...")
    generate_brief(final_df, median_ratio)
    
    print("\n🎉 Analysis complete! Check the following outputs:")
    print("   • esg_data_analysis.xlsx - Excel workbook with data and charts")
    print("   • visualizations/ - Folder with PNG charts")
    print("   • esg_brief.txt - Executive summary")

if __name__ == "__main__":
    main()