import pandas as pd
import numpy as np
import random

def generate_sample_esg_data():
    """
    Generate sample ESG data for demonstration purposes
    Based on realistic CO2 and GDP data patterns
    """
    countries = [
        'United States', 'China', 'Japan', 'Germany', 'India', 'United Kingdom', 
        'France', 'Italy', 'Brazil', 'Canada', 'Russia', 'South Korea', 
        'Spain', 'Australia', 'Mexico', 'Indonesia', 'Netherlands', 'Saudi Arabia', 
        'Turkey', 'Taiwan', 'Belgium', 'Ireland', 'Argentina', 'Israel', 
        'Thailand', 'Egypt', 'South Africa', 'Philippines', 'Singapore', 'Norway'
    ]
    
    country_codes = [
        'USA', 'CHN', 'JPN', 'DEU', 'IND', 'GBR', 'FRA', 'ITA', 'BRA', 'CAN',
        'RUS', 'KOR', 'ESP', 'AUS', 'MEX', 'IDN', 'NLD', 'SAU', 'TUR', 'TWN',
        'BEL', 'IRE', 'ARG', 'ISR', 'THA', 'EGY', 'ZAF', 'PHL', 'SGP', 'NOR'
    ]
    
    # Set seed for reproducible results
    np.random.seed(42)
    random.seed(42)
    
    data = []
    
    for i, (country, code) in enumerate(zip(countries, country_codes)):
        # Generate realistic CO2 per capita (metric tons)
        # Developed countries: 8-20, Developing: 2-12, Oil producers: 15-25
        if code in ['USA', 'AUS', 'CAN', 'SAU']:  # High CO2 emitters
            co2_base = np.random.uniform(12, 20)
        elif code in ['CHN', 'IND', 'BRA', 'MEX', 'IDN', 'THA', 'EGY', 'ZAF', 'PHL']:  # Developing
            co2_base = np.random.uniform(2, 8)
        else:  # Developed European and others
            co2_base = np.random.uniform(6, 12)
            
        # Generate realistic GDP per capita (current USD)
        if code in ['USA', 'NOR', 'SGP']:  # Very high income
            gdp_base = np.random.uniform(60000, 80000)
        elif code in ['DEU', 'GBR', 'FRA', 'JPN', 'AUS', 'CAN', 'BEL', 'NLD', 'IRE']:  # High income
            gdp_base = np.random.uniform(35000, 55000)
        elif code in ['KOR', 'ESP', 'ITA', 'ISR']:  # Upper middle income developed
            gdp_base = np.random.uniform(25000, 40000)
        elif code in ['SAU', 'RUS', 'ARG', 'TUR']:  # Oil/resource rich or upper middle
            gdp_base = np.random.uniform(15000, 30000)
        else:  # Developing economies
            gdp_base = np.random.uniform(3000, 15000)
        
        # Add some variation for recent years (2018-2022)
        for year in [2018, 2019, 2020, 2021, 2022]:
            year_variation = np.random.uniform(0.95, 1.05)
            co2_value = co2_base * year_variation
            gdp_value = gdp_base * year_variation
            
            # COVID impact for 2020
            if year == 2020:
                co2_value *= np.random.uniform(0.85, 0.95)  # CO2 dropped
                gdp_value *= np.random.uniform(0.90, 0.98)  # GDP dropped
            
            data.append({
                'country': country,
                'country_code': code,
                'year': year,
                'co2_per_capita': round(co2_value, 2),
                'gdp_per_capita': round(gdp_value, 0)
            })
    
    df = pd.DataFrame(data)
    return df

def save_sample_data():
    """
    Generate and save sample data
    """
    df = generate_sample_esg_data()
    df.to_csv('sample_esg_data.csv', index=False)
    print(f"Generated {len(df)} records for {len(df['country'].unique())} countries")
    print("Sample data saved to 'sample_esg_data.csv'")
    return df

if __name__ == "__main__":
    save_sample_data()