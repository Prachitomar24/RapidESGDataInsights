import pandas as pd
import numpy as np
import requests
import json
from datetime import datetime

class WorldBankDataProcessor:
    """
    A class to handle World Bank data retrieval and processing for ESG analysis
    """
    
    def __init__(self):
        self.base_url = "https://api.worldbank.org/v2"
        self.countries = [
            'USA', 'CHN', 'JPN', 'DEU', 'IND', 'GBR', 'FRA', 'ITA', 'BRA', 'CAN',
            'RUS', 'KOR', 'ESP', 'AUS', 'MEX', 'IDN', 'NLD', 'SAU', 'TUR', 'TWN',
            'BEL', 'IRE', 'ARG', 'ISR', 'THA', 'EGY', 'ZAF', 'PHL', 'SGP', 'NOR'
        ]
        
    def get_indicator_data(self, indicator, start_year=2018, end_year=2022):
        """
        Retrieve indicator data from World Bank API
        """
        all_data = []
        
        # World Bank API has limitations, so we'll fetch data in smaller chunks
        for country in self.countries:
            url = f"{self.base_url}/country/{country}/indicator/{indicator}"
            
            params = {
                'format': 'json',
                'date': f"{start_year}:{end_year}",
                'per_page': 100
            }
            
            try:
                response = requests.get(url, params=params, timeout=10)
                response.raise_for_status()
                data = response.json()
                
                if isinstance(data, list) and len(data) > 1 and data[1]:
                    all_data.extend(data[1])
                    
            except requests.RequestException as e:
                print(f"Warning: Could not fetch {indicator} data for {country}: {e}")
                continue
                
        print(f"Fetched {len(all_data)} records for indicator {indicator}")
        return all_data
    
    def get_co2_emissions_data(self):
        """
        Get CO2 emissions data (metric tons per capita)
        Updated 2024 indicator code
        """
        return self.get_indicator_data('EN.GHG.CO2.PC.CE.AR5')
    
    def get_gdp_per_capita_data(self):
        """
        Get GDP per capita data (current US$)
        """
        return self.get_indicator_data('NY.GDP.PCAP.CD')
    
    def get_gdp_data(self):
        """
        Get GDP data (current US$)
        """
        return self.get_indicator_data('NY.GDP.MKTP.CD')
    
    def process_data_to_dataframe(self, data, indicator_name):
        """
        Convert World Bank API response to pandas DataFrame
        """
        processed_data = []
        
        for item in data:
            if item['value'] is not None:
                processed_data.append({
                    'country': item['country']['value'],
                    'country_code': item['countryiso3code'],
                    'year': int(item['date']),
                    indicator_name: float(item['value'])
                })
        
        df = pd.DataFrame(processed_data)
        print(f"Processed {len(df)} records for {indicator_name}")
        return df
    
    def calculate_co2_gdp_ratio(self, co2_data, gdp_data):
        """
        Calculate CO2 emissions per GDP ratio
        """
        # Merge the datasets
        merged_data = pd.merge(
            co2_data, 
            gdp_data, 
            on=['country', 'country_code', 'year'], 
            how='inner'
        )
        
        # Calculate CO2/GDP ratio (CO2 per capita / GDP per capita * 1000 for better scale)
        merged_data['co2_gdp_ratio'] = (merged_data['co2_per_capita'] / merged_data['gdp_per_capita']) * 1000
        
        return merged_data
    
    def get_latest_year_data(self, df):
        """
        Get the most recent year data for each country
        """
        return df.groupby('country').apply(
            lambda x: x.loc[x['year'].idxmax()]
        ).reset_index(drop=True)
    
    def categorize_countries(self, df, co2_gdp_column='co2_gdp_ratio'):
        """
        Categorize countries as leaders or laggards based on CO2/GDP ratio
        """
        df = df.copy()
        median_ratio = df[co2_gdp_column].median()
        
        df['category'] = df[co2_gdp_column].apply(
            lambda x: 'Leader' if x < median_ratio else 'Laggard'
        )
        
        return df, median_ratio