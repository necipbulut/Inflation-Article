
import pandas as pd

file_path = r'C:\Users\necip.bulut\OneDrive - Bahceşehir Üniversitesi\Masaüstü\inflation\roe\data.xlsx'

capital_df = pd.read_excel(file_path, sheet_name='capital')
profit_df = pd.read_excel(file_path, sheet_name='profit')

# Year/Quarter sütunlarının isimlerini çek
time_columns = capital_df.columns[4:48]

sector_mapping = {
    1: 'Brokerage Houses',
    2: 'Fishery Products',
    3: 'Banks',
    4: 'Information Services',
    5: 'IT',
    6: 'Leisure',
    7: 'Office',
    8: 'Other Manufacturing',
    9: 'Electricity, Gas, Steam',
    10: 'Finance',
    11: 'Real Estate Investment Trusts',
    12: 'Real Estate Activities',
    13: 'Food',
    14: 'Venture Capital',
    15: 'Crude Petroleum, Natural Gas',
    16: 'Holding',
    17: 'Legal and Accounting',
    18: 'Construction',
    19: 'Paper and Printing',
    20: 'Rental and Leasing',
    21: 'Chemicals',
    22: 'Accommodation',
    23: 'Coal Mining',
    24: 'Leasing/Factoring',
    25: 'Basic Metals',
    26: 'Metal Products',
    27: 'Metallic Ores',
    28: 'Architecture and Engineering',
    29: 'Furniture',
    30: 'Retail Trade',
    31: 'Pre-Market',
    32: 'Advertising and Marketing',
    33: 'Health',
    34: 'Defense',
    35: 'Travel Agency and Tourism',
    36: 'Insurance',
    37: 'Sports Services',
    38: 'Agriculture and Livestock',
    39: 'Stone and Soil',
    40: 'Textile',
    41: 'Telecommunication',
    42: 'Wholesale Trade',
    43: 'Transportation',
    44: 'Asset Management',
    45: 'Publishing',
    46: 'Food and Beverage'
}

# Sektörel ROE hesaplama fonksiyonu
def calculate_sectoral_roe(capital_df, profit_df, time_columns):
    sectoral_roes = {}

    for col in time_columns:
        # Sermaye ve kar verilerini al ve 0 değerlerini filtrele
        capital_data = capital_df[['firm_id', 'sector_id', col]].dropna()
        capital_data = capital_data[capital_data[col] != 0]
        profit_data = profit_df[['firm_id', col]].dropna()
        profit_data = profit_data[profit_data[col] != 0]

        # Firmaların sektörel sermaye ağırlıklarını hesapla
        sector_capital_sum = capital_data.groupby('sector_id')[col].sum().reset_index()
        capital_data = capital_data.merge(sector_capital_sum, on='sector_id', suffixes=('', '_total'))
        capital_data['weight'] = capital_data[col] / capital_data[col + '_total']

        # Firmaların ROE'sini hesapla
        firm_data = capital_data[['firm_id', 'sector_id', 'weight', col]].merge(profit_data, on='firm_id', suffixes=('_capital', '_profit'))
        
        # Sütun adlarını kontrol edelim
        capital_col = col + '_capital'
        profit_col = col + '_profit'

        firm_data['roe'] = firm_data[profit_col] / firm_data[capital_col]  # profit / capital
        
        # Sektörel ROE'yi hesapla (ağırlıklı ortalama)
        sector_roe = firm_data.groupby('sector_id').apply(lambda x: (x['roe'] * x['weight']).sum()).reset_index()
        sector_roe.columns = ['sector_id', col]
        
        sectoral_roes[col] = sector_roe

    return sectoral_roes

# Sektörel ROE'leri hesapla
sectoral_roes = calculate_sectoral_roe(capital_df, profit_df, time_columns)



# Sonuçları DataFrame'e dönüştür
result_df = sectoral_roes[time_columns[0]].copy()
for col in time_columns[1:]:
    result_df = result_df.merge(sectoral_roes[col], on='sector_id', how='outer')
result_df['sector'] = result_df['sector_id'].map(sector_mapping)

# Sonuçları kaydet
result_df.to_excel(r'C:\Users\necip.bulut\OneDrive - Bahceşehir Üniversitesi\Masaüstü\inflation\roe\sectoral_roe.xlsx', index=False)

