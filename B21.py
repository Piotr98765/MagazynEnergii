import itertools
import pandas as pd

# Constants
HOURS_IN_YEAR = 8760
HOLIDAY_MULTIPLIER = 0.1
PRICE_PER_KW = 2.333

############################Data###################################
# Sample energy consumption list for 24 hours in the summer period
summer_energy_consumption_24h = [762.34, 760.68, 748.73, 750.72, 866.58, 1044.92, 1120.25, 1201.87, 1390.13, 1472.8, 1514.94, 1550.73, 1590.24, 1571.61, 1600, 1475.29, 1404.06, 1358.35, 1360.91, 1286.4, 1165.12, 1033.45, 914.99, 784.74]
# Sample energy consumption list for 24 hours in the winter period
winter_energy_consumption_24h = [1090.51, 1097.77, 1079.91, 1067.44, 1219.34, 1401.5, 1478.97, 1512.34, 1600, 1575.48, 1545.97, 1535.45, 1526.82, 1512.34, 1581.99, 1587.36, 1569.66, 1572.37, 1551.89, 1518.39, 1434.77, 1278.16, 1147.82, 1019.5]
###################################################################

# Function to extend energy consumption list
def extend_energy_consumption(original_list, hours):
    return list(itertools.islice(itertools.cycle(original_list), hours))

def holiday_energy_consumption(origina_list):
    return [x * 0.1 for x in origina_list]

# Function to create DataFrame
def create_energy_consumption_dataframe():
    # Creating a DataFrame with the extended data list
    df = pd.DataFrame(columns=['Hour', 'Date', 'Energy Consumption [kWh]', 'Period', 'Is Holiday'])

    hours_range = range(1, 8761)
    df['Hour'] = hours_range

    # Adding a column with the hour number
    df['Date'] = pd.date_range('2023-01-01 00:00:00', periods=8760, freq='H')

    # Adding a column with information about the summer/winter period
    df['Period'] = df['Date'].apply(lambda x: 'Summer' if (x.month >= 4 and x.month <= 9) else 'Winter')

    # Adding a column with information about holidays
    df['Is Holiday'] = df['Date'].apply(lambda x: 'Yes' if (x.weekday() == 6 or x.date() in [
        pd.Timestamp(2023, 1, 1), pd.Timestamp(2023, 1, 6), pd.Timestamp(2023, 4, 9), pd.Timestamp(2023, 4, 10),
        pd.Timestamp(2023, 5, 1), pd.Timestamp(2023, 5, 3), pd.Timestamp(2023, 5, 28), pd.Timestamp(2023, 6, 8), pd.Timestamp(2023, 8, 15),
        pd.Timestamp(2023, 11, 1), pd.Timestamp(2023, 11, 11), pd.Timestamp(2023, 11, 13), pd.Timestamp(2023, 12, 25), pd.Timestamp(2023, 12, 26)
    ]) else 'No')

    # Adding a column with information about peak/off-peak zone
    df['Price per kW [PLN]'] = 2.333

    # Initializing the 'Energy Consumption' column with empty values
    df['Energy Consumption [kWh]'] = ''
    
    return df

def main():
    summer_energy_consumption_8760h = extend_energy_consumption(summer_energy_consumption_24h, HOURS_IN_YEAR)
    winter_energy_consumption_8760h = extend_energy_consumption(winter_energy_consumption_24h, HOURS_IN_YEAR)

    holiday_summer_energy_consumption_24h = holiday_energy_consumption(summer_energy_consumption_24h)
    holiday_winter_energy_consumption_24h = holiday_energy_consumption(winter_energy_consumption_24h)

    holiday_summer_energy_consumption_8760h = extend_energy_consumption(holiday_summer_energy_consumption_24h, HOURS_IN_YEAR)
    holiday_winter_energy_consumption_8760h = extend_energy_consumption(holiday_winter_energy_consumption_24h, HOURS_IN_YEAR)

    df = create_energy_consumption_dataframe()

    # Replacing energy consumption values for summer and holidays
    df.loc[(df['Period'] == 'Summer') & (df['Is Holiday'] == 'Yes'), 'Energy Consumption [kWh]'] = pd.Series(holiday_summer_energy_consumption_8760h)

    # Replacing energy consumption values for summer and workdays
    df.loc[(df['Period'] == 'Summer') & (df['Is Holiday'] == 'No'), 'Energy Consumption [kWh]'] = pd.Series(summer_energy_consumption_8760h)

    # Replacing energy consumption values for winter and holidays
    df.loc[(df['Period'] == 'Winter') & (df['Is Holiday'] == 'Yes'), 'Energy Consumption [kWh]'] = pd.Series(holiday_winter_energy_consumption_8760h)
    df.loc[(df['Period'] == 'Winter') & (df['Is Holiday'] == 'No'), 'Energy Consumption [kWh]'] = pd.Series(winter_energy_consumption_8760h)

    # Replacing energy consumption values for winter and workdays

    df['Cost [PLN]'] = df['Energy Consumption [kWh]'] * df['Price per kW [PLN]']
    # Saving the data to an Excel file
    df.to_excel('B21.xlsx', index=False)

    print("Data has been saved to the 'B21.xlsx' file")

if __name__ == '__main__':
    main()