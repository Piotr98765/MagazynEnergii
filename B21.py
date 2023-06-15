import pandas as pd
import itertools

# Przykładowa lista zużycia energii dla 24 godzin w okresie letnim
zużycie_energii_24h_lato = [762.34, 760.68, 748.73, 750.72, 866.58, 1044.92, 1120.25, 1201.87, 1390.13, 1472.8, 1514.94, 1550.73, 1590.24, 1571.61, 1600, 1475.29, 1404.06, 1358.35, 1360.91, 1286.4, 1165.12, 1033.45, 914.99, 784.74]

# Przykładowa lista zużycia energii dla 24 godzin w okresie zimowym
zużycie_energii_24h_zima = [1090.51, 1097.77, 1079.91, 1067.44, 1219.34, 1401.5, 1478.97, 1512.34, 1600, 1575.48, 1545.97, 1535.45, 1526.82, 1512.34, 1581.99, 1587.36, 1569.66, 1572.37, 1551.89, 1518.39, 1434.77, 1278.16, 1147.82, 1019.5]

# Rozszerzenie listy zużycia energii dla 24 godzin w okresie letnim do 8760 godzin
zużycie_energii_8760h_lato = list(itertools.islice(itertools.cycle(zużycie_energii_24h_lato), 8760))

# Rozszerzenie listy zużycia energii dla 24 godzin w okresie zimowym do 8760 godzin
zużycie_energii_8760h_zima = list(itertools.islice(itertools.cycle(zużycie_energii_24h_zima), 8760))

# Przykładowa lista zużycia energii dla 24 godzin w okresie letnim w dni wolne
zużycie_energii_24h_lato_dzien_wolny = [x * 0.1 for x in zużycie_energii_24h_lato]
zużycie_energii_24h_zima_dzien_wolny = [x * 0.1 for x in zużycie_energii_24h_zima]

# Rozszerzenie listy zużycia energii dla 24 godzin w okresie letnim w dni wolne do 8760 godzin
zużycie_energii_8760h_lato_dzien_wolny = list(itertools.islice(itertools.cycle(zużycie_energii_24h_lato_dzien_wolny), 8760))
zużycie_energii_8760h_zima_dzien_wolny = list(itertools.islice(itertools.cycle(zużycie_energii_24h_zima_dzien_wolny), 8760))

# Rozszerzenie listy zużycia energii dla 24 godzin w okresie zimowym do 8760 godzin
zużycie_energii_8760h_zima = []
for i in range(8760):
    zużycie_energii_8760h_zima.extend(zużycie_energii_24h_zima)

# Rozszerzenie listy zużycia energii dla 24 godzin w okresie letnim w dniu wolnym do 8760 godzin
zużycie_energii_8760h_lato_dzien_wolny = list(itertools.islice(itertools.cycle(zużycie_energii_24h_lato_dzien_wolny), 8760))

# Rozszerzenie listy zużycia energii dla 24 godzin w okresie zimowym w dniu wolnym do 8760 godzin
zużycie_energii_8760h_zima_dzien_wolny = []
for i in range(8760):
    zużycie_energii_8760h_zima_dzien_wolny.extend(zużycie_energii_24h_zima_dzien_wolny)

# Utworzenie DataFrame z rozszerzoną listą danych
df = pd.DataFrame(columns=['Godzina', 'Data', 'Zużycie energii [kWh]', 'Okres', 'Czy wolne'])

hours_range = range(1, 8761)
df['Godzina'] = hours_range

# Dodanie kolumny z numerem godziny
df['Data'] = pd.date_range('2023-01-01 00:00:00', periods=8760, freq='H')

# Dodanie kolumny z informacją o okresie letnim/zimowym
df['Okres'] = df['Data'].apply(lambda x: 'Letni' if (x.month >= 4 and x.month <= 9) else 'Zimowy')

# Dodanie kolumny z informacją o dniu wolnym
df['Czy wolne'] = df['Data'].apply(lambda x: 'Tak' if (x.weekday() == 6 or x.date() in [
    pd.Timestamp(2023, 1, 1), pd.Timestamp(2023, 1, 6), pd.Timestamp(2023, 4, 9), pd.Timestamp(2023, 4, 10),
    pd.Timestamp(2023, 5, 1), pd.Timestamp(2023, 5, 3), pd.Timestamp(2023, 5, 28), pd.Timestamp(2023, 6, 8), pd.Timestamp(2023, 8, 15),
    pd.Timestamp(2023, 11, 1), pd.Timestamp(2023, 11, 11), pd.Timestamp(2023, 11, 13), pd.Timestamp(2023, 12, 25), pd.Timestamp(2023, 12, 26)
]) else 'Nie')

# Dodanie kolumny z informacją o strefie szczytowej/pozaszczytowej
df['Cena za kW [zł]'] = 2.333

# Inicjalizacja kolumny 'Zużycie energii' pustymi wartościami
df['Zużycie energii [kWh]'] = ''

# Zastąpienie wartości zużycia energii dla okresu letniego i dni wolnych
df.loc[(df['Okres'] == 'Letni') & (df['Czy wolne'] == 'Tak'), 'Zużycie energii [kWh]'] = pd.Series(zużycie_energii_8760h_lato_dzien_wolny)

# Zastąpienie wartości zużycia energii dla okresu letniego i dni roboczych
df.loc[(df['Okres'] == 'Letni') & (df['Czy wolne'] == 'Nie'), 'Zużycie energii [kWh]'] = pd.Series(zużycie_energii_8760h_lato)

# Zastąpienie wartości zużycia energii dla okresu zimowego i dni wolnych
df.loc[(df['Okres'] == 'Zimowy') & (df['Czy wolne'] == 'Tak'), 'Zużycie energii [kWh]'] = pd.Series(zużycie_energii_8760h_zima_dzien_wolny)
df.loc[(df['Okres'] == 'Zimowy') & (df['Czy wolne'] == 'Nie'), 'Zużycie energii [kWh]'] = pd.Series(zużycie_energii_8760h_zima)
# Zastąpienie wartości zużycia energii dla okresu zimowego i dni roboczych

df['Koszt [zł]'] = df['Zużycie energii [kWh]'] * df['Cena za kW [zł]']
# Zapisanie danych do pliku Excel
df.to_excel('B21.xlsx', index=False)

print("Dane zostały zapisane do pliku 'B21.xlsx'")