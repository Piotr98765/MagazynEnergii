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
df['Okres'] = df['Data'].apply(lambda x: 'Letni' if ((x.month >= 4 and x.month <= 9) or (x.month == 10 and x.day == 1)) else 'Zimowy')


# Dodanie kolumny z informacją o dniu wolnym
df['Czy wolne'] = df['Data'].apply(lambda x: 'Tak' if (x.weekday() == 6 or x.date() in [
    pd.Timestamp(2023, 1, 1), pd.Timestamp(2023, 1, 6), pd.Timestamp(2023, 4, 9), pd.Timestamp(2023, 4, 10),
    pd.Timestamp(2023, 5, 1), pd.Timestamp(2023, 5, 3), pd.Timestamp(2023, 6, 8), pd.Timestamp(2023, 8, 15),
    pd.Timestamp(2023, 11, 1), pd.Timestamp(2023, 11, 11), pd.Timestamp(2023, 12, 25), pd.Timestamp(2023, 12, 26)
]) else 'Nie')


# Dodanie kolumny z informacją o strefie szczytowej/pozaszczytowej
df['Cena za kW [zł]'] = df.apply(lambda x: '2.499' if (
    (x['Data'].month in [1, 2, 11, 12] and (
        (x['Data'].hour >= 8 and x['Data'].hour < 11) or
        (x['Data'].hour >= 16 and x['Data'].hour < 21)
    )) or
    (x['Data'].month in [3, 10] and (
        (x['Data'].hour >= 8 and x['Data'].hour < 11) or
        (x['Data'].hour >= 18 and x['Data'].hour < 21)
    )) or
    (x['Data'].month in [4, 9] and (
        (x['Data'].hour >= 8 and x['Data'].hour < 11) or
        (x['Data'].hour >= 19 and x['Data'].hour < 21)
    )) or
    (x['Data'].month in [5, 6, 7, 8] and (
        (x['Data'].hour >= 8 and x['Data'].hour < 11) or
        (x['Data'].hour >= 20 and x['Data'].hour < 21)
    ))
) else '2.264', axis=1)



df['Stan magazynu'] = df.apply(lambda x: 'Rozładowywanie' if (
    (x['Data'].month in [1, 2, 11, 12] and (
        (x['Data'].hour >= 8 and x['Data'].hour < 11) or
        (x['Data'].hour >= 16 and x['Data'].hour < 21)
    )) or
    (x['Data'].month in [3, 10] and (
        (x['Data'].hour >= 8 and x['Data'].hour < 11) or
        (x['Data'].hour >= 18 and x['Data'].hour < 21)
    )) or
    (x['Data'].month in [4, 9] and (
        (x['Data'].hour >= 8 and x['Data'].hour < 11) or
        (x['Data'].hour >= 19 and x['Data'].hour < 21)
    )) or
    (x['Data'].month in [5, 6, 7, 8] and (
        (x['Data'].hour >= 8 and x['Data'].hour < 11) or
        (x['Data'].hour >= 20 and x['Data'].hour < 21)
    ))
) else 'Ładowanie', axis=1)


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

# Zapisanie danych do pliku Excel
df.to_excel('TABELA1.xlsx', index=False)

print("Dane zostały zapisane do pliku 'zużycie_energii1.xlsx'")

df = pd.read_excel('TABELA1.xlsx')

dict_ladowanie = {}
dict_rozladowanie = {}

for index, row in df.iterrows():
    godzina = row['Godzina']
    zużycie_energii = row['Zużycie energii [kWh]']
    stan_magazynu = row['Stan magazynu']

    if stan_magazynu == 'Ładowanie':
        if godzina in dict_ladowanie:
            dict_ladowanie[godzina].append({'Zużycie energii [kWh]': zużycie_energii})
        else:
            dict_ladowanie[godzina] = [{'Zużycie energii [kWh]': zużycie_energii}]
    elif stan_magazynu == 'Rozładowywanie':
        if godzina in dict_rozladowanie:
            dict_rozladowanie[godzina].append({'Zużycie energii [kWh]': zużycie_energii})
        else:
            dict_rozladowanie[godzina] = [{'Zużycie energii [kWh]': zużycie_energii}]

sequences_ladowanie = []
current_sequence_ladowanie = []
for godzina, dane in dict_ladowanie.items():
    current_sequence_ladowanie.append(godzina)

    if godzina + 1 not in dict_ladowanie:
        sequences_ladowanie.append(current_sequence_ladowanie)
        current_sequence_ladowanie = []

sequences_rozladowanie = []
current_sequence_rozladowanie = []
current_zuzycie_energii = 0
for godzina, dane in dict_rozladowanie.items():
    current_sequence_rozladowanie.append(godzina)
    current_zuzycie_energii += sum([data['Zużycie energii [kWh]'] for data in dane])

    if godzina + 1 not in dict_rozladowanie:
        sequences_rozladowanie.append((current_sequence_rozladowanie, current_zuzycie_energii))
        current_sequence_rozladowanie = []
        current_zuzycie_energii = 0

for i, sequence in enumerate(sequences_ladowanie):
    if len(sequence) > 1:
        start_godzina = sequence[0]
        end_godzina = sequence[-1]
        #print(f"Ciąg {i+1}: {start_godzina}-{end_godzina}")

average_energy_consumption = {}
liczba_pracujacych_magazynow = {}
klimatyzacja = {}

for i, (sequence, zuzycie_energii) in enumerate(sequences_rozladowanie):
    if len(sequence) >= 1:
        ciag_ladowania = sequences_ladowanie[i]
        liczba_godzin_ladowania = len(ciag_ladowania)
        liczba_godzin_klimatyzacji = len(ciag_ladowania) + len(sequence) - 1
        if zuzycie_energii > 6210:
            srednie_zuzycie_energii= 6210 / liczba_godzin_ladowania
            odciazenie_ladowowania = 6210 / len(sequence) 
            for godzina in ciag_ladowania:
                average_energy_consumption[godzina] = srednie_zuzycie_energii
                liczba_pracujacych_magazynow[godzina] = 3
                klimatyzacja[godzina] = 390 / liczba_godzin_klimatyzacji
            for godzina in sequence:
                average_energy_consumption[godzina] = -odciazenie_ladowowania
                liczba_pracujacych_magazynow[godzina] = 3
                klimatyzacja[godzina] = 390 / liczba_godzin_klimatyzacji

        else:
            srednie_zuzycie_energii = zuzycie_energii / liczba_godzin_ladowania
            for godzina in ciag_ladowania:
                average_energy_consumption[godzina] = srednie_zuzycie_energii
                if zuzycie_energii <= 4140 and zuzycie_energii >= 2070:
                    liczba_pracujacych_magazynow[godzina] = 2
                    klimatyzacja[godzina] = 260 / liczba_godzin_klimatyzacji
                else:
                    liczba_pracujacych_magazynow[godzina] = 1
                    klimatyzacja[godzina] = 130 / liczba_godzin_klimatyzacji
            for godzina in sequence:
                value = dict_rozladowanie[godzina][0]['Zużycie energii [kWh]']
                negative_value = -abs(value)
                average_energy_consumption[godzina] = negative_value
                if zuzycie_energii <= 4140 and zuzycie_energii >= 2070:
                    liczba_pracujacych_magazynow[godzina] = 2
                    klimatyzacja[godzina] = 260 / liczba_godzin_klimatyzacji
                else:
                    liczba_pracujacych_magazynow[godzina] = 1
                    klimatyzacja[godzina] = 130 / liczba_godzin_klimatyzacji


        start_godzina = sequence[0]
        end_godzina = sequence[-1]
        
        
last_ladowanie_sequence = sequences_ladowanie[-1]
if last_ladowanie_sequence[-1] == df['Godzina'].iloc[-1]:
    for godzina in last_ladowanie_sequence:
        average_energy_consumption[godzina] = 0

df['Zużycie energii przez magazyn [kWh]'] = df['Godzina'].map(average_energy_consumption)
df['Ile magazynów pracuje?'] = df['Godzina'].map(liczba_pracujacych_magazynow)
df['Klimatyzacja [kWh]'] = df['Godzina'].map(klimatyzacja)

df['Bilans [kWh]'] = df['Zużycie energii [kWh]'] + df['Zużycie energii przez magazyn [kWh]'] + df['Klimatyzacja [kWh]']


def color_rows(row):
    if row['Stan magazynu'] == 'Ładowanie':
        return ['background-color: #97CFA4'] * len(row)
    elif row['Stan magazynu'] == 'Rozładowywanie':
        return ['background-color: orange'] * len(row)
    else:
        return [''] * len(row)

# Pokoloruj wiersze za pomocą Styler
df_styled = df.style.apply(color_rows, axis=1)
df_styled.to_excel('TABELA2.xlsx', index=False, engine='openpyxl')