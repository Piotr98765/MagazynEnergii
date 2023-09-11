import pandas as pd
import B22

def load_data(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except FileNotFoundError:
        raise Exception("File not found: " + file_path)

def process_data(df):
    charging_data = {}
    discharging_data = {}
    
    for index, row in df.iterrows():
        hour = row['Hour']
        energy_consumption = row['Energy Consumption [kWh]']
        storage_status = row['Status']

        if storage_status == 'CHARGING':
            if hour in charging_data:
                charging_data[hour].append({'Energy Consumption [kWh]': energy_consumption})
            else:
                charging_data[hour] = [{'Energy Consumption [kWh]': energy_consumption}]
        elif storage_status == 'DISCHARGING':
            if hour in discharging_data:
                discharging_data[hour].append({'Energy Consumption [kWh]': energy_consumption})
            else:
                discharging_data[hour] = [{'Energy Consumption [kWh]': energy_consumption}]

    return charging_data, discharging_data

def load_unload_sequences(charging_data, discharging_data):
    loading_sequences = []
    current_loading_sequence = []
    for hour, data in charging_data.items():
        current_loading_sequence.append(hour)

        if hour + 1 not in charging_data:
            loading_sequences.append(current_loading_sequence)
            current_loading_sequence = []

    unloading_sequences = []
    current_unloading_sequence = []
    current_energy_consumption = 0
    for hour, data in discharging_data.items():
        current_unloading_sequence.append(hour)
        current_energy_consumption += sum([data['Energy Consumption [kWh]'] for data in data])

        if hour + 1 not in discharging_data:
            unloading_sequences.append((current_unloading_sequence, current_energy_consumption))
            current_unloading_sequence = []
            current_energy_consumption = 0
    
    return loading_sequences, unloading_sequences

def calculate_average_energy(loading_sequences, unloading_sequences, discharging_data, df):
    average_energy_consumption = {}
    working_storage_count = {}
    air_conditioning = {}

    for i, (sequence, energy_consumption) in enumerate(unloading_sequences):
        if len(sequence) >= 1:
            loading_sequence = loading_sequences[i]
            loading_hours_count = len(loading_sequence)
            air_conditioning_hours_count = len(loading_sequence) + len(sequence) - 1
            if energy_consumption > 6210:
                average_energy_per_hour = 6210 / loading_hours_count
                unloading_relief = 6210 / len(sequence)
                for hour in loading_sequence:
                    average_energy_consumption[hour] = average_energy_per_hour
                    working_storage_count[hour] = 3
                    air_conditioning[hour] = 390 / air_conditioning_hours_count
                for hour in sequence:
                    average_energy_consumption[hour] = -unloading_relief
                    working_storage_count[hour] = 3
                    air_conditioning[hour] = 390 / air_conditioning_hours_count

            else:
                average_energy_per_hour = energy_consumption / loading_hours_count
                for hour in loading_sequence:
                    average_energy_consumption[hour] = average_energy_per_hour
                    if energy_consumption <= 4140 and energy_consumption >= 2070:
                        working_storage_count[hour] = 2
                        air_conditioning[hour] = 260 / air_conditioning_hours_count
                    else:
                        working_storage_count[hour] = 1
                        air_conditioning[hour] = 130 / air_conditioning_hours_count
                for hour in sequence:
                    value = discharging_data[hour][0]['Energy Consumption [kWh]']
                    negative_value = -abs(value)
                    average_energy_consumption[hour] = negative_value
                    if energy_consumption <= 4140 and energy_consumption >= 2070:
                        working_storage_count[hour] = 2
                        air_conditioning[hour] = 260 / air_conditioning_hours_count
                    else:
                        working_storage_count[hour] = 1
                        air_conditioning[hour] = 130 / air_conditioning_hours_count

    last_loading_sequence = loading_sequences[-1]
    if last_loading_sequence[-1] == df['Hour'].iloc[-1]:
        for hour in last_loading_sequence:
            average_energy_consumption[hour] = 0

    return average_energy_consumption, working_storage_count, air_conditioning


def color_rows(row):
    if row['Status'] == 'CHARGING':
        return ['background-color: #97CFA4'] * len(row)
    elif row['Status'] == 'DISCHARGING':
        return ['background-color: orange'] * len(row)
    else:
        return [''] * len(row)

def main():
    B22.main()
    file_path = 'B22.xlsx'
    df = load_data(file_path)

    charging_data, discharging_data = process_data(df)

    loading_sequences, unloading_sequences = load_unload_sequences(charging_data, discharging_data)

    average_energy_consumption, working_storage_count, air_conditioning = calculate_average_energy(loading_sequences, unloading_sequences, discharging_data, df)

    df['Energy Consumption by Storage [kWh]'] = df['Hour'].map(average_energy_consumption)
    df['Number of Working Storages'] = df['Hour'].map(working_storage_count)
    df['Air Conditioning [kWh]'] = df['Hour'].map(air_conditioning)

    df['Balance [kWh]'] = df['Energy Consumption [kWh]'] + df['Energy Consumption by Storage [kWh]'] + df['Air Conditioning [kWh]']

    df_styled = df.style.apply(color_rows, axis=1)
    df_styled.to_excel('simStorage1111.xlsx', index=False, engine='openpyxl')

if __name__ == "__main__":
    main()
