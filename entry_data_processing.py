import itertools

# Constants
HOLIDAY_MULTIPLIER = 0.1

# Function to extend energy consumption list
def extend_energy_consumption(original_list, hours):
    return list(itertools.islice(itertools.cycle(original_list), hours))

# Function to calculate energy consumption for holiday hours
def holiday_energy_consumption(origina_list):
    return [x * HOLIDAY_MULTIPLIER for x in origina_list]


