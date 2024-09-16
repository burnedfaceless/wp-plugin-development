#!/usr/bin/python3
territory_names = ['Alaska', 'Alabama', 'Arkansas', 'American Samoa', 'Arizona', 'California', 'Colorado', 'Connecticut', 'District of Columbia', 'Delaware', 'Florida', 'Georgia', 'Guam', 'Hawaii', 'Iowa', 'Idaho', 'Illinois', 'Indiana', 'Kansas', 'Kentucky', 'Louisiana', 'Massachusetts', 'Maryland', 'Maine', 'Michigan', 'Minnesota', 'Missouri', 'Mississippi', 'Montana', 'North Carolina', 'North Dakota', 'Nebraska', 'New Hampshire', 'New Jersey', 'New Mexico', 'Nevada', 'New York', 'Ohio', 'Oklahoma', 'Oregon', 'Pennsylvania', 'Puerto Rico', 'Rhode Island', 'South Carolina', 'South Dakota', 'Tennessee', 'Texas', 'Utah', 'Virginia', 'Virgin Islands', 'Vermont', 'Washington', 'Wisconsin', 'West Virginia', 'Wyoming']

territories_in_order = ['North Carolina', 'Georgia', 'Iowa', 'Alabama', 'Arkansas', 'Texas', 'Washington', 'Colorado', 'Nevada', 'New Mexico', 'Idaho', 'Virginia', 'Washington D.C.', 'Florida', 'Maine', 'Massachusetts', 'New York', 'Illinois']

# Convert the territories list to a set for faster lookup
territories_list_set = set(territory_names)

# Initialize an empty list to store the missing states
missing_territories = []

# Iterate over the second_array and check if each state is present in the state_names_set
for territory in territory_names:
    if territory not in territories_in_order:
        missing_territories.append(territory)

# Print the missing states
print('Missing territories:')
for territory in missing_territories:
    print(territory)
