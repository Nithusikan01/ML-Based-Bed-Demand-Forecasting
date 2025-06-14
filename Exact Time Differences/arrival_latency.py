'''import pandas as pd

# Load the Excel file into a DataFrame
df = pd.read_excel('Data.xlsx')

# Choose a specific range of rows (e.g., from row 0 to 4)
start_row = 0
end_row = 4
selected_data = df.loc[start_row:end_row].copy()  # Use .copy() to create a copy of the DataFrame

# Convert 'injury_time' and 'arrival_time' to datetime objects
selected_data.loc[:, 'injury_time'] = pd.to_datetime(selected_data['injury_time'])
selected_data.loc[:, 'arrival_time'] = pd.to_datetime(selected_data['arrival_time'])

# Calculate the time difference for each row
selected_data.loc[:, 'time_difference'] = selected_data['arrival_time'] - selected_data['injury_time']

# Extract days, hours, minutes, and seconds
selected_data.loc[:, 'days'] = selected_data['time_difference'].dt.days
selected_data.loc[:, 'hours'], remainder = divmod(selected_data['time_difference'].dt.seconds, 3600)
selected_data.loc[:, 'minutes'], selected_data.loc[:, 'seconds'] = divmod(remainder, 60)

# Convert the formatted time difference to seconds and store as integer
selected_data.loc[:, 'formatted_time_difference'] = (
    selected_data['days']*24*60*60 +
    selected_data['hours']*60*60 +
    selected_data['minutes']*60 +
    selected_data['seconds']
)

# Save the modified DataFrame to 'data.xlsx'
selected_data.to_excel('Data.xlsx', index=False)'''


################################################################################
'''import pandas as pd

# Load the Excel file into a DataFrame
df = pd.read_excel('data.xlsx')

# Convert 'injury_time' and 'arrival_time' to datetime objects for the entire DataFrame
df.loc[:, 'injury_time'] = pd.to_datetime(df['injury_time'])
df.loc[:, 'arrival_time'] = pd.to_datetime(df['arrival_time'])

# Calculate the time difference for each row
df.loc[:, 'time_difference'] = df['arrival_time'] - df['injury_time']

# Extract days, hours, minutes, and seconds
df.loc[:, 'days'] = df['time_difference'].dt.days
df.loc[:, 'hours'], remainder = divmod(df['time_difference'].dt.seconds, 3600)
df.loc[:, 'minutes'], df.loc[:, 'seconds'] = divmod(remainder, 60)

# Convert the formatted time difference to seconds and store as integer
df.loc[:, 'formatted_time_difference'] = (
    df['days']*24*60*60 +
    df['hours']*60*60 +
    df['minutes']*60 +
    df['seconds']
)

# Save the modified DataFrame to 'data.xlsx'
df.to_excel('data.xlsx', index=False)'''

###########################################################

'''import pandas as pd

# Load the Excel file into a DataFrame
df = pd.read_excel('data.xlsx')
print(df.columns)

# Convert 'injury_time' and 'arrival_time' to datetime objects for the entire DataFrame
df.loc[:, 'arrival_time'] = pd.to_datetime(df['arrival_time'])
df.loc[:, 'eddischarge_time'] = pd.to_datetime(df['eddischarge_time'])

# Calculate the time difference for each row
df.loc[:, 'time_difference'] = df['eddischarge_time'] - df['arrival_time']

# Extract days, hours, minutes, and seconds
df.loc[:, 'days'] = df['time_difference'].dt.days
df.loc[:, 'hours'], remainder = divmod(df['time_difference'].dt.seconds, 3600)
df.loc[:, 'minutes'], df.loc[:, 'seconds'] = divmod(remainder, 60)

# Convert the formatted time difference to seconds and store as integer
df.loc[:, 'formatted_time_difference'] = (
    df['days']*24*60*60 +
    df['hours']*60*60 +
    df['minutes']*60 +
    df['seconds']
)

# Save the modified DataFrame to 'data.xlsx'
df.to_excel('data.xlsx', index=False)'''

'''import pandas as pd
from openpyxl import load_workbook

# Load the Excel file into a DataFrame
# df = pd.read_excel('Data.xlsx')
try:
    df = pd.read_excel('Data.xlsx', engine='openpyxl')
except ValueError:
    # If the 'openpyxl' engine doesn't work, try 'xlrd' engine
    df = pd.read_excel('Data.xlsx', engine='xlrd')


# Convert 'injury_time' and 'arrival_time' to datetime objects for the entire DataFrame
df['injury_time'] = pd.to_datetime(df['injury_time'])
df['arrival_time'] = pd.to_datetime(df['arrival_time'])

# Calculate the time difference for each row
df['time_difference'] = df['arrival_time'] - df['injury_time']

# Convert the time difference to a formatted string in the "day:hour:minute:second" format
df['formatted_time_difference'] = df['time_difference'].apply(
    lambda x: '{:02}:{:02}:{:02}:{:02}'.format(x.days, x.seconds // 3600, (x.seconds // 60) % 60, x.seconds % 60)
)

# Load the existing Excel file
book = load_workbook('data.xlsx')

# Create a Pandas Excel writer using the existing Excel file
writer = pd.ExcelWriter('data.xlsx', engine='openpyxl')
writer.book = book

# Add the new column to the existing sheet
df[['formatted_time_difference_injury_vs_arrival']].to_excel(writer, sheet_name='Sheet1', index=False, startcol=df.shape[1])

# Save changes
writer.save()'''


### Correct Code ###
'''import pandas as pd

# Load the Excel file into a DataFrame
df = pd.read_excel('Data.xlsx')

# Convert 'injury_time' and 'arrival_time' to datetime objects for the entire DataFrame
df.loc[:, 'injury_time'] = pd.to_datetime(df['injury_time'])
df.loc[:, 'arrival_time'] = pd.to_datetime(df['arrival_time'])

# Calculate the time difference for each row
df.loc[:, 'time_difference(days)_arrival_vs_injury'] = df['arrival_time'] - df['injury_time']

# Extract days, hours, minutes, and seconds
df.loc[:, 'days'] = df['time_difference(days)_arrival_vs_injury'].dt.days
df.loc[:, 'hours'], remainder = divmod(df['time_difference(days)_arrival_vs_injury'].dt.seconds, 3600)
df.loc[:, 'minutes'], df.loc[:, 'seconds'] = divmod(remainder, 60)

# Create a new column with formatted time difference
df['formatted_time_difference_injury_vs_injury'] = (
    df['days'].astype(str) + ' days ' +
    df['hours'].astype(str) + ' hours ' +
    df['minutes'].astype(str) + ' minutes ' +
    df['seconds'].astype(str) + ' seconds '
)

# Save the modified DataFrame to 'data.xlsx'
df.to_excel('Data.xlsx', index=False)'''

'''import pandas as pd

# Load the Excel file into a DataFrame
df = pd.read_excel('Data.xlsx')

# Convert 'injury_time' and 'arrival_time' to datetime objects for the entire DataFrame
df.loc[:, 'arrival_time'] = pd.to_datetime(df['arrival_time'])
df.loc[:, 'eddischarge_time'] = pd.to_datetime(df['eddischarge_time'])

# Calculate the time difference for each row
df.loc[:, 'time_difference(days)_arrival_vs_eddischarge'] = df['eddischarge_time'] - df['arrival_time']

# Extract days, hours, minutes, and seconds
df.loc[:, 'days_arrival_vs_eddischarge'] = df['time_difference(days)_arrival_vs_eddischarge'].dt.days
df.loc[:, 'hours_arrival_vs_eddischarge'], remainder = divmod(df['time_difference(days)_arrival_vs_eddischarge'].dt.seconds, 3600)
df.loc[:, 'minutes_arrival_vs_eddischarge'], df.loc[:, 'seconds'] = divmod(remainder, 60)

# Create a new column with formatted time difference
df['formatted_time_difference_injury_vs_injury'] = (
    df['days'].astype(str) + ' days ' +
    df['hours'].astype(str) + ' hours ' +
    df['minutes'].astype(str) + ' minutes ' +
    df['seconds'].astype(str) + ' seconds '
)

# Save the modified DataFrame to 'data.xlsx'
df.to_excel('Data.xlsx', index=False)'''

import pandas as pd

# Load the Excel file into a DataFrame
df = pd.read_excel('Data.xlsx')

# Convert 'injury_time' and 'arrival_time' to datetime objects for the entire DataFrame
df.loc[:, 'arrival_time'] = pd.to_datetime(df['arrival_time'])
df.loc[:, 'discharge_time'] = pd.to_datetime(df['discharge_time'])

# Calculate the time difference for each row
df.loc[:, 'time_difference(days)_arrival_vs_discharge'] = df['discharge_time'] - df['arrival_time']

# Extract days, hours, minutes, and seconds
df.loc[:, 'days_arrival_vs_discharge'] = df['time_difference(days)_arrival_vs_discharge'].dt.days
df.loc[:, 'hours_arrival_vs_discharge'], remainder = divmod(df['time_difference(days)_arrival_vs_discharge'].dt.seconds, 3600)
df.loc[:, 'minutes_arrival_vs_discharge'], df.loc[:, 'seconds'] = divmod(remainder, 60)

# Create a new column with formatted time difference
'''df['formatted_time_difference_injury_vs_injury'] = (
    df['days'].astype(str) + ' days ' +
    df['hours'].astype(str) + ' hours ' +
    df['minutes'].astype(str) + ' minutes ' +
    df['seconds'].astype(str) + ' seconds '
)'''

# Save the modified DataFrame to 'data.xlsx'
df.to_excel('Data.xlsx', index=False)





