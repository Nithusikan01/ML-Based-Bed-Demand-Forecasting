import pandas as pd

# Read your dataset into a Pandas DataFrame
df = pd.read_excel("Data.xlsx")

# Convert "arrival time" column to datetime format
df['arrival_time'] = pd.to_datetime(df['arrival_time'])

# Extract just the date from the "arrival time"
df['arrival_date'] = df['arrival_time'].dt.date

# Initialize a list to store the results
result_data = []

# Loop through each unique date in the dataset
for date in df['arrival_date'].unique():
    # Filter the dataframe to include patients who are still in the hospital on the given date
    patients_today = df[(df['arrival_date'] <= date) & (pd.to_datetime(df['arrival_date']) + pd.to_timedelta(df['total_los'], unit='D') >= pd.to_datetime(date))].shape[0]
    
    # Calculate the number of patients for day before, 3 days before, 5 days before, and 7 days before
    patients_day_before = df[(pd.to_datetime(df['arrival_date']) <= pd.to_datetime(date) - pd.Timedelta(days=1)) & (pd.to_datetime(df['arrival_date']) + pd.to_timedelta(df['total_los'], unit='D') >= pd.to_datetime(date) - pd.Timedelta(days=1))].shape[0]
    patients_3_days_before = df[(pd.to_datetime(df['arrival_date']) <= pd.to_datetime(date) - pd.Timedelta(days=2)) & (pd.to_datetime(df['arrival_date']) + pd.to_timedelta(df['total_los'], unit='D') >= pd.to_datetime(date) - pd.Timedelta(days=3))].shape[0]
    patients_5_days_before = df[(pd.to_datetime(df['arrival_date']) <= pd.to_datetime(date) - pd.Timedelta(days=3)) & (pd.to_datetime(df['arrival_date']) + pd.to_timedelta(df['total_los'], unit='D') >= pd.to_datetime(date) - pd.Timedelta(days=5))].shape[0]
    patients_7_days_before = df[(pd.to_datetime(df['arrival_date']) <= pd.to_datetime(date) - pd.Timedelta(days=7)) & (pd.to_datetime(df['arrival_date']) + pd.to_timedelta(df['total_los'], unit='D') >= pd.to_datetime(date) - pd.Timedelta(days=7))].shape[0]
    
    # Append the results to the result list
    result_data.append({'Date': date, 
                        'Patients': patients_today, 
                        'Patients_Day_Before': patients_day_before, 
                        'Patients_2_Days_Before': patients_3_days_before, 
                        'Patients_3_Days_Before': patients_5_days_before, 
                        'Patients_7_Days_Before': patients_7_days_before})

# Create a DataFrame from the result list
result_df = pd.DataFrame(result_data)

# Write the new DataFrame to an Excel file
result_df.to_excel("patient_counts12.xlsx", index=False)
