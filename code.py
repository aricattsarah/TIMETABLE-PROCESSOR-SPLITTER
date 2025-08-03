#Import necessary libraries
import pandas as pd
from google.colab import files
#Upload and load the Excel file
uploaded = files.upload()
file_path = next(iter(uploaded))
xls = pd.ExcelFile(file_path)
# List of weekdays (sheet names) can be altered
weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

def extract_class_schedule_flat(df, day):
    df.columns = pd.MultiIndex.from_arrays([df.iloc[0], df.iloc[1]])
    df = df.iloc[2:].copy()
    df.columns = ['{} - {}'.format(str(a).strip(), str(b).strip()) if pd.notna(b) else str(a).strip()
                  for a, b in df.columns]
    df.rename(columns={df.columns[0]: 'Class'}, inplace=True)
    df['Day'] = day
    return df

all_days_data = []
for day in weekdays:
    df = xls.parse(day)
    cleaned_df = extract_class_schedule_flat(df, day)
    all_days_data.append(cleaned_df)
weekly_timetable = pd.concat(all_days_data, ignore_index=True)
classes = weekly_timetable["Class"].unique()
used_sheet_names = set()

