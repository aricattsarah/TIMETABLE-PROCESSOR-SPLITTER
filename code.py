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

def get_unique_sheet_name(base_name):
    name = base_name[:31]
    original = name
    i = 1
    while name.lower() in used_sheet_names:
        suffix = f"_{i}"
        name = (original[:31 - len(suffix)] + suffix) if len(original) + len(suffix) > 31 else original + suffix
        i += 1
    used_sheet_names.add(name.lower())
    return name

with pd.ExcelWriter("Classwise_Timetable.xlsx", engine="xlsxwriter") as writer:
    for cls in classes:
        if pd.isna(cls):
            continue
        class_df = weekly_timetable[weekly_timetable["Class"] == cls].copy()
        cols = ['Day'] + [col for col in class_df.columns if col not in ['Class', 'Day']]
        class_df = class_df[cols]
        sheet_name = get_unique_sheet_name(str(cls).strip())
        class_df.to_excel(writer, sheet_name=sheet_name, index=False)

#Download the output files
files.download("enter_name.xlsx")
