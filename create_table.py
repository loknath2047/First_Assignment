import pandas as pd
input_file = 'input_data.xlsx'
df = pd.read_excel(input_file)
df_sorted = df.sort_values(by='Priority').reset_index(drop=True)
days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
periods_per_day = 6
timetable = pd.DataFrame(index=days, columns=[f'Period {i+1}' for i in range(periods_per_day)])
def find_next_available_period(timetable, subject):
    for day in timetable.index:
        for period in timetable.columns:
            if pd.isna(timetable.at[day, period]):
                return day, period
    return None, None
for _, row in df_sorted.iterrows():
    teacher = row['Teacher']
    subject = row['Subject']
    num_classes = row['Number of classes per week']
    
    for _ in range(num_classes):
        day, period = find_next_available_period(timetable, subject)
        if day and period:
            timetable.at[day, period] = f'{subject} ({teacher})'
        else:
            ("No available periods to schedule all classes. Adjust timetable configuration.")
output_file = 'output_data.xlsx'
timetable.to_excel(output_file, engine='openpyxl')
print(f'Timetable has been saved to {output_file}')