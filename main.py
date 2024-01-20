import pandas as pd

input_file = 'Assignment_Timecard.xlsx'
output_file = 'output.txt'

# Read Excel Sheet to Dataframe
df = pd.read_excel(input_file, sheet_name='Sheet1')

# Convert string to datetime type
df['Time'] = pd.to_datetime(df['Time'], format='%m/%d/%Y %I:%M %p')
df['Time Out'] = pd.to_datetime(df['Time Out'], format='%m/%d/%Y %I:%M %p')
df['Timecard Hours (as Time)'] = pd.to_datetime(df['Timecard Hours (as Time)'], format='%H:%M')

def a_consecutive_days(df):

    # Find the min and max time in and time out for a particular day in case of multiple entries
    df = df.groupby(['Position ID', 'Employee Name', df['Time'].dt.date])['Time Out'].agg(['min', 'max']).reset_index()

    df['Time'] = pd.to_datetime(df['Time'])
    df['Date'] = pd.to_datetime(df['Time'].dt.date)

    # finds diff in date in consecutive rows
    df['Day Diff'] = df.groupby(['Position ID','Employee Name'])['Date'].diff().dt.days

    # Identify consecutive working days with a difference of 1 day and fill missing values with 0
    # Group by 'Position ID' and calculate the sum of consecutive day indicators for each position
    # Filter positions where the sum of consecutive days is greater than or equal to 6 (for 7 consecutive days, the sum of consecutive days will be 6)
    consecutive_7_days_df = df[df['Day Diff'].fillna(0).eq(1).groupby(df['Position ID']).transform('sum') >= 6]

    employee_name_id_df = consecutive_7_days_df[['Employee Name', 'Position ID']].drop_duplicates()
    employee_name_id_dict = dict(zip(employee_name_id_df['Employee Name'], employee_name_id_df['Position ID']))
    return employee_name_id_dict

def b_time_shift_difference(df):
    # find time difference between two timeouts
    # then subtract timecard (time worked) in 1 shift
    # that times total time spent between 2 shifts
    df['Time Out Difference'] = df.groupby(['Position ID','Employee Name'])['Time Out'].diff()
    hours = df['Timecard Hours (as Time)'].dt.hour
    minutes = df['Timecard Hours (as Time)'].dt.minute
    df['Time Between Shifts'] = df['Time Out Difference'] - pd.to_timedelta(hours * 60 + minutes, unit='m')

    less_than_10_more_than_1_df = df[(df['Time Between Shifts'] > pd.to_timedelta('1 hour')) & (df['Time Between Shifts'] < pd.to_timedelta('10 hours'))]

    employee_name_id_df = less_than_10_more_than_1_df[['Employee Name', 'Position ID']].drop_duplicates()
    employee_name_id_dict = dict(zip(employee_name_id_df['Employee Name'], employee_name_id_df['Position ID']))
    return employee_name_id_dict

def c_shift_duration(df):
    df['Shift Duration'] = df['Timecard Hours (as Time)'].dt.hour + df['Timecard Hours (as Time)'].dt.minute / 60

    more_than_14_hour_df = df[df['Shift Duration'] > 14]

    employee_name_id_df = more_than_14_hour_df[['Employee Name', 'Position ID']].drop_duplicates()
    employee_name_id_dict = dict(zip(employee_name_id_df['Employee Name'], employee_name_id_df['Position ID']))
    return employee_name_id_dict


dict_a = a_consecutive_days(df)
dict_b = b_time_shift_difference(df)
dict_c = c_shift_duration(df)


with open(output_file, "w+") as file:

    for question, d in zip(
        [
            "a. who has worked for 7 consecutive days.",
            "b. who have less than 10 hours of time between shifts but greater than 1 hour",
            "c. Who has worked for more than 14 hours in a single shift",
        ],
        [dict_a, dict_b, dict_c]
    ):

        print(question)
        file.write(question + "\n")

        for ename, eid in d.items():
            print(eid + "--" + ename)
            file.write(eid + "--" + ename + "\n")
        print("--"*10)
        file.write("--"*10 + "\n")

