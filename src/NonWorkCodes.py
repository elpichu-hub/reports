def run_report(file_path):
    import os
    import pandas as pd
    
    # read the CSV file into a pandas dataframe
    df = pd.read_csv(file_path)

    # specify the column names you want to keep
    UserId = 'UserId'
    StatusDateTime = 'StatusDateTime'
    StatusKey = 'StatusKey'
    EndDateTime = 'EndDateTime'
    StateDuration = 'StateDuration'
    LastName = 'LastName'
    FirstName = 'FirstName'

    # keep only selected columns in the dataframe
    df_custom = df[[UserId, StatusDateTime, StatusKey, EndDateTime, StateDuration, LastName, FirstName]]

    # list of non-work codes
    non_work_codes = ['meeting', 'at a training session', 'it', 'mentoring', 'coaching', 'floor support', 'huddle session', 'after call work', 'in a meeting', 'at a meeting']
    break_codes = ['at lunch', 'scheduled break', 'unscheduled break',]

    # filter the dataframe based on the 'StatusKey' column containing any value in the non_work_codes list
    df_filtered = df_custom[df_custom[StatusKey].isin(non_work_codes)].copy()


    # fitler the dataframe based on the break_codes
    df_filtered_breaks = df_custom[df_custom[StatusKey].isin(break_codes)].copy()

    # add a new column called 'date' containing the date extracted from the 'StatusDateTime' column
    df_filtered['date'] = pd.to_datetime(df_filtered[StatusDateTime]).dt.date
    df_filtered['Agent Name'] = df_filtered['LastName'] + ', ' + df_filtered['FirstName']


    # convert 'StatusDateTime' and 'EndDateTime' columns to datetime format
    df_filtered[StatusDateTime] = pd.to_datetime(df_filtered[StatusDateTime])
    df_filtered[EndDateTime] = pd.to_datetime(df_filtered[EndDateTime])

    # filter the dataframe based on the 'StatusKey' column containing any value in the non_work_codes list
    df_filtered_15_more = df_filtered[(df_filtered[EndDateTime] - df_filtered[StatusDateTime]).dt.seconds >= 300].copy()


    df_filtered_breaks[StatusDateTime] = pd.to_datetime(df_filtered_breaks[StatusDateTime])
    df_filtered_breaks[EndDateTime] = pd.to_datetime(df_filtered_breaks[EndDateTime])
    

    # functions to generate reports
    def generate_report(df, report_name):
        df_export = df_filtered_15_more[['date', 'Agent Name', 'StatusKey', 'StatusDateTime', 'EndDateTime']]
        df_export.to_excel('replica.xlsx', index=False)

        with open('report.txt', 'a') as file:
            file.write('\n\n\n')
            file.write(f"Report: {report_name}\n")
            file.write("Created by: Lazaro Gonzalez\n\n")

            grouped_df = df.groupby([UserId, StatusKey, StateDuration]).agg({'StateDuration': 'count'}).rename(columns={'StateDuration': 'Count'}).reset_index()
            grouped_df = grouped_df.groupby([UserId, StatusKey]).agg({'Count': 'sum', 'StateDuration': 'sum'}).reset_index()

            for index, row in grouped_df.iterrows():
                user_id = row[UserId]
                first_name = df.loc[df[UserId] == user_id, FirstName].iloc[0]
                last_name = df.loc[df[UserId] == user_id, LastName].iloc[0]
                status_key = row[StatusKey]
                duration_minutes = row[StateDuration] / 60
                count = row['Count']

                filtered_by_user_code = df[(df[UserId] == user_id) & (df[StatusKey] == status_key)]
                
                if status_key == 'at lunch' and duration_minutes > 35:
                    exeed_warning_breaks = "**********************"
                elif status_key == 'scheduled break' and duration_minutes > 35:
                    exeed_warning_breaks = "**********************"
                elif status_key == 'unscheduled break' and duration_minutes > 10:
                    exeed_warning_breaks = "**********************"
                elif status_key == 'after call work' and duration_minutes > 3:
                    exeed_warning_breaks = "!!!!!!!!!!!!!!!!!!!!!!!!!!"
                else:
                    exeed_warning_breaks = ""

                file.write(f"{first_name} {last_name}: {status_key.capitalize()} {count} times, {duration_minutes:.2f} minutes {exeed_warning_breaks} \n")

                for _, event in filtered_by_user_code.iterrows():
                    start_time = event[StatusDateTime].strftime('%H:%M:%S')
                    end_time = event[EndDateTime].strftime('%H:%M:%S')
                    duration_minutes = event[StateDuration] / 60

                    exceed_warning = "*******************************" if status_key in non_work_codes and duration_minutes > 5 else ""

                    # print(f"{first_name} {last_name} {status_key} from {start_time} to {end_time} total {duration_minutes:.2f} minutes on {event[StatusDateTime].strftime('%m/%d/%Y')}.")
                    file.write(f"{first_name} {last_name} {status_key} from {start_time} to {end_time} total {duration_minutes:.2f} minutes on {event[StatusDateTime].strftime('%m/%d/%Y')}.{exceed_warning}\n")
                    # file.write(f"{event[StatusDateTime].strftime('%m/%d/%Y')} {last_name}, {first_name} {status_key} {start_time} {end_time} (COMMENT) Completed {exceed_warning}\n")
                file.write('-------------------------------------------------------------------------------------------\n')

            

    # run both reports
    generate_report(df_filtered, "Non-Work Codes Usage")
    # generate_report(df_filtered_breaks, "Break Codes Usage")



    # open the report file using the default application
    os.startfile('report.txt')

# run_report()