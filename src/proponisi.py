def seconds_to_time_format(seconds):
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"



# date_string = '4.24.2023'
def run_report(date_string, file_path1):
    import os
    import pandas as pd
    
    try:
        # read the CSV file into a pandas dataframe
        df = pd.read_csv(file_path1)   

        # specify the column names you want to keep
        user_id = 'i3user'
        acd_calls = 'totEnteredACD' # calls offered
        acd_answered = 'totnAnsweredACD' # calls answered
        avg_hold = 'totTHoldACD'
        avg_acw = 'totTACW'
        tot_talk = 'totTTalkACD'
        full_name = 'DisplayUserName'
        

        # keep only selected columns in the dataframe
        df_custom = df[[user_id, acd_calls, acd_answered, avg_hold, avg_acw, tot_talk, full_name]]

        # group the data by the 'i3user' column and calculate the total 'totEnteredACD',   'totTHoldACD', 'totTACW', and 'totTTalkACD'
        df_grouped = df_custom.groupby(user_id).agg({
            acd_calls: 'sum',
            acd_answered: 'sum',
            avg_hold: 'sum',
            avg_acw: 'sum',
            tot_talk: 'sum',
            full_name: 'first'
        }).reset_index()

        # Get the date from user input
        # date = input("Enter the date in yyyy-mm-dd format: ")

    
        if df_grouped.isna().any().any():
            df_grouped = df_grouped.fillna(0)
            print("Warning: The grouped data contains NaN values.")

        # Add the QA, Perfect Attendance, and Date columns to the dataframe
        df_grouped['InternalId'] = df_grouped[user_id].str.upper()
        df_grouped['FullName'] = df_grouped[full_name]
        df_grouped['ACD Calls'] = df_grouped[acd_answered]
        
        
        epsilon = 1e-8  # Small constant to avoid division by zero

        df_grouped['AVG Hold'] = (df_grouped[avg_hold] / (df_grouped[acd_answered] + epsilon)).apply(lambda x: round(x)).apply(seconds_to_time_format)
        df_grouped['AVG ACW'] = (df_grouped[avg_acw] / (df_grouped[acd_answered] + epsilon)).apply(lambda x: round(x)).apply(seconds_to_time_format)
        df_grouped['AHT (seconds)'] = ((df_grouped[tot_talk] + df_grouped[avg_acw]) / (df_grouped[acd_answered] + epsilon)).apply(lambda x: round(x))
        df_grouped['QA'] = ''
        df_grouped['Perfect Attendance'] = ''
        df_grouped['Date'] = date_string

        df_grouped['ACD Calls'] = df_grouped[acd_answered]

         # Remove rows where ACD Calls is 0
        df_grouped = df_grouped[df_grouped['ACD Calls'] != 0]

        # Remove extra spaces in FullName column
        df_grouped['FullName'] = df_grouped['FullName'].str.strip().str.replace(r'\s+', ' ', regex=True)

        # Sort the dataframe based on the 'FullName' column (A to Z)
        df_grouped = df_grouped.sort_values(by='FullName')
        
        # Reorder the columns as required
        columns = ['InternalId', 'FullName', 'ACD Calls', 'AHT (seconds)', 'AVG Hold', 'AVG ACW', 'QA', 'Perfect Attendance', 'Date']
        df_grouped = df_grouped[columns]

        # creates the dataframe to keep the user list,
        # every time the report is run on f
        df_custom_agents = df[[user_id, full_name]]
        df_custom_agents_grouped = df_custom_agents.groupby(user_id).agg({
            full_name:'first'
        }).reset_index()
        df_custom_agents_grouped[user_id] = df_custom_agents_grouped[user_id].str.upper()
        # Remove extra spaces in FullName column
        df_custom_agents_grouped[full_name] = df_custom_agents_grouped[full_name].str.strip().str.replace(r'\s+', ' ', regex=True)
        
        columns_agents_file = ['InternalId', 'FullName']
        df_custom_agents_grouped.columns = columns_agents_file

        print(df_custom_agents_grouped)
        
        # ...
        # Load existing agents list if exists, else create an empty DataFrame
        if os.path.exists('agents_list_do_not_delete.csv'):
            df_agents_list = pd.read_csv('agents_list_do_not_delete.csv')
        else:
            df_agents_list = pd.DataFrame(columns=['InternalId', 'FullName'])

        # Merge df_custom_agents_grouped with df_agents_list using 'outer' method
        df_updated_agents_list = pd.merge(df_agents_list, df_custom_agents_grouped, how='outer')

        # Remove duplicates by keeping the first occurrence of each unique agent
        df_updated_agents_list = df_updated_agents_list.drop_duplicates(subset=['InternalId'])

        # Save the updated agents list
        df_updated_agents_list.to_csv('agents_list_do_not_delete.csv', index=False, encoding='utf-8-sig')
        # ...

        # Export the data to a CSV file
        file_name = f"Proponisi_report_lazaro_gonzalez_{date_string}.csv"
        df_grouped.to_csv(file_name, index=False, encoding='utf-8-sig')

        os.startfile(file_name)

        dir_path = os.path.dirname(os.path.abspath(file_name))
        print("Report exported to:", dir_path)
        return f"Proponisi Report exported to: ${dir_path}"
    except Exception as e:
        print(e)
        return e


# run_report(date_string)

