def qa_stats_proponisi_click(file_path1, file_path2, file_path3, file_path4=None):
    import pandas as pd
    import numpy as np

    # Read the existing data
    df1 = pd.read_csv(file_path1)
    date_string = df1['Date'].iloc[0]  

    # Read the agents list
    df2 = pd.read_csv(file_path2)

    # Merge the two dataframes based on 'InternalId' and 'FullName'
    merged = pd.merge(df1, df2[['InternalId', 'FullName']], on=['InternalId', 'FullName'], how='outer')


    # Fill missing values in specific columns
    merged['ACD Calls'].fillna(0, inplace=True)
    merged['AHT (seconds)'].fillna(0, inplace=True)
    merged['AVG Hold'].fillna('0:00:00', inplace=True)
    merged['AVG ACW'].fillna('0:00:00', inplace=True)
    merged['Date'].fillna(date_string, inplace=True)
    merged['QA'].fillna(0, inplace=True)  # Initialize 'QA' to zero where it's missing

    # Write the merged data back to the first file
    merged.to_csv(file_path1, index=False)

    # Read the data from file_path3, keeping only 'Scored User' and 'Total Score' columns
    df3 = pd.read_csv(file_path3)[['Scored User', 'Total Score']]

    # Remove extra spaces in 'Scored User' column
    df3['Scored User'] = df3['Scored User'].str.strip().str.replace(r'\s+', ' ', regex=True).str.replace(r'\(.*\)', '', regex=True)

    # Remove the percentage sign and convert the 'Total Score' column to numeric
    df3['Total Score'] = pd.to_numeric(df3['Total Score'].str.rstrip('%'), errors='coerce')

    # Initialize df_score as df3
    df_score = df3

    # If a fourth file path is provided, read the data from file_path4 as well
    if file_path4 is not None:
        df4 = pd.read_csv(file_path4)[['Scored User', 'Total Score']]

        # Remove extra spaces in 'Scored User' column
        df4['Scored User'] = df4['Scored User'].str.strip().str.replace(r'\s+', ' ', regex=True).str.replace(r'\(.*\)', '', regex=True)

        df4['Total Score'] = pd.to_numeric(df4['Total Score'].str.rstrip('%'), errors='coerce')

        # Concatenate df3 and df4
        df_score = pd.concat([df3, df4])

    # Reset the index of the new DataFrame
    df_score.reset_index(drop=True, inplace=True)

     # Group by 'Scored User' and calculate the average 'Total Score' for each user
    df_score = df_score.groupby('Scored User', as_index=False).agg({'Total Score': lambda x: x.mean(numeric_only=True)})

    print(df_score)


   # Merge the dataframes on 'FullName'
    merged_score = pd.merge(merged, df_score, left_on='FullName', right_on='Scored User', how='left')

    # If 'Total Score' is NaN after the merge, replace it with 0
    merged_score['Total Score'].fillna(0, inplace=True)

    # Add the 'Total Score' values to the 'QA' column
    merged_score['QA'] = merged_score['QA'] + merged_score['Total Score']

    # Replace 0 with NaN in the 'QA' column
    merged_score['QA'].replace(0, np.nan, inplace=True)

    # Drop the 'Total Score' and 'Scored User' columns as they're no longer needed
    merged_score.drop(['Total Score', 'Scored User'], axis=1, inplace=True)

    # Round the 'QA' column to the nearest whole number
    merged_score['QA'] = np.ceil(merged_score['QA'])

    # Fill NaN values in 'QA' column with a placeholder (e.g., '-1')
    merged_score['QA'].fillna(-1, inplace=True)

    # Convert 'QA' column values to strings and append '%' symbol, while replacing the placeholder with an empty string
    merged_score['QA'] = merged_score['QA'].astype(int).astype(str).replace('-1', '').apply(lambda x: x + '%' if x != '' else x)

    # Write the merged data back to the first file
    merged_score.to_csv(file_path1, index=False)

    # print(merged_score)

    # Write the merged data back to the first file
    merged_score.to_csv(file_path1, index=False)

    return