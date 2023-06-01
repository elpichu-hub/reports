

file_path_1 = 'Quality Scoring Details ALL.csv'
file_path_2 = 'agents_list_do_not_delete.csv'

# Create a function to send the email
def send_email(subject, recipient, body, img_path=None, img_path_100=100):
    import email_config
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from email.mime.image import MIMEImage
    import os

    EMAIL_ADDRESS = email_config.EMAIL_ADDRESS
    EMAIL_PASSWORD = email_config.EMAIL_PASSWORD

    # Create the email message
    message = MIMEMultipart()
    message['From'] = EMAIL_ADDRESS
    message['To'] = recipient
    message['Subject'] = subject
    message.attach(MIMEText(body, 'html'))

    # If an image path is provided, add the image as an inline attachment
    if img_path is not None:
        with open(img_path, 'rb') as img_file:
            img_data = img_file.read()
        img_mime = MIMEImage(img_data)
        img_mime.add_header('Content-ID', '<{}>'.format(os.path.basename(img_path)))
        img_mime.add_header('Content-Disposition', 'inline', filename=os.path.basename(img_path))
        message.attach(img_mime)

    # If an img_path_100 is provided, add the image as an inline attachment
    if img_path_100 is not None:
        with open(img_path_100, 'rb') as img_file:
            img_data = img_file.read()
        img_mime = MIMEImage(img_data)
        img_mime.add_header('Content-ID', '<{}>'.format(os.path.basename(img_path_100)))
        img_mime.add_header('Content-Disposition', 'inline', filename=os.path.basename(img_path_100))
        message.attach(img_mime)

    # Connect to the Gmail SMTP server and send the email
    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(message)


def send_qa_to_work_email(file_path_1, file_path_2, date_range):
    import pandas as pd
    import chardet
    from datetime import datetime
    import os
    import encouraging_and_success_messages
    import random
    from email_config import team_lead_elite_signature

    # Detect the file encoding
    with open(file_path_1, 'rb') as file:
        result = chardet.detect(file.read())

    # Read the CSV file with the detected encoding without assigning column names
    df_temp = pd.read_csv(file_path_1, encoding=result['encoding'], header=None)

    # Assign column names
    column_names = ['A', 'dateTime', 'C', 'media'] + [f'column_{i}' for i in range(5, 10)] + ['user'] + ['column_11', 'totalScore'] + [f'column_{i}' for i in range(13, 18)] + ['escalationAgent'] + [f'column_{i}' for i in range(19, 25)] + ['gradingArea'] + [f'column_{i}' for i in range(26, 35)] + ['gradingAreaDescription', 'yesOrNot', 'column_AL', 'pointsWorth', 'pointsLost', 'column_AO', 'comments'] + [f'column_{i}' for i in range(41, len(df_temp.columns) + 1)]

    # Read the CSV file again with the assigned column names
    df = pd.read_csv(file_path_1, encoding=result['encoding'], header=None, names=column_names)

    df_columns_needed = df[['dateTime', 'media', 'user', 'totalScore', 'escalationAgent', 'gradingArea', 'gradingAreaDescription', 'yesOrNot', 'pointsWorth', 'pointsLost', 'comments']]

    # filter the rows of data that user and escalation agent are the same, if they
    # are not the same means that the call graded is for the escalation agent not 
    # my teams agent
    df_filtered = df_columns_needed[df_columns_needed['user'] == df_columns_needed['escalationAgent']]

    # Read the second Excel file
    df_agents = pd.read_csv(file_path_2)

    # Rename 'InternalId' column in df_agents to 'user' to match with df_filtered
    df_agents.rename(columns={'InternalId': 'user'}, inplace=True)

    df_filtered = df_filtered.copy()
    df_filtered['user'] = df_filtered['user'].astype(str)

    df_agents = df_agents.copy()
    df_agents['user'] = df_agents['user'].astype(str)

    # Merge the two dataframes
    df_filtered = pd.merge(df_filtered, df_agents[['user', 'FullName']], on='user', how='left')

    # Calculate the average score for each user
    average_scores = df_filtered.groupby('user')['totalScore'].mean()

    # Get the current date and time
    now = datetime.now()
    greeting = ""

    # Determine the appropriate greeting based on the current time
    if now.hour < 12:
        greeting = "Good morning"
    elif now.hour < 18:
        greeting = "Good afternoon"
    else:
        greeting = "Good evening"

    # Iterate over unique users and create the result string
    for user in df_filtered['user'].unique():
        user_data = df_filtered[df_filtered['user'] == user]
        average_score = average_scores.loc[user]
        # Here we get the FullName corresponding to the user
        full_name = user_data['FullName'].iloc[0]
        # If full_name is NaN or None, use the user
        if pd.isna(full_name) or full_name is None:
            full_name = user

        if average_score >= 85:
            random_success_message = encouraging_and_success_messages.success_messages[random.randint(0, len(encouraging_and_success_messages.success_messages) - 1)]
            result_string = f"<h3>{greeting} {full_name}, \n\nYour average QA for the week of {date_range} is <span style='color: green;'>{average_score:.2f}%</span>. {random_success_message} \n</h3>"
        else:
            random_ancouraging_message = encouraging_and_success_messages.encouraging_messages[random.randint(0, len(encouraging_and_success_messages.encouraging_messages) - 1)]
            result_string = f"<h3>{greeting} {full_name}, \n\nYour average QA for the week of {date_range} is <span style='color: red;'>{average_score:.2f}%</span>. <span style='background-color: #59FFA0;'>{random_ancouraging_message}</span> \n</h3>"

        # Get the current Directory
        current_dir = os.getcwd()

        # subject for Email
        subject = f'{full_name} QA Score Week Of {date_range}'
        # Find all calls that have comments and add them to the message
        # calls_with_comments = user_data[user_data['comments'].notna()]
        calls_with_comments = user_data
        call_number = 1
        prev_date_time = None
        call_index = 1
        image_path_100 = None
        for _, call in calls_with_comments.iterrows():
            date_time = call['dateTime']
            gradingArea = call['gradingAreaDescription']
            comments = call['comments']
            yes_or_not = call['yesOrNot']
            total_score = call['totalScore']
            points_worth = call['pointsWorth']

            if prev_date_time != date_time:
                if total_score >= 85 and total_score < 100:
                    result_string += f"<h3 style='text-decoration:underline;'>\n- Call {call_number} Recorded on {date_time} Scored <span style='color: green;'>{total_score}%</span></h3>"
                elif total_score == 100:
                    image_path_100 = os.path.join(current_dir, 'Perfect.jpg')
                    result_string += f"<h3 style='text-decoration:underline;'>\n- Call {call_number} Recorded on {date_time} Scored <span style='color: green;'>{total_score}%</span> \n</h3>"
                    result_string += "<img src='cid:{}' alt='Passed Image' style='height: 100px; width: 100px'><br>".format(os.path.basename(image_path_100))
                else:
                    result_string += f"<h3 style='text-decoration:underline;'>\n- Call {call_number} Recorded on {date_time} Scored <span style='color: red;'>{total_score}%</span></h3>"
                call_number += 1
                prev_date_time = date_time

            if pd.isna(comments):
                comments = " "  # Replace NaN with a space

    
            if yes_or_not == "No":
                result_string += f"<h4 style='margin-top: 0; margin-bottom: 0'>{call_index} - {gradingArea} <span style='background-color: #C1292E; color: white'>{yes_or_not}, {comments} / Points Lost: {points_worth}</span> \n</h4>"
                call_index += 1
            # else:
                # result_string += f"<h4 style='margin-top: 0; margin-bottom: 0'>{call_index} - {gradingArea} {yes_or_not}! \n</h4>"
                # call_index += 1

        # Body for email
        current_dir = os.getcwd()
        if average_score >= 85:
            image_path = os.path.join(current_dir, 'Passed.jpg')
        else:
            image_path = os.path.join(current_dir, 'Failed.jpg')
        body = "<html>" \
                "<body>" \
                "<img src='cid:{}' alt='Passed Image' style='width: 180px; height: 180px;'><br>".format(os.path.basename(image_path)) \
                + result_string.replace('\n', '<br>') + \
                team_lead_elite_signature + \
                "</body>" \
                "</html>"


        # Send the email for the current user
        send_email(subject, 'lazaro.gonzalez@conduent.com', body, image_path, image_path_100)

            #  # Write the message to a .txt file for the current user, sanitizing the filename
            # filename = sanitize_filename.sanitize(f'{full_name}_quality_assurance.txt')
            # with open(f'qas/{filename}', 'w', encoding='utf-8') as file:
            #     file.write(result_string)

    # Save the filtered data to a new CSV file
    # df_filtered.to_csv('filtered_data.csv', index=False)


# send_qa_to_work_email(file_path_1, file_path_2, date_range)