#!/usr/bin/env python3
"""Send training alert email."""

# Import libraries
from common import get_timestamp, read_configuration_file
from treport import check_failed_training_records
import pandas as pd
import win32com.client
import os


# Function for sending alert email for failed training
def send_failed_training_alert_email(config, df, display=False):
    """Send alert email for failed training."""
    # Iterate through each failed case
    for i, r in df.iterrows():
        # Load alert email template
        with open("template/t_reminder_failed.html", "r") as file:
            content = file.read()

        # Replace sender's details placeholders in email
        for key, value in config["email_sender"].items():
            try:
                content = content.replace("{{ " + key + " }}", value)

            except BaseException:
                pass

        # Replace placeholders in email content
        content = content.replace("{{ staff_name }}", r["Staff Name"])

        # Filter the dataframe by staff name
        df_t = df[df["Staff Name"] == r["Staff Name"]].copy()

        # Drop uneccessary columns
        df_t.drop(["Staff Name", "Staff No", "Refresh", "PassFlag",
                   "Organization Unit", "Organization Unit Desc",
                   "Remarks", "End_d", "Days Passed"
                   ], axis=1, inplace=True)

        # Convert dataframe to HTML table
        email_table = df_t.to_html(
            index=False, justify="center").replace(
                    '<td', '<td align="center"').replace(
                            '<table',
                            '<table border="1" class="dataframe"' +
                            'style="font-family: Courier New" cellpadding=24')

        email_table = email_table.replace(
                'Course Desc', 'Course Description')
        email_table = email_table.replace(
                'Start', 'Start Date')
        email_table = email_table.replace(
                'End', 'End Date')

        # Replace reminder table in email
        content = content.replace("{{ email_table }}", email_table)

        # Open Outlook application
        outlook = win32com.client.Dispatch('outlook.application')

        # Create new email
        mail = outlook.CreateItem(0)

        # Insert corporate logo in email
        attachment = mail.Attachments.Add(
                os.path.abspath(config["email_sender"]["corp_logo"]))
        attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo")

        # Receiver's email
        mail.To = config["email_sender"]["admin_email"]

        # Send email copy
        cc_list = []
        for cc in config["email_cc"]:
            cc_list.append(cc)
        for cc_exp in config["email_cc_expiry"]:
            cc_list.append(cc_exp)

        # Set CC
        mail.CC = "; ".join(cc_list)

        # Email subject
        mail.Subject = "Failed Training Assessment Alert on " + \
            get_timestamp(format="%d/%m/%Y")

        # Email body content in HTML
        mail.HTMLBody = content

        # Display email for checking
        if display:
            mail.Display()

        # Send email
        else:
            mail.Send()
            # Print confirmation on console
            print('[' + get_timestamp() +
                  "] Sent failed training alert email (" +
                  r["Staff Name"] +
                  ") to admin.")


if __name__ == '__main__':
    # Read configuration file
    config = read_configuration_file()

    # Check failed training records
    df_failed = check_failed_training_records(config, test_date=None)
    print(df_failed)

    # Send alert email
    send_failed_training_alert_email(config, df_failed, display=True)
