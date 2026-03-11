#!/usr/bin/env python3
"""Send training reminder email to staff."""

# Import libraries
from common import get_timestamp, read_configuration_file
from treport import check_passed_training_records
import win32com.client
import pandas as pd
import os


# Function for building training reminder content
def build_training_reminder_content(config, df):
    """Process email content for training reminder."""
    # Load reminder email template
    with open("template/t_reminder_passed.html", "r") as file:
        reminder_html = file.read()

    # Drop uneccessary columns
    df.drop(["Staff Name", "Staff No", "Course Code", "Start", "End",
             "Refresh", "PassFlag", "Organization Unit",
             "Organization Unit Desc", "Remarks", "Expiry_d", "Expiry", "End_d"
             ], axis=1, inplace=True)

    # Change remaining days column to red in colour
    df = df.style.set_properties(
            **{"color": "red"}, subset=["Days Remaining"]).hide(
                    axis='index')

    # Convert dataframe to HTML table
    email_table = df.to_html(
            index=False, justify="center").replace(
                    '<td', '<td align="center"').replace(
                            '<table',
                            '<table border="1" class="dataframe"' +
                            'style="font-family: Courier New" cellpadding=24')

    email_table = email_table.replace(
            'Job Attachment Required', 'Job Attachment<br>Required')
    email_table = email_table.replace(
            'Course Desc', 'Course Description')

    # Open Outlook application
    outlook = win32com.client.Dispatch('outlook.application')

    # Create new email
    mail = outlook.CreateItem(0)

    # Insert corporate logo in email
    attachment = mail.Attachments.Add(
            os.path.abspath(config["email_sender"]["corp_logo"]))
    attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo")

    # Compose email body content
    content = reminder_html

    # Replace reminder table in email
    content = content.replace("{{ email_table }}", email_table)

    # Replace sender's details placeholders in email
    for key, value in config["email_sender"].items():
        try:
            content = content.replace("{{ " + key + " }}", value)

        except BaseException:
            pass

    return mail, content


# Function for sending training reminder email to staff
def send_training_reminder_email(config, display=False, test_date=None):
    """Send training reminder email to staff."""
    # Check training records
    df_passed = check_passed_training_records(config, test_date=test_date)

    if df_passed.empty:
        print('[' + get_timestamp() +
              "] No staff requires training reminder email.")
        return

    # Read staff list
    df_staff = pd.read_csv(config["staff_list_path"], dtype="string")

    # Get list of all staff recieving the email
    staff_list = df_passed["Staff No"].unique()

    for s in staff_list:
        # Filter by staff ID
        df_email_pass = df_passed[df_passed["Staff No"] == s]

        # Build reminder content
        mail, content = build_training_reminder_content(config, df_email_pass)

        # Get staff name for email
        staff_email_name = df_staff.loc[
                df_staff["Staff Number"] == str(s), "Email Name"].item()

        # Replace staff name placeholder in email
        content = content.replace("{{ staff_name }}", staff_email_name)

        # Receiver's email
        receipient_email = df_staff.loc[
            df_staff["Staff Number"] == str(s), "Corporate Email"].item()
        mail.To = receipient_email

        # Send email copy
        cc_list = []
        for cc in config["email_cc"]:
            cc_list.append(cc)

        # Get team admin list if any qualification is expiring within 30 days
        if (df_email_pass["Days Remaining"] <= 30).any():
            g = df_staff.loc[df_staff["Staff Number"] == str(s), "Team"].item()
            team_admin_list = df_staff.loc[df_staff["Staff Number"].isin(
                config["team_admin"][g])]["Corporate Email"].tolist()

            # Remove admin email from team admin list
            if config["email_sender"]["admin_email"] in team_admin_list:
                team_admin_list.remove(config["email_sender"]["admin_email"])

            # Concatenate CC list with team admin list
            if receipient_email not in team_admin_list:
                cc_list = cc_list + team_admin_list

        # Add extra CC if any qualification is expiring today
        if (df_email_pass["Days Remaining"] == 0).any():
            for cc_exp in config["email_cc_expiry"]:
                cc_list.append(cc_exp)

        # Remove duplicate recipients
        cc_list = list(set(cc_list))

        # Remove receipient email from CC list if exists
        if receipient_email in cc_list:
            cc_list.remove(receipient_email)

        # Set CC
        mail.CC = "; ".join(cc_list)

        # Email subject
        mail.Subject = "Reminder to Complete Job Attachement and/or " + \
                "Oral Examination on " + get_timestamp(format="%d/%m/%Y")

        # Email body content in HTML
        mail.HTMLBody = content

        # Display email for inspection
        if display:
            mail.Display()

            # Print confirmation on console
            print('[' + get_timestamp() +
                  "] Prepared training reminder email sending to " +
                  df_staff.loc[df_staff["Staff Number"] == str(s),
                               "Name"].item() + '.')

        # Send email
        else:
            mail.Send()

            # Print confirmation on console
            print('[' + get_timestamp() +
                  "] Sent training reminder email to " +
                  df_staff.loc[df_staff["Staff Number"] == str(s),
                               "Name"].item() + '.')


if __name__ == "__main__":
    # Read configuration file
    config = read_configuration_file()

    # Send training reminder email to staff
    send_training_reminder_email(config, display=True, test_date=None)
