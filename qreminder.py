#!/usr/bin/env python3
"""Send reminder email to staff."""

# Import libraries
from common import get_timestamp, read_configuration_file
from qrecord import fetch_practice_record
from qreport import analyse_report
import win32com.client
import numpy as np
import pandas as pd
import os


# Function for building reminder content
def build_reminder_content(config, mode, df):
    """Process email content for daily reminder."""
    # Load reminder email template
    with open("template/q_reminder_" + mode + ".html", "r") as file:
        reminder_html = file.read()

    # Fetch practice records
    df = fetch_practice_record(config, df)

    # Drop uneccessary columns
    df.drop(["Staff ID", "Qualification Code", "Status", "Note",
             "Organization Unit", "Organization Unit Desc", "Expiry_d",
             "First Obtain_d", "Due for Refresh/Examination", "Last Refresh_d"
             ], axis=1, inplace=True)

    # Drop name for daily reminder
    if mode == "daily":
        df.drop("Name", axis=1, inplace=True)

    # Change expiry date column to red in colour
    df = df.style.set_properties(
            **{"color": "red"}, subset=["Expiry"]).hide(
                    axis='index')

    # Convert dataframe to HTML table
    email_table = df.to_html(
            index=False, justify="center").replace(
                    '<td', '<td align="center"').replace(
                            '<table',
                            '<table border="1" class="dataframe"' +
                            'style="font-family: Courier New" cellpadding=24')

    # Make remaining days red in colour
    if mode == "daily":
        for rdr in config["remaining_days_red"]:
            try:
                email_table = email_table.replace(
                        'col5" >' + str(rdr) + '<',
                        'col5" style="color:red;">' + str(rdr) + '<')

            except BaseException:
                pass

    # Make practice count red in colour
    col_num = {"daily": '4', "quarterly": '5'}
    for pr in config["practice_red"]:
        try:
            email_table = email_table.replace(
                    'col' + col_num[mode] + '" >' + str(pr) + '<',
                    'col' + col_num[mode] + '" style="color:red;">' + str(
                        pr) + '<')

        except BaseException:
            pass

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


# Function for sending daily reminder email to staff
def send_daily_reminder_email(config, display=False, test_date=None):
    """Send daily reminder email to staff."""
    # Analyse report
    df_reminder = analyse_report(config, quarter_range=None,
                                 test_date=test_date)

    if df_reminder.empty:
        print('[' + get_timestamp() + "] No staff requires reminder email.")
        return

    # Read staff list
    df_staff = pd.read_csv(config["staff_list_path"], dtype="string")

    # Get list of all staff recieving the email
    staff_list = df_reminder["Staff ID"].unique()

    for s in staff_list:
        # Filter by staff ID
        df_email = df_reminder[df_reminder["Staff ID"] == s]

        # Build reminder content
        mail, content = build_reminder_content(config, "daily", df_email)

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
        if (df_email["Days Remaining"] <= 30).any():
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
        if (df_email["Days Remaining"] == 0).any():
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
        mail.Subject = "Reminder of Qualification Renewal on " + \
            get_timestamp(format="%d/%m/%Y")

        # Email body content in HTML
        mail.HTMLBody = content

        # Display email for inspection
        if display:
            mail.Display()

            # Print confirmation on console
            print('[' + get_timestamp() +
                  "] Prepared reminder email sending to " +
                  df_staff.loc[df_staff["Staff Number"] == str(s),
                               "Name"].item() + '.')

        # Send email
        else:
            mail.Send()

            # Print confirmation on console
            print('[' + get_timestamp() + "] Sent reminder email to " +
                  df_staff.loc[df_staff["Staff Number"] == str(s),
                               "Name"].item() + '.')


# Function for sending quarterly reminder to team head
def send_quarterly_reminder_email(
        config, q_num, month_start, month_end, display=False,
        test_date=None):
    """Send quarterly reminder email to team head."""
    # Form year string
    year = month_start.split("-")[0]

    # Form quarter date range
    quarter_range = np.arange(
            month_start, month_end, dtype="datetime64[D]"
            ).astype(str).tolist()

    # Analyse report
    df_reminder = analyse_report(config, quarter_range=quarter_range,
                                 test_date=test_date)

    # Read staff list
    df_staff = pd.read_csv(config["staff_list_path"], dtype="string")

    # Iterate through all teams
    for g in config["team_admin"]:

        # Get team member list
        df_team = df_staff[df_staff["Team"] == g]
        member_list = df_team["Staff Number"].tolist()

        # Filter by staff number in member list
        df_email = df_reminder[df_reminder["Staff ID"].astype(str).isin(
            member_list)]

        if df_email.empty:
            continue

        # Build reminder content
        mail, content = build_reminder_content(config, "quarterly",
                                               df_email)

        # Get team admin list
        team_admin_list = df_staff.loc[df_staff["Staff Number"].isin(
            config["team_admin"][g])]["Email Name"].tolist()

        if len(team_admin_list) <= 2:
            team_admin_name = " and ".join(team_admin_list)
        else:
            team_admin_name = "all"

        # Replace staff name placeholder in email
        content = content.replace("{{ team_admin_name }}", team_admin_name)

        # Replace year placeholder in email
        content = content.replace("{{ year }}", year)

        # Replace quarter placeholder in email
        content = content.replace("{{ quarter }}", q_num)

        # Replace team placeholder in email
        content = content.replace("{{ team }}", g)

        # Get team admin email
        team_admin_email = df_staff.loc[df_staff["Staff Number"].isin(
            config["team_admin"][g])]["Corporate Email"].tolist()

        # Receiver's email
        mail.To = "; ".join(team_admin_email)

        # Send email copy
        cc_list = []
        for cc in config["email_cc"]:
            cc_list.append(cc)
        mail.CC = "; ".join(cc_list)

        # Email subject
        mail.Subject = "Quarterly Reminder of Qualification Renewal in " +\
            year + " Q" + q_num

        # Email body content in HTML
        mail.HTMLBody = content

        # Display email for inspection
        if display:
            mail.Display()

            # Print confirmation on console
            print('[' + get_timestamp() +
                  "] Prepared reminder email sending to team admin of team " +
                  g + '.')

        # Send email
        else:
            mail.Send()

            # Print confirmation on console
            print('[' + get_timestamp() +
                  "] Sent reminder email to team admin of team " + g + '.')


if __name__ == "__main__":
    # Read configuration file
    config = read_configuration_file()

    test_mode = "daily"

    if test_mode == "daily":
        # Generate reminder email
        send_daily_reminder_email(config, display=True, test_date=None)

    elif test_mode == "quarterly":
        # Generate quarterly reminder email
        send_quarterly_reminder_email(config, "1", "2025-01",
                                      "2025-04", display=True, test_date=None)

    else:
        pass
