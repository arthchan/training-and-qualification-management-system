#!/usr/bin/env python3
"""Send alert email."""

# Import libraries
from common import get_timestamp, read_configuration_file
import pandas as pd
import win32com.client
import os


# Function for sending an alert email
def send_alert_email(config, html, failed=None, display=False):
    """Send an alert email."""
    # Open Outlook application
    outlook = win32com.client.Dispatch('outlook.application')

    # Create new email
    mail = outlook.CreateItem(0)

    # Insert corporate logo in email
    attachment = mail.Attachments.Add(
            os.path.abspath(config["email_sender"]["corp_logo"]))
    attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo")

    # Load alert email template
    with open("template/" + html + ".html", "r") as file:
        content = file.read()

    # Replace sender's details placeholders in email
    for key, value in config["email_sender"].items():
        try:
            content = content.replace("{{ " + key + " }}", value)

        except BaseException:
            pass

    # Get staff names in failed list
    if failed is not None:
        # Read staff list
        df_staff = pd.read_csv(config["staff_list_path"], dtype="string")
        failed_name = ""

        # Get number of failed cases
        number_of_failed = str(len(failed))

        # Iterate through all failed cases
        for s in failed:
            failed_name = failed_name + "<br> - " + df_staff.loc[
                    df_staff["Staff Number"] == str(s),
                    "Name"].values[0]

        # Replace placeholders in email
        content = content.replace("{{ number_of_failed }}", number_of_failed)
        content = content.replace("{{ failed_name }}", failed_name)

    else:
        pass

    # Receiver's email
    mail.To = config["email_sender"]["admin_email"]

    # Email subject
    mail.Subject = "Daily Qualification Enquiry Report on " + \
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
              "] Sent alert email to admin.")


if __name__ == '__main__':
    # Read configuration file
    config = read_configuration_file()

    # Send alert email
    send_alert_email(config, "q_alert_success",
                     failed=None, display=True)
