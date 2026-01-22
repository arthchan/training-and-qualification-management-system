#!/usr/bin/env python3
"""Running the command line interface console."""

# Import libraries
from common import get_timestamp, read_configuration_file
from qalert import send_alert_email
from qrecord import fetch_qualification_record
from qreminder import send_daily_reminder_email, send_quarterly_reminder_email
from qreport import generate_qualification_report
from trecord import fetch_training_record
from treport import generate_training_report
import pandas as pd
import schedule
import time


# Function for fetching qualification records and send reminders daily
def run_daily_enquiry_routine():
    """Run daily enquiry routine."""
    # Read configuration file
    config = read_configuration_file()

    # Fetch qualification records
    failed = fetch_qualification_record(config)

    # Send alert email to admin
    if len(failed) == 0:
        send_alert_email(config, "q_alert_success")

    elif len(failed) < len(pd.read_csv(config["staff_list_path"])):
        send_alert_email(config, "q_alert_partial_success", failed)

    else:
        send_alert_email(config, "q_alert_failure")

    # Fetch training records
    fetch_training_record(config)

    # Generate training report
    generate_training_report(config)


# Function for sending daily reminder email
def run_reminder_routine():
    """Run reminder routine."""
    # Read configuration file
    config = read_configuration_file()

    # Generate report
    generate_qualification_report(config)

    # Send daily reminder email
    send_daily_reminder_email(config)

    # Get current date
    ddmm = get_timestamp(format="%d/%m")
    yyyy = get_timestamp(format="%Y")

    # Check and send quarterly reminder email
    if ddmm == "01/12":
        yyyy = str(int(yyyy) + 1)
        send_quarterly_reminder_email(config, "1", yyyy + "-01", yyyy + "-04")
    elif ddmm == "01/03":
        send_quarterly_reminder_email(config, "2", yyyy + "-04", yyyy + "-07")
    elif ddmm == "01/06":
        send_quarterly_reminder_email(config, "3", yyyy + "-07", yyyy + "-10")
    elif ddmm == "01/09":
        send_quarterly_reminder_email(config, "4", yyyy + "-10",
                                      str(int(yyyy) + 1) + "-01")
    else:
        pass


if __name__ == "__main__":
    # Start the command line interface console
    print("[" + get_timestamp() + "] Starting the programme...")

    # Read configuration file
    config = read_configuration_file()

    # Schedule the routines
    schedule.every().day.at(config["fetch_time"]).do(run_daily_enquiry_routine)
    schedule.every().day.at(config["reminder_time"]).do(run_reminder_routine)

    while True:
        try:
            # Run the routines
            schedule.run_pending()
            time.sleep(1)

        except KeyboardInterrupt:
            break

    # Quit the programme
    print("[" + get_timestamp() + "] The programme has been terminated.")
    quit()
