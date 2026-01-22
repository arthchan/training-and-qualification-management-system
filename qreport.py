#!/usr/bin/env python3
"""Process qualification report."""

# Import libraries
from common import get_timestamp, read_configuration_file
import pandas as pd
import numpy as np
import glob

# Configure Pandas
pd.set_option('mode.chained_assignment', None)


# Function for generating report in CSV format
def generate_qualification_report(config):
    """Generate report in CSV format."""
    # Initialise dataframe for all data
    df_all = pd.DataFrame([])

    # Read staff list
    df_staff = pd.read_csv(config["staff_list_path"], dtype="string")

    # Read individual reports
    files = glob.glob("reports/Q_*.csv")
    for f in files:
        # Filter away former staff
        sid = f.split('\\')[1].split("_")[2]
        if sid not in df_staff["Staff Number"].values:
            continue

        # Import report as dataframe
        df = pd.read_csv(f)

        # Set Note column as string
        df["Note"] = df["Note"].astype(str)

        # Insert staff number in dataframe
        df.insert(0, "Staff ID", sid)
        df.insert(1, "Name", f.split('\\')[1].split("_")[1])

        # Mark implied qualification
        for iq in config["implied_qualification"]:
            if df["Qualification Code"].str.contains(
                    "|".join(iq)).values.sum() >= 2:
                im_flag = False
                for i, iqq in enumerate(iq):
                    if (not im_flag) and (
                            iqq in df["Qualification Code"].values):
                        im_flag = True

                    elif iqq in df["Qualification Code"].values:
                        df.loc[df["Qualification Code"] == iqq,
                               "Note"] = "Implied"

                    else:
                        pass

            else:
                pass

        # Replace "nan" with empty string
        df["Note"] = df["Note"].replace("nan", '')

        # Append data to main dataframe
        df_all = pd.concat([df_all, df], ignore_index=True)

    # Export report as CSV file in local folder
    df_all.to_csv(config["q_report_path"], index=False, encoding='utf-8-sig')

    # Export report as CSV file to Personal OneDrive
    try:
        df_all.to_csv(config["q_report_abs_path"], index=False,
                      encoding='utf-8-sig')

    except BaseException:
        df_all.to_csv(config["q_report_abs_path"].split(".csv")[0] + '_' +
                      get_timestamp(format="%Y%m%d-%H%M") + ".csv",
                      index=False, encoding='utf-8-sig')

    finally:
        pass

    return df_all


# Function for analysing report
def analyse_report(config, quarter_range=None, test_date=None):
    """Analyse report."""
    # Get date for testing
    if test_date is not None:
        today = np.datetime64(test_date, 'D')
    # Get today's date
    else:
        today = np.datetime64("today", 'D')

    # Initialise dataframe for reminder
    df_reminder = pd.DataFrame([])

    # Read report
    df = pd.read_csv("temp/QReport.csv")

    # Move due dates to expiry dates
    df["Expiry"] = df["Expiry"].combine_first(
        df["Due for Refresh/Examination"])

    # Get list of available qualifications
    q_code_list = df["Qualification Code"].unique()

    for q in q_code_list:
        # Filter by qualification
        df_q = df[df["Qualification Code"] == q]

        # Skip if there is no expiry date or the qualification is bypassed
        if df_q["Expiry"].isnull().all() \
                or q in config["bypass_qualification"]:
            continue

        # If the qualification has an expiry date
        else:
            # Remove all rows without an expiry date
            df_qe = df_q[df_q["Expiry"].notnull()]

            # Change data type of expiry dates
            df_qe["Expiry_d"] = pd.to_datetime(
                df_qe["Expiry"], format="%d/%m/%Y")

            # Change data type of first obtain date
            df_qe["First Obtain_d"] = pd.to_datetime(
                df_qe["First Obtain"], format="%d/%m/%Y")

            # No need to check for remaining days for quarterly report
            if quarter_range is not None:
                df_qe = df_qe[df_qe["Expiry_d"].astype(
                    str).isin(quarter_range)]
                df_reminder = pd.concat([df_reminder, df_qe],
                                        ignore_index=True)
                continue

            # Get the number of day(s) between today and expiry date
            df_qe["Days Remaining"] = df_qe["Expiry_d"] - today

            # Change data type of remaining days
            df_qe["Days Remaining"] = df_qe["Days Remaining"].dt.days

            # Append data to reminder dataframe according to table in config
            if q in config["remaining_days_table"]:
                for r in config["remaining_days_table"][q]:
                    df_reminder = pd.concat([df_reminder, df_qe[df_qe[
                        "Days Remaining"] == r]], ignore_index=True)
            else:
                for r in config["remaining_days_table"]["DEFAULT"]:
                    df_reminder = pd.concat([df_reminder, df_qe[df_qe[
                        "Days Remaining"] == r]], ignore_index=True)

    # Add column for refresher
    if df_reminder.empty is False:
        df_reminder["Refresher"] = np.nan
    else:
        return df_reminder

    # Remove implied qualification
    df_reminder.drop(df_reminder[df_reminder["Note"] == "Implied"].index,
                     inplace=True)

    # Check if refresher training is required for all staff
    for s in df_reminder["Staff ID"].unique():
        df_reminder_s = df_reminder[
                df_reminder["Staff ID"] == s].convert_dtypes()

        for t in config["has_refresher"]:
            has_refresher_q = t[0:-1]
            repeat_year = t[-1]

            for q in has_refresher_q:
                if df_reminder_s[
                        df_reminder_s["Qualification Code"] == q].empty:
                    continue

                else:
                    df_refresher = df_reminder_s[df_reminder_s[
                        "Qualification Code"] == q]

                    # Calculate if refresher training is required
                    training_array = (pd.DatetimeIndex(
                        df_refresher["Expiry_d"]).year - pd.DatetimeIndex(
                            df_refresher["First Obtain_d"]).year) % repeat_year

                    df_refresher["Refresher"] = training_array

                    # Update results to reminder dataframe
                    df_reminder.update(df_refresher)

    # Change data type of refresher training column
    df_reminder["Refresher"] = df_reminder["Refresher"].astype("string")

    # Indicate if refresher training is required
    df_reminder["Refresher"] = df_reminder["Refresher"].replace(np.nan, '-')
    df_reminder["Refresher"] = df_reminder["Refresher"].replace("0.0", 'Y')
    df_reminder["Refresher"] = df_reminder["Refresher"].replace(
            r'\d+.\d+', 'N', regex=True)

    if quarter_range is None:
        # Change remaining days to integers
        df_reminder["Days Remaining"] = df_reminder[
                "Days Remaining"].astype(int)

        # Sort reminder dataframe by staff ID and days remaining
        df_reminder.sort_values(
                ["Staff ID", "Days Remaining", "Qualification Code"],
                inplace=True)

    else:
        # Sort reminder dataframe by expiry date and staff ID
        df_reminder.sort_values(
                ["Expiry_d", "Staff ID", "Qualification Code"],
                inplace=True)

    return df_reminder


if __name__ == "__main__":
    # Read configuration file
    config = read_configuration_file()

    # Generate report
    generate_qualification_report(config)

    # Analyse report
    #print(analyse_report(config, quarter_range=None, test_date=None))
