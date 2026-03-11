#!/usr/bin/env python3
"""Process training report."""

# Import libraries
from common import get_timestamp, read_configuration_file
import pandas as pd
import numpy as np
import glob

# Configure Pandas
pd.set_option('mode.chained_assignment', None)


# Function for generating report in CSV format
def generate_training_report(config):
    """Generate report in CSV format."""
    # Initialise dataframe for all data
    df_all = pd.DataFrame([])

    # Read staff list
    df_staff = pd.read_csv(config["staff_list_path"], dtype="string")

    # Read individual reports
    files = glob.glob("temp/T_*.csv")
    for f in files:
        # Filter away former staff
        sid = f.split('\\')[1].split("_")[2]
        if sid not in df_staff["Staff Number"].values:
            continue

        df = pd.read_csv(f)

        # Append data to main dataframe
        df_all = pd.concat([df_all, df], ignore_index=True)

    # Export report as CSV file in local folder
    df_all.to_csv(config["t_report_path"], index=False, encoding="utf_8_sig")

    # Export report as CSV file to Personal OneDrive
    try:
        df_all.to_csv(config["t_report_abs_path"], index=False,
                      encoding="utf_8_sig")
    except BaseException:
        df_all.to_csv(config["t_report_abs_path"].split(".csv")[0] + '_' +
                      get_timestamp(format="%Y%m%d-%H%M") + ".csv",
                      index=False, encoding="utf_8_sig")

    finally:
        pass

    return df_all


# Function for checking passed training records
def check_passed_training_records(config, test_date=None):
    """Check passed training records."""
    # Get date for testing
    if test_date is not None:
        today = np.datetime64(test_date, 'D')
    # Get today's date
    else:
        today = np.datetime64("today", 'D')

    # Initialise dataframe for reminder
    df_p = pd.DataFrame([])

    # Read report
    df = pd.read_csv(config["t_report_path"])
    df_q = pd.read_csv(config["q_report_path"])

    # Filter records which require job attachment
    for course in config["has_attachment"].keys():
        df_c = df[df["Course Code"] == course]

        # Change data type of course end dates
        df_c["End_d"] = pd.to_datetime(df_c["End"], format="%d/%m/%Y")

        # Calculate expiry dates
        df_c["Expiry_d"] = df_c["End_d"] + pd.Timedelta(
                days=config["has_attachment"][course][2])

        # Change data type of expiry dates to string format
        df_c["Expiry"] = df_c["Expiry_d"].dt.strftime("%d/%m/%Y")

        # Insert number of job attachment required
        df_c["Job Attachment Required"] = config["has_attachment"][course][1]

        # Calculate days to expiry
        df_c["Days Remaining"] = df_c["Expiry_d"] - today

        # Change data type of days remaining to integer
        df_c["Days Remaining"] = df_c["Days Remaining"].dt.days

        # Sort dataframe by days remaining
        df_c.sort_values(by="Days Remaining", ascending=False, inplace=True,
                         ignore_index=True)

        # Filter passed records based on reminder days
        df_cp = df_c[(df_c["Days Remaining"].isin(config[
            "has_attachment"][course][-1])) & (df_c["PassFlag"] == "Passed")]

        # Iterate through records to check if qualification is attained
        for i, r in df_cp.iterrows():
            if df_q[(df_q["Staff ID"] == r["Staff No"]) & (
                df_q["Qualification Code"] == config[
                    "has_attachment"][course][0])].empty is False:
                # Drop the record if qualification is attained
                df_cp.drop(i, inplace=True)

            else:
                pass

        # Append records to passed dataframe
        df_p = pd.concat([df_p, df_cp], ignore_index=True)

    return df_p


# Function for checking failed training records
def check_failed_training_records(config, test_date=None):
    """Check failed training records."""
    # Get date for testing
    if test_date is not None:
        today = np.datetime64(test_date, 'D')
    # Get today's date
    else:
        today = np.datetime64("today", 'D')

    # Initialise dataframe for reminder
    df_fo = pd.DataFrame([])

    # Read report
    df = pd.read_csv(config["t_report_path"])

    # Filter failed records
    df_f = df[df["PassFlag"] == "Failed"]

    if df_f.empty is False:
        # Change data type of course end dates
        df_f["End_d"] = pd.to_datetime(df_f["End"], format="%d/%m/%Y")

        # Calculate days passed since course end
        df_f["Days Passed"] = today - df_f["End_d"]

        # Change data type of days passed to integer
        df_f["Days Passed"] = df_f["Days Passed"].dt.days

        # Sort dataframe by days passed
        df_f.sort_values(by="Days Passed", inplace=True, ignore_index=True)

        # Filter failed records based on days passed since course end
        df_f = df_f[df_f["Days Passed"] <= 365]
        df_fo = df_f.copy()

        # Compare with past report
        if df_f.empty is False:
            try:
                df_fp = pd.read_csv("temp/F_Report.csv")

                for i, r in df_fo.iterrows():
                    if df_fp[(df_fp["Staff No"] == r["Staff No"]) & (
                        df_fp["Course Code"] == r["Course Code"]) & (
                            df_fp["End"] == r["End"])].empty is False:
                        # Drop the record if it is already in past report
                        df_fo.drop(i, inplace=True)

            except BaseException:
                pass

            finally:
                # Export failed records as CSV file in temp folder
                df_f.drop(columns=["End_d", "Days Passed"], inplace=True)
                df_f.to_csv(
                        "temp/F_Report.csv", index=False, encoding="utf_8_sig")

    return df_fo


if __name__ == "__main__":
    # Read configuration file
    config = read_configuration_file()

    # Generate report
    generate_training_report(config)

    # Check passed training records
    df_passed = check_passed_training_records(config, test_date=None)
    print(df_passed)

    # Check failed training records
    df_failed = check_failed_training_records(config, test_date=None)
    print(df_failed)
