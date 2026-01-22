#!/usr/bin/env python3
"""Fetch training record."""

# Import libraries
from common import get_timestamp, get_time_difference, read_configuration_file
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
import glob
import os
import time


# Function to fetch training record
def fetch_training_record(config):
    """Fetch training record."""
    # Read staff list
    df = pd.read_csv(config['staff_list_path'], dtype="string")

    # Initialise webdriver
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    web = webdriver.Chrome(options=options)

    # Get user path
    user_path = "C:\\Users\\" + os.getlogin()

    # Initialise an array to store all failed cases
    failed = []

    print("[" + get_timestamp() + "] Fetching staff training record...")

    # Iterate through all staff
    for d, s in enumerate(df["Staff Number"]):
        try:
            staff_id = s

            # Remove all training record files in Downloads folder
            file_list = glob.glob(user_path + "\\Downloads\\TrainResult*")

            if len(file_list) > 0:
                for file in file_list:
                    os.remove(file)

            # Browse webpage
            web.get(config["enquiry_training_link"])

            # Find input field for staff number
            staff_id_input = WebDriverWait(web, 10).until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        '//*[@id="ctl00_cphContent_txtTrainingStaffNo_' +
                        'txtStaffNo"]')))

            # Fill in staff number
            staff_id_input.send_keys(staff_id)

            # Find "Data Download" button
            download_button = WebDriverWait(web, 10).until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        '//*[@id="ctl00_cphContent_btnDown"]')))

            # Click "Data Download" button
            download_button.click()

            # Record download start time and initialise end time
            start_time = datetime.now()
            end_time = datetime.now()

            # Check if the file exists
            download_flag = False
            while get_time_difference(start_time, end_time) < 180:
                file_list = glob.glob(user_path +
                                      "\\Downloads\\TrainResult*.xls")
                if len(file_list) > 0:
                    download_flag = True
                    break

                time.sleep(5)
                end_time = datetime.now()

            # Continue iteration if file download is unsuccessful
            if not download_flag:
                print("[" + get_timestamp() +
                      "] Failed to download training record for " +
                      df[df["Staff Number"] == s]["Name"].values[0] + '.')
                failed.append(s)
                continue

            # Read the downloaded file
            df_record_s = pd.read_excel(file_list[0], skiprows=7)

            # Drop blank columns
            df_record_s = df_record_s.drop(columns=[df_record_s.columns[5],
                                                    df_record_s.columns[11]])

            # Get staff name
            name = df_record_s.iloc[0, 0]

            # Remove previous files
            try:
                for previous_file in glob.glob("reports/T_" + name + "*"):
                    os.remove(previous_file)
            except BaseException:
                pass

            # Save dataframe as CSV file
            file_name = "reports/T_" + name + "_" + staff_id + "_" + \
                get_timestamp(format="%Y%m%d") + ".csv"
            df_record_s.to_csv(file_name, index=False, encoding="utf_8_sig")

            # Go to previous page
            web.back()

            # Find input field for staff number
            staff_id_input = WebDriverWait(web, 10).until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        '//*[@id="ctl00_cphContent_txtTrainingStaffNo_' +
                        'txtStaffNo"]')))

            # Clear previous search
            staff_id_input.clear()

        except BaseException:
            # Append staff number to failed array
            failed.append(s)

            print("[" + get_timestamp() +
                  "] Failed to fetch training record for " +
                  df[df["Staff Number"] == s]["Name"].values[0] + '.')

            # Go to original page
            web.get(config["enquiry_training_link"])

    # Quit web
    web.quit()

    # Remove all training record files in Downloads folder
    file_list = glob.glob(user_path + "\\Downloads\\TrainResult*")

    if len(file_list) > 0:
        for file in file_list:
            os.remove(file)

    print("[" + get_timestamp() +
          "] Completed with " + str(len(failed)) + " failed case(s).")

    # Return failed cases
    return failed


if __name__ == "__main__":
    # Read configuration file
    config = read_configuration_file()

    # Fetch training record
    fetch_training_record(config)
