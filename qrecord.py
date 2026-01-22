#!/usr/bin/env python3
"""Process qualification and practice record."""

# Import libraries
from common import get_timestamp, read_configuration_file
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import numpy as np
import pandas as pd
import glob
import os


# Function to fetch qualification record
def fetch_qualification_record(config):
    """Fetch qualification record."""
    # Read staff list
    df = pd.read_csv(config["staff_list_path"], dtype="string")

    # Initialise webdriver
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    web = webdriver.Chrome(options=options)

    # Initialise an array to store all failed cases
    failed = []

    print("[" + get_timestamp() +
          "] Fetching staff qualification record...")
    # Iterate through all staff
    for s in df["Staff Number"]:
        staff_id = s

        # Start trial loop
        for trial in range(3):
            try:
                # Browse webpage
                web.get(config["enquiry_qualification_link"])

                # Find input field for staff number
                staff_id_input = WebDriverWait(web, 10).until(
                        EC.presence_of_element_located((
                            By.XPATH,
                            '//*[@id="ctl00_cphContent_txtEnquiryStaffNo_' +
                            'txtStaffNo"]')))

                # Fill in staff number
                staff_id_input.send_keys(staff_id)

                # Find "Search" button
                search_button = WebDriverWait(web, 10).until(
                        EC.presence_of_element_located((
                            By.XPATH,
                            '//*[@id="ctl00_cphContent_btnEnquiry"]')))

                # Click "Search" button
                search_button.click()

                # Find "Data Download" button
                WebDriverWait(web, 10).until(
                        EC.presence_of_element_located((
                            By.XPATH,
                            '//*[@id="ctl00_cphContent_btnExport"]')))

                # Get page source
                page_source = web.page_source
                soup = BeautifulSoup(page_source, 'lxml')

                # Find staff name, organisation unit and description
                name_and_id = soup.find(
                    "span",
                    id="ctl00_cphContent_MtrcMaster_ctl02_dgrdStaff_ctl02_" +
                    "Label8")
                name_and_id_string = name_and_id.text.lstrip().rstrip()
                name = name_and_id_string.replace(staff_id, "").rstrip()
                unit = soup.find(
                    "span",
                    id="ctl00_cphContent_MtrcMaster_ctl02_Label3").text
                unit_desc = soup.find(
                    "span",
                    id="ctl00_cphContent_MtrcMaster_ctl02_Label5").text

                # Find data table
                table = soup.find(
                    "table",
                    id="ctl00_cphContent_MtrcMaster_ctl02_dgrdStaff_ctl02_" +
                    "dgrdStaffQual")
                entries = table.find_all("td")

                # Initialise dataframe
                df_record = pd.DataFrame(
                    columns=[
                        "Qualification Code",
                        "Qualification",
                        "First Obtain",
                        "Last Refresh",
                        "Expiry",
                        "Due for Refresh/Examination",
                        "Last Practice/Attachment",
                        "Status",
                        "Note"
                    ]
                )

                # Iterate through all entries in the table
                row = []
                for i, e in enumerate(entries):
                    text = e.text.lstrip().rstrip()

                    # Handle qualification code and qualification
                    if i % 8 == 0:
                        if " " in text:
                            # Get qualification code
                            text_1 = text.split(" ")[0]
                            row.append(text_1)

                            # Get qualification
                            text_2 = text.replace(text_1, "").lstrip()
                            row.append(text_2)

                        else:
                            row.append("")
                            row.append(text)

                    else:
                        # Get date
                        row.append(text)

                    # Write row to dataframe
                    if i % 8 == 7:
                        df_record.loc[i // 8] = row
                        row.clear()

                # Add organisation unit and description columns
                df_record["Organization Unit"] = unit
                df_record["Organization Unit Desc"] = unit_desc

                # Remove previous files
                try:
                    for previous_file in glob.glob("reports/Q_" + name + '*'):
                        os.remove(previous_file)
                except BaseException:
                    pass

                # Save dataframe as CSV file
                file_name = "reports/Q_" + name + "_" + staff_id + "_" + \
                    get_timestamp(format="%Y%m%d") + ".csv"
                df_record.to_csv(file_name, index=False, encoding="utf-8-sig")

                # Exit trial loop if qualification record is fetched
                break

            except BaseException:

                print("[" + get_timestamp() +
                      "] Failed to fetch qualification record for " +
                      df[df["Staff Number"] == s]["Name"].values[0] +
                      " (Trial #" + str(trial + 1) + ").")

                # Exit trial loop if last trial
                if trial == 2:
                    # Append staff number to failed array
                    failed.append(s)

                    break

                # Continue trial loop if not last trial
                continue

    # Quit web
    web.quit()

    print("[" + get_timestamp() +
          "] Completed with " + str(len(failed)) + " failed case(s).")

    # Return failed cases
    return failed


# Function for fetching CQAS practice records
def fetch_practice_record(config, df):
    """Fetch CQAS practice records."""
    # Combine last refresh and first obtain dates
    df["Last Refresh_d"] = df["Last Refresh"].combine_first(df["First Obtain"])

    # Return the original dataframe if there is no practice to be fetched
    if df[df["Qualification Code"].isin(config["has_practice"])].empty is True:
        pass

    else:
        # Initialise webdriver
        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        web = webdriver.Chrome(options=options)

        print("[" + get_timestamp() + "] Fetching staff practice record...")

        # Get staff ID list
        df["Staff ID"] = df["Staff ID"].astype(str)
        staff_id_list = df["Staff ID"].unique()

        # Iterate through all staff
        for sid in staff_id_list:
            df_person = df[df["Staff ID"] == sid]

            # Continue the iteration if there is no practice to be fetched
            if df_person[df_person["Qualification Code"].isin(
                    config["has_practice"])
            ].empty is True:
                continue

            # Iterate through all qualifications
            for i, row in df_person.iterrows():

                # If there is practice requirement
                if row["Qualification Code"] in config["has_practice"]:

                    # Start trial loop
                    for trial in range(3):
                        try:
                            # Browse webpage
                            web.get(config["enquiry_practice_link"])

                            # Find "Clear" button
                            clear_button = WebDriverWait(web, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH,
                                     '//*[@id="ctl00_cphContent_btnClear_' +
                                     'Pract"]')))

                            # Click "Clear" button
                            clear_button.click()

                            # Find input field for staff number
                            WebDriverWait(web, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH,
                                     '//*[@id="ctl00_cphContent_txtSearch' +
                                     'Staff_Pract_txtStaffNo"]')))

                            # Fill in staff number
                            web.execute_script(
                                "document.getElementById(" +
                                "'ctl00_cphContent_txtSearchStaff_" +
                                "Pract_txtStaffNo'" +
                                ").setAttribute('value', '" +
                                sid + "')")

                            # Find "Search" button
                            search_button = WebDriverWait(web, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH,
                                     '//*[@id="ctl00_cphContent_btnSearch_' +
                                     'Pract"]')))

                            # Click "Search" button
                            search_button.click()

                            # Find "Cancel" button
                            cancel_button = WebDriverWait(web, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH,
                                     '//*[@id="ctl00_cphContent_btnBack"]')))

                            # Click "Cancel" button
                            cancel_button.click()

                            # Find input field for qualification code
                            WebDriverWait(web, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH,
                                     '//*[@id="ctl00_cphContent_txtQual_' +
                                     'Pract"]')))

                            # Fill in qualification code
                            web.execute_script(
                                "document.getElementById(" +
                                "'ctl00_cphContent_txtQual_Pract'" +
                                ").setAttribute('value', '" +
                                row["Qualification Code"] + "')")

                            # Find start date input field
                            WebDriverWait(web, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH,
                                     '//*[@id="ctl00_cphContent_' +
                                     'txtDateForSearchFrom_dateTextBox"]')))

                            # Input start date
                            web.execute_script(
                                "document.getElementById(" +
                                "'ctl00_cphContent_txtDateForSearchFrom_" +
                                "dateTextBox'" +
                                ").setAttribute('value', '" +
                                row["Last Refresh_d"] + "')")

                            # Find "Search" button
                            search_button = WebDriverWait(web, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH,
                                     '//*[@id="ctl00_cphContent_btnSearch_' +
                                     'Pract"]')))

                            # Click "Search" button
                            search_button.click()

                            # Find "Data Download" button
                            WebDriverWait(web, 10).until(
                                    EC.presence_of_element_located((
                                        By.XPATH,
                                        '//*[@id="ctl00_cphContent_' +
                                        'btnDownLoad"]')))

                            # Get page source
                            page_source = web.page_source
                            soup = BeautifulSoup(page_source, 'lxml')

                            # Find number of record found
                            df.at[i, "Last Practice/Attachment"] = soup.find(
                                "span", id="ctl00_cphContent_lblRecordCount"
                            ).text.split(":")[1]

                            # Find "Cancel" button
                            cancel_button = WebDriverWait(web, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH,
                                     '//*[@id="ctl00_cphContent_btnBack"]')))

                            # Click "Cancel" button
                            cancel_button.click()

                            # Exit trial loop if practice record is fetched
                            break

                        except BaseException:
                            print("[" + get_timestamp() +
                                  "] Failed to fetch practice record for " +
                                  df[df["Staff ID"] == sid][
                                      "Name"].values[0] +
                                  " (Trial #" + str(trial + 1) + ").")

                            # Exit trial loop if last trial
                            if trial == 2:
                                df.at[i, "Last Practice/Attachment"] = '?'
                                break

                            # Continue trial loop if not last trial
                            continue

                else:
                    pass

        # Quit web
        web.quit()

        print("[" + get_timestamp() + "] Completed.")

    # Replace NaN by '-'
    df["Last Refresh"] = df["Last Refresh"].replace(np.nan, '-')
    df["Last Practice/Attachment"] = df["Last Practice/Attachment"].replace(
            np.nan, '-')

    # Rename column name
    df = df.rename(columns={"Last Practice/Attachment": "Practice Done"})

    # Replace all dates by '-' in Practice Done column
    df.loc[df["Practice Done"].str.contains('/'), "Practice Done"] = '-'

    return df


if __name__ == "__main__":
    # Read configuration file
    config = read_configuration_file()

    # Fetch qualification record
    fetch_qualification_record(config)
