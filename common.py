#!/usr/bin/env python3
"""Running some common functions."""

# Import libraries
from datetime import datetime
import logging
import os
import json


# Funtion to get timestamp
def get_timestamp(dt=None, format="%d/%m/%Y %H:%M:%S"):
    """Get timestamp."""
    # Get current timestamp if no datetime object received
    if dt is None:
        dt = datetime.now()

    # Generate and return timestamp in formatted string
    return dt.strftime(format)


# Function to calculate time difference
def get_time_difference(start_time, end_time):
    """Get time difference."""
    # Calculate time difference
    time_diff = end_time - start_time

    # Return time difference in seconds
    return time_diff.total_seconds()


# Function to print and record error message
def handle_error_message(error):
    """Handle error message."""
    # Print error message on console
    print("[" + get_timestamp() + "] " + str(error))

    # Create log to record error message
    logging.basicConfig(
        filename="logs/log_" + get_timestamp(format="%Y%m%d") + ".log",
        filemode="a")
    logging.error("[" + get_timestamp() + "] " + str(error))


# Function to read configuration file
def read_configuration_file():
    """Read configuration file."""
    # Open configuration file
    with open("config.json") as config_file:

        # Load configuration file
        config = json.load(config_file)

        # Set absolute file path with username
        path_prefix = "C:/Users/" + os.getlogin() + '/'

        for i in config:
            if "abs_path" in i:
                config[i] = path_prefix + config[i]

    # Close configuration file
    config_file.close()

    return config


if __name__ == "__main__":
    pass
