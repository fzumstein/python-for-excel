"""This module contains all functions that are either called from Excel
or manipulate Excel.
"""

import datetime as dt

from dateutil import tz
import requests
import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw

import database


# This is the part of the URL that is the same for every request
BASE_URL = "https://pypi.org/pypi"


def add_package():
    """ Adds a new package including the version history to the database.
    Triggers an update of the dropdown on the Tracker tab.
    """
    # Excel objects
    db_sheet = xw.Book.caller().sheets["Database"]
    package_name = db_sheet["new_package"].value
    feedback_cell = db_sheet["new_package"].offset(column_offset=1)

    # Clear feedback cell
    feedback_cell.clear_contents()

    # Check if the package exists on PyPI
    if not package_name:
        feedback_cell.value = "Error: Please provide a name!"
        return
    if requests.get(f"{BASE_URL}/{package_name}/json",
                    timeout=6).status_code != 200:
        feedback_cell.value = "Error: Package not found!"
        return

    # Insert the package name into the packages table
    error = database.store_package(package_name)
    db_sheet["new_package"].clear_contents()

    # Show any errors, otherwise kick off a database update and
    # refresh the dropdown so you can select the new package
    if error:
        feedback_cell.value = f"Error: {error}"
    else:
        feedback_cell.value = f"Added {package_name} successfully."
        update_database()
        refresh_dropdown()


def update_database():
    """ Deletes all records from the versions table, fetches all
    data again from PyPI and stores the versions again in the table.
    """
    # Excel objects
    sheet_db = xw.Book.caller().sheets["Database"]

    # Clear logs
    sheet_db["log"].expand().clear_contents()

    # Keeping things super simple: Delete all versions for all packages
    # and repopulate the package_versions table from scratch
    database.delete_versions()
    df_packages = database.get_packages()
    logs = []

    # Query the PyPI REST API
    for package_id, row in df_packages.iterrows():
        ret = requests.get(f"{BASE_URL}/{row['package_name']}/json",
                           timeout=6)
        if ret.status_code == 200:
            ret = ret.json()  # parse the JSON string into a dictionary
            logs.append(f"INFO: {row['package_name']} downloaded successfully")
        else:
            logs.append(f"ERROR: Could not download data for {row['package_name']}")
            continue

        # Instantiate a DataFrame by extracting data from the REST API response
        releases = []
        for version, files in ret["releases"].items():
            if ret["releases"][version]:  # ignore releases without info
                releases.append((files[0]["upload_time"], version, package_id))
        df_releases = pd.DataFrame(columns=["uploaded_at", "version_string", "package_id"],
                                   data=releases)
        df_releases["uploaded_at"] = pd.to_datetime(df_releases["uploaded_at"])
        df_releases = df_releases.sort_values("uploaded_at")
        database.store_versions(df_releases)
        logs.append(f"INFO: {row['package_name']} stored to database successfully")

    # Write out the last updated timestamp and logs
    sheet_db["updated_at"].value = (f"Last updated: "
                                    f"{dt.datetime.now(tz.UTC).isoformat()}")
    sheet_db["log"].options(transpose=True).value = logs


def show_history():
    """ Shows the latest release and plots the release history
    (number of releases per year)
    """
    # Excel objects
    book = xw.Book.caller()
    tracker_sheet = book.sheets["Tracker"]
    package_name = tracker_sheet["package_selection"].value
    feedback_cell = tracker_sheet["package_selection"].offset(column_offset=1)
    picture_cell = tracker_sheet["latest_release"].offset(row_offset=2)

    # Use the "seaborn" style for the Matplotlib plots produced by pandas
    plt.style.use("seaborn")

    # Check input
    if not package_name:
        feedback_cell.value = ("Error: Please select a package first! "
                               "You may first have to add one to the database.")
        return

    # Clear output cells and picture
    feedback_cell.clear_contents()
    tracker_sheet["latest_release"].clear_contents()
    if "releases_per_year" in tracker_sheet.pictures:
        tracker_sheet.pictures["releases_per_year"].delete()

    # Get all versions of the package from the database
    try:
        df_releases = database.get_versions(package_name)
    except Exception as e:
        feedback_cell.value = repr(e)
        return
    if df_releases.empty:
        feedback_cell.value = f"Error: Didn't find any releases for {package_name}"
        return

    # Calculate the number of releases per year and plot it
    df_releases_yearly = df_releases.resample("Y").count()
    df_releases_yearly.index = df_releases_yearly.index.year
    df_releases_yearly.index.name = "Years"
    df_releases_yearly = df_releases_yearly.rename(
        columns={"version_string": "Number of Releases"})
    ax = df_releases_yearly.plot.bar(
        title=f"Number of Releases per Year "
              f"({tracker_sheet['package_selection'].value})")

    # Write the results and plot to Excel
    version = df_releases.loc[df_releases.index.max(), "version_string"]
    tracker_sheet["latest_release"].value = (
        f"{version} ({df_releases.index.max():%B %d, %Y})")
    tracker_sheet.pictures.add(ax.get_figure(), name="releases_per_year",
                               top=picture_cell.top,
                               left=picture_cell.left)


def refresh_dropdown():
    """ Refreshes the dropdown on the Tracker tab with the content of
    the packages table.
    """
    # Excel objects
    book = xw.Book.caller()
    dropdown_sheet = book.sheets["Dropdown"]
    tracker_sheet = book.sheets["Tracker"]

    # Clear the current value in the dropdown
    tracker_sheet["package_selection"].clear_contents()

    # If the Excel table has non-empty rows, delete them before repopulating
    # it again with the values from the packages database table
    if dropdown_sheet["dropdown_content"].value:
        dropdown_sheet["dropdown_content"].delete()
    dropdown_sheet["dropdown_content"].options(
        header=False, index=False).value = database.get_packages()


if __name__ == "__main__":
    xw.Book("packagetracker.xlsm").set_mock_caller()
    add_package()
