from datetime import datetime, timedelta
import pandas as pd


def analyze_consecutive_days(sheet):
    employee_data = {}

    for index, row in sheet.iterrows():
        employee_name = row["Employee Name"]
        time_in = row["Time"].date()
        time_out = row["Time Out"].date()
        position_status = row["Position Status"]

        if employee_name not in employee_data:
            employee_data[employee_name] = {
                "consecutive_days_count": 0,
                "last_date": None,
            }

        if employee_data[employee_name]["last_date"] is not None:
            if employee_data[employee_name]["last_date"] + timedelta(days=1) == time_in:
                employee_data[employee_name]["consecutive_days_count"] += 1
            elif employee_data[employee_name]["last_date"] == time_in:
                pass
            else:
                employee_data[employee_name]["consecutive_days_count"] = 1
        else:
            employee_data[employee_name]["consecutive_days_count"] = 1

        employee_data[employee_name]["last_date"] = time_in

        if employee_data[employee_name]["consecutive_days_count"] == 7:
            print(f"{employee_name} : {position_status}")


def check(time_str1, time_str2):
    time1 = datetime.strptime(time_str1, "%H:%M:%S")
    time2 = datetime.strptime(time_str2, "%H:%M:%S")
    time_diff = str(time2 - time1).split(",")
    if len(time_diff) == 2:
        time_diff = time_diff[1].split(":")[0]
    else:
        time_diff = time_diff[0].split(":")[0]

    # print(time_str1,time_str2,time_diff)
    return int(time_diff) >= 14


def workedForMoreThan14HoursInSingleShift(df):
    for index, row in df.iterrows():
        employee_name = row["Employee Name"]
        position_status = row["Position Status"]
        time_in = row["Time"]
        time_out = row["Time Out"]
        if time_in is not pd.NaT and time_out is not pd.NaT:
            time_in = str(time_in.time())
            time_out = str(time_out.time())
            # print(time_in,time_out)
            if check(time_in, time_out):
                print(f"{employee_name} : {position_status}")


def time_diff(time1, time2):
    time1 = datetime.strptime(time1, "%H:%M:%S")
    time2 = datetime.strptime(time2, "%H:%M:%S")
    time_diff = str(time2 - time1).split(",")
    if len(time_diff) == 2:
        time_diff = str(time_diff[1])
    else:
        time_diff = str(time_diff[0])
    return time_diff


def haveDiffOf1To10Hours(df):
    documents = {}
    for index, row in df.iterrows():
        employee_name = row["Employee Name"]
        position_status = row["Position Status"]
        time_in = row["Time"]
        time_out = row["Time Out"]

        if employee_name not in documents:
            documents[employee_name] = {"date": None, "last_time": None, "diff": 0}

        if time_in is not pd.NaT and time_out is not pd.NaT:
            date = time_in.date()
            if (
                documents[employee_name]["date"] is None
                or documents[employee_name]["date"] != date
            ):
                time_out_only = time_out.time()
                documents[employee_name] = {
                    "date": date,
                    "last_time": time_out_only,
                    "diff": 0,
                }
            else:
                diff = time_diff(
                    str(documents[employee_name]["last_time"]), str(time_in.time())
                )
                diff = int(diff.split(":")[0])
                documents[employee_name] = {"date": None, "last_time": None, "diff": 0}
                if diff >= 1 and diff < 10:
                    print(f"{employee_name} : {position_status}")


# Example usage:
file_path = "Assignment_Timecard.xlsx"
df = pd.read_excel((file_path))
print("---------- First ------------")
analyze_consecutive_days(df)
print()
print("---------- Second -----------")
workedForMoreThan14HoursInSingleShift(df)
print()
print("--------- Third ---------")
haveDiffOf1To10Hours(df)
