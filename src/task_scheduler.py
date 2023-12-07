__author__ = "Amit Yadav"
__email__ = "amityadav4664@gmail.com"

import sys
from os.path import dirname, realpath, join
import subprocess
import pandas as pd

messages = {"TASK_TYPE": "Choose scheduling option: \n1. Daily \n2. Weekly \n3. Monthly \n4. Every X Minutes.\n(1,2,"
                         "3,4): ",
            "INVALID_INPUT": "The Input is not valid",
            "INPUT_TIME": "Enter the time for daily schedule (HH:mm): ",
            "INPUT_GROUP_ID": "Enter the group_id (get it from config.xlsx): ",
            "INPUT_DAY": "Enter the day for weekly schedule (MON, TUE, WED, etc.): ",
            "INPUT_DATE": "Enter the day of the month for monthly schedule (1-31): ",
            "INPUT_MINUTE": "Enter the number of minutes for task scheduling: ",
            "TEST_TASK": "Do you want to test the scheduled task? (Y/N): ",
            "EXIT": "press any key to exit.."}

PYTHON_PATH = sys.executable
SCRIPT_PATH = join(dirname(realpath(__file__)), "auto_mailer.py")
TASK_EXCEL_PATH = join(dirname(realpath(__file__)), "scheduled_tasks.xlsx")

if __name__ == "__main__":
    task_df = pd.read_excel(TASK_EXCEL_PATH, sheet_name="tasks")
    task_df.set_index("Id", inplace=True)
    group_id = input(messages["INPUT_GROUP_ID"])
    if " " in group_id:
        print("group_id should not contain any spaces. it should be a single word")
    task_type = int(input(messages["TASK_TYPE"]))
    task_name = f"AutoMailerTask_{group_id}"

    if task_type == 1:
        time = input(messages["INPUT_TIME"])
        cmd = (f'schtasks /create /tn "{task_name}" /tr "{PYTHON_PATH} {SCRIPT_PATH} {group_id}" /sc '
               f'daily /st {time} /f')
        details = {"group_id": group_id, "time": time, "frequency": "Daily"}
    elif task_type == 2:
        time = input(messages["INPUT_TIME"])
        day = input(messages["INPUT_DAY"])
        cmd = (f'schtasks /create /tn "{task_name} /tr "{PYTHON_PATH} {SCRIPT_PATH} {group_id}" /sc '
               f'weekly /d {day} /st {time} /f')
        details = {"group_id": group_id, "time": time, "day": day, "frequency": "Weekly"}
    elif task_type == 3:
        time = input(messages["INPUT_TIME"])
        date = input(messages["INPUT_DAY"])
        cmd = (f'schtasks /create /tn "{task_name}" /tr "{PYTHON_PATH} {SCRIPT_PATH} {group_id}" /sc '
               f'monthly /d {date} /st {time} /f')
        details = {"group_id": group_id, "time": time, "date": date, "frequency": "Monthly"}
    elif task_type == 4:
        minute = input(messages["INPUT_MINUTE"])
        cmd = (f'schtasks /create /tn "{task_name}" /tr "{PYTHON_PATH} {SCRIPT_PATH} {group_id}" /sc '
               f'minute /mo {minute} /f')
        details = {"group_id": group_id, "frequency": f"Every {minute}"}
    else:
        print(messages["INVALID_INPUT"])
        input(messages["EXIT"])
        sys.exit()

    return_code = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    if return_code.returncode == 0:
        print(f"Task with name {task_name} scheduled successfully.")
        test_task = input(messages["TEST_TASK"])
        if test_task.lower() == "y":
            print(f"Running {task_name} task now..")
            subprocess.run(f'schtasks /run /tn "{task_name}"')
            print(f"ran {task_name} task..")
        data = {"task_name": task_name, "details": details}
        task_df.loc[task_df.shape[0]] = data
        task_df.to_excel(TASK_EXCEL_PATH, sheet_name="tasks")
    else:
        print("Error while scheduling the task..")
        print(return_code.stderr)

    input(messages["EXIT"])
    sys.exit()
