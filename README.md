# Auto Mailer For Windows Outlook Client
#### This script is built for automating sending outlook emails at scheduled interval using windows outlook client and task scheduler.

### How to Install
1. Use `git clone https://github.com/CraftyPythonDeveloper/Auto-Mailer.git` to clone the repository.
2. Open cmd or anaconda prompt and make sure python is accessible. To check type `python --version`
3. Run `pip install -r src/requirements.txt` to install required libraries.
3. Open **run.bat** in notepad and change PYTHONPATH to the actual location of your python interpreter.
4. To check PYTHONPATH, run below command in python interpreter.
   ````
   import sys
   sys.executabe
   ````
5. Change the same PYTHONPATH in **create_task.bat**

### How to use
1. Open **config.xlsx** and open sheet with name **messages**  
2. Add the group_id, make sure not to add any space or special characters in group_id, it's unique.
3. Add subject and attachment path.
4. Open outlook and type the formatted message body and save it in draft. Make sure the subject in config and in outlook are same.
5. Open emails sheet in excel and add the email address of users with group_id.
6. Once done to test the script run `python automailer.py <group_id>` or double-click on `run.bat`
7. To schedule the email, eun `create_task.bat` and follow the instructions.
8. To delete a task, open `scheduled_tasks.xlsx` and get the task name.
9. Run **delete_scheduled_task.bat** and give task name to delete it.