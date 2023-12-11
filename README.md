# Auto Mailer For Windows Outlook Client
#### This script is built for automating sending outlook emails at scheduled interval using windows outlook client and task scheduler.

### How to Install
1. Use `git clone https://github.com/CraftyPythonDeveloper/Auto-Mailer.git` to clone the repository.
2. Open cmd or anaconda prompt and make sure python is accessible. To check type `python --version`
3. Run `pip install -r src/requirements.txt` to install required libraries.
4. Open **run.bat** in notepad and change PYTHONPATH to the actual location of your python interpreter.
5. To check PYTHONPATH, run below command in python interpreter.
   ````
   import sys
   sys.executable
   ````
6. Replace the \\\ with \ in the path and Change the same PYTHONPATH in **create_task.bat**

### How to use
1. Open **config.xlsx** and open sheet with name **messages**  
2. Add the group_id, make sure not to add any space or special characters in group_id, it's unique.
3. Add subject and attachment path, send_email_to, send_email_cc and send_email_bcc. To add multiple attachments, emails use **;** as a seperator.
4. Meaning of each column names are as below,
   1. subject: Subject of your email. (should match the subject of email you will save in draft)
   2. attachment_path: the location of your attachment separated by **;**.
   3. send_email_to: email addresses separated by **;** to send To.
   4. send_email_cc: email address separated by **;** to keep them in cc
   5. send_email_bb: email address separated by **;** to keep them in bcc
5. Open outlook and type the formatted message body and save it in draft. Make sure the subject in config and in outlook are same.
6. Open emails sheet in excel and add the email address of users with group_id.
7. Once done to test the script run `python automailer.py <group_id>` or double-click on `run.bat`
8. To schedule the email, eun `create_task.bat` and follow the instructions.
9. To delete a task, open `scheduled_tasks.xlsx` and get the task name.
10. Run **delete_scheduled_task.bat** and give task name to delete it.