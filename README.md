# Outlook-MSG-to-Excel
Exporting a group of Outlook msg emails to Excel

This C# proram gets a textfile that contains the paths of all msg files that you want to export into Excel.
Both Outlook and Excel must be installed on the PC.

The columns of the Excel sheet are:

SenderName, smtpAddress, Subject, Recipients' smtpAddress, Recipients' names, sentTime, msgBody, msg file location, list of attachments


I used this program (in windows 10 environment) to export company's archieved emails into excel and from there to R for further analysis.
The program works fine, but it doesn't end the excel process after execution. After the program's execution, you still can see excel in Task Manager. You can kill the process manually. I didn't fix it because I only used (needed) this program once. If by any chance you fixed it, let me know :-)
