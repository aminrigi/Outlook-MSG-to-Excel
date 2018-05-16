# Outlook-MSG-to-Excel
Exporting a group of Outlook msg emails to Excel

This C# proram gets a textfile that contains the paths of all msg file that you want to export to Excel.
Both Outlook and Excel must be installed on the PC.

The columns of the Excel sheet are:

SenderName, smtpAddress, Subject, Recipients' smtpAddress, Recipients' names, sentTime, msgBody, msg file location, list of attachments


I used this program (in windows 10 environment) to export archieved company emails into excel and from there to R for further analysis.
The program works fine, but it doesn't end the excel process after execution. After the end of program you still can see excel in Task Manager. You can kill the process manually. I didn't fix it because I only used this program once.

