using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;

namespace MsgParser
{
    class Program
    {
        static void Main(string[] args)
        {
            Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            Outlook._NameSpace nameSpace = app.GetNamespace("MAPI");
            nameSpace.Logon(null, null, false, false);
            Outlook.Folder folder = (Outlook.Folder)app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts);

            Outlook.MailItem msg;
            
            //Making an Excel file
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

			//Passing the files with address of msg file to StreamReader
            System.IO.StreamReader file = new System.IO.StreamReader(@"list_of_msg_files.txt");

            const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

            int count = 1; //In excel count starts from 1
            string filePath, smtpAddress="";
			
			//reading name of each msg file from 
            while ((filePath = file.ReadLine()) != null)
            {

                try
                {
                    check++;
                    msg = (Outlook.MailItem)app.Session.OpenSharedItem(filePath);
                }
                catch (Exception ex)
                {
                    System.Console.WriteLine(filePath);
                    if (check > 11093) break;
                    else continue;
                }

                
                try
                {
                    //Fiding Sender Information
                    Outlook._MailItem temp = ((Outlook._MailItem)msg).Reply();

                    Outlook.Recipient sender = temp.Recipients[1];
                    Outlook.PropertyAccessor xx = sender.PropertyAccessor;


                    try //For some companies, when an employee leaves, the local email address is deleted
                    {
                        smtpAddress = xx.GetProperty(PR_SMTP_ADDRESS).ToString();
                    }
                    catch (Exception ex)
                    {
                        smtpAddress = "NAN";
                    }

					temp.Delete();

                    xlWorkSheet.Cells[count, 1] = (sender.Name == null) ? "" : sender.Name;

                }
                catch(Exception e) //ignore!!
                {
                    xlWorkSheet.Cells[count, 1] = (msg.SenderName == null) ? "" : msg.SenderName;
                }

                
                xlWorkSheet.Cells[count, 2] = (smtpAddress == null) ? "" : smtpAddress;

                //subject
                xlWorkSheet.Cells[count, 3] = (msg.Subject == null) ? "" : msg.Subject;

                //Recipients
                smtpAddress = "";
                string names = "";
                Outlook.Recipients recips = msg.Recipients;
                
                foreach (Outlook.Recipient recip in recips)
                {
                    Outlook.PropertyAccessor pa = recip.PropertyAccessor;

                    string tmp;
                    try // if email address is not there
                    {
                         tmp = pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                    }
                    catch(Exception ex)
                    {
                        tmp = recip.Name;
                    }

                    smtpAddress = smtpAddress + "," + tmp;
                    names = names +","+ recip.Name;

                                        
                }

				
                xlWorkSheet.Cells[count, 4] = (smtpAddress == null) ? "" : smtpAddress;
                xlWorkSheet.Cells[count, 5] = (names == null) ? "" : names;

				//In case you need CreationTime
                //xlWorkSheet.Cells[count, 6] = (msg.CreationTime == null) ? "" : msg.CreationTime.ToShortDateString();
                //xlWorkSheet.Cells[count, 7] = (msg.ReceivedTime == null) ? "" : msg.ReceivedTime.ToShortDateString();

                xlWorkSheet.Cells[count, 6] = (msg.SentOn == null) ? "" : msg.SentOn.Day + "-" + msg.SentOn.Month + "-" + msg.SentOn.Year;
                xlWorkSheet.Cells[count, 7] = (msg.Body == null) ? "" : msg.Body;
                xlWorkSheet.Cells[count, 8] = filePath;


                string attached = "";
                for(int i=0; i<msg.Attachments.Count; i++)
                {
                    attached += "," + msg.Attachments[i+1].FileName;
                }


                xlWorkSheet.Cells[count, 9] = attached;

                count++;

            }


            file.Close();

            xlWorkBook.SaveAs("Emails.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            System.Console.ReadLine();
        }


        
        
    }
}
