using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace SendEmailSMTP
{
    public class SendEmailClass
    {
        public static string sendEmail(
            string fromAddress,
            string toAddress,
            string ccAddress,
            string bccAddress,
            string replyToAddress,
            string subject,
            bool isHtmlBody,
            string body,
            string severHost,
            int serverPort,
            bool secureConnection,
            string username,
            string password,
            string attachments)
        {

            try
            {
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                message.From = new MailAddress(fromAddress);
                if (toAddress.Length > 0)
                {
                    string[] toAdrs = toAddress.Split(',');
                    foreach (string toEmail in toAdrs)
                    {
                        message.To.Add(new MailAddress(toEmail));
                        LogToFile(" INFO : New Email added as TO Address : " + toEmail);
                    }

                }
                if (ccAddress.Length > 0)
                {
                    string[] CCId = ccAddress.Split(',');
                    foreach (string CCEmail in CCId)
                    {
                        message.CC.Add(new MailAddress(CCEmail)); //Adding Multiple CC email Id  
                        LogToFile(" INFO : New email address added to CC: " + CCEmail);
                    }

                }

                if (bccAddress.Length > 0)
                {
                    string[] BCCId = bccAddress.Split(',');
                    foreach (string BCCEmail in BCCId)
                    {
                        message.Bcc.Add(new MailAddress(BCCEmail)); //Adding Multiple BCC email Id  
                        LogToFile(" INFO : New email address added to bcc : " + BCCEmail);
                    }
                }

                if (replyToAddress.Length > 0)
                {
                    string[] replyToAdrs = replyToAddress.Split(',');
                    foreach (string replyToEmail in replyToAdrs)
                    {
                        message.ReplyToList.Add(new MailAddress(replyToEmail)); //Adding Multiple ReplyTo Email addresses
                        LogToFile(" INFO : New email address added to ReplyTo : " + replyToEmail);
                    }
                }

                message.Subject = subject;
                message.IsBodyHtml = isHtmlBody;
                message.Body = body;
                System.Net.Mail.Attachment attachment;
                if (attachments.Length > 0)
                {
                    string[] attachmentsArr = attachments.Split(',');
                    for (int i = 0; i < attachmentsArr.Length; i++)
                    {
                        attachment = new System.Net.Mail.Attachment(attachmentsArr[i]);
                        message.Attachments.Add(attachment);
                    }
                }
                else
                {
                    LogToFile(" INFO : No Attachments to be sent");
                }
                smtp.Port = serverPort;
                smtp.Host = severHost;
                LogToFile(" INFO : Configured smtp host : " + severHost);
                LogToFile(" INFO : Configured smtp port : " + serverPort);
                smtp.EnableSsl = secureConnection;
                smtp.UseDefaultCredentials = false;
                LogToFile(" INFO : SSL enable " + secureConnection);
                if (username.Length == 0)
                {
                    LogToFile(" INFO : Username is empty");
                    return "Username is empty";
                }
                if (password.Length == 0)
                {
                    LogToFile(" INFO : Password is empty");
                    return "Password is empty";
                }
                smtp.Credentials = new NetworkCredential(username, password);
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
                LogToFile(" INFO : Email sent successfully");
                return "Email sent successfully";
            }
            catch (Exception ex)
            {
                LogToFile(" ERROR: " + ex.Message);
                return ex.Message;
            }
        }

        private static void LogToFile(string line)
        {
            string logFile = "EmailSendSmtp.log";
            string logFileFolderPath = System.Environment.GetEnvironmentVariable("USERPROFILE") + "\\.logs\\";
            string logFilePath = logFileFolderPath + logFile;

            DateTime dateNow = DateTime.Now;
            //2022 - Mar - 23 Wed 09:09:50.623
            string now = dateNow.ToString("yyyy-MMM-dd ddd HH:mm:ss");
            // Create the folder where the .log file will be saved
            if (!Directory.Exists(logFileFolderPath))
            {
                DirectoryInfo di = Directory.CreateDirectory(logFileFolderPath);
                di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
            }
            // Create the .log file
            if (!File.Exists(logFilePath))
            {
                File.Create(logFilePath).Dispose();
            }
            // Log the line to the .log file
            using (TextWriter tw = new StreamWriter(logFilePath, true))
            {
                tw.WriteLine(now + ": " + line);
            }
        }
    }
}