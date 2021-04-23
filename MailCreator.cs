using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Web;
using System.IO;
using System.Net;
using System.Net.Mime;

namespace CommentToAttendence
    {
     public class MailCreator 
        {
        public string defalultSender = "enter_your_email";
        public string emailSubject = "Attendance Sheet: ";
        public string email_body_header = "This is an automated Mail<br>_____________________________________________<br>";    
        public string email_body_footer = "<br><br>---------------------------------------------------<br><b>Developed by JITU<br>CSE Batch-15, PSTU<br>";
        public string soft_detail = "<br>Software Build Version: 1.0.0.0<br>Database: MS SQL SERVER<br>Framework: DOT NET 4.7.2<br>Copyright ©  2021<br>";


        /*
        //Create a temporary DataTable for user Emails.
                DataTable dtEmails = new DataTable();
                dtEmails.Columns.AddRange(new DataColumn[] {
                    new DataColumn("Name", typeof(string)),
                    new DataColumn("Email",typeof(string)) });
                dtEmails.Rows.Add("John Hammond", "john.hammond@test.com");
                dtEmails.Rows.Add("Mudassar Khan", "mudassar.khan@test.com");
                dtEmails.Rows.Add("Suzanne Mathews", "suzzane.mathews@test.com");
                dtEmails.Rows.Add("Robert Schidner", "robert.schidner@test.com");
 
                //string subject = "Welcome Email";
                //string body = "Hello {0},<br /><br />Welcome to ASPSnippets<br /><br />Thanks.";
 
                //Using Parallel Multi-Threading send multiple bulk email.
                Parallel.ForEach(dtEmails.AsEnumerable(), row =>
                {
                    //Send Email with Excel attachment.
                    SendEmail(row["Email"].ToString(), subject, string.Format(body, row["Name"]), bytes);
                }); 
        */

        
        public void SendEmail(string sender, string recipient, string subject, string body, byte[] attachment,string filenameDotExtension)
            {
            MailMessage mailMessage = new MailMessage(sender, recipient);
            mailMessage.Subject = subject;
            mailMessage.Body = body;

            //Add Byte array as Attachment.
            mailMessage.Attachments.Add(new Attachment(new MemoryStream(attachment), filenameDotExtension)); ///ekhane buffer null aschete

            mailMessage.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient();
            smtp.Host = "smtp.gmail.com";
            smtp.EnableSsl = true;
            NetworkCredential NetworkCred = new NetworkCredential();
            
            NetworkCred.UserName = "your_email";
            NetworkCred.Password = "password";
            
            smtp.UseDefaultCredentials = true;
            smtp.Credentials = NetworkCred;
            smtp.Port = 587;
            smtp.Send(mailMessage);
            }

        //MailMessage mailMessage = new MailMessage();
        //Attachment attachment = new Attachment(myMemorySteam, new ContentType(MediaTypeNames.Application.Octet));

        //attachment.ContentDisposition.FileName = "myFile.xlsx";
        //attachment.ContentDisposition.Size = attachment.Length;
       //mailMessage.Attachments.Add(attachment);
        


        public void SendEmail(string sender, string recipient, string subject, string body,  MemoryStream memoryStream, string filenameDotExtension)
            {
            MailMessage mailMessage = new MailMessage(sender, recipient);
            mailMessage.Subject = subject;
            mailMessage.Body = body;
            byte[] memory = memoryStream.ToArray();
            //Add Byte array as Attachment.
            //mailMessage.Attachments.Add(new Attachment(memoryStream, filenameDotExtension)); ///ekhane buffer null aschete
            
            Attachment attachment = new Attachment(memoryStream, new ContentType(MediaTypeNames.Application.Octet));
            attachment.ContentDisposition.FileName = filenameDotExtension;
            //attachment.ContentDisposition.Size = attachment.Length;
            mailMessage.Attachments.Add(attachment);

            mailMessage.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient();
            smtp.Host = "smtp.gmail.com";
            smtp.EnableSsl = true;
            NetworkCredential NetworkCred = new NetworkCredential();
            
            NetworkCred.UserName = "your_email";
            NetworkCred.Password = "password";
            
            smtp.UseDefaultCredentials = true;
            smtp.Credentials = NetworkCred;
            smtp.Port = 587;
            smtp.Send(mailMessage);
            }

        public void SendEmail(string sender, string recipient, string subject, string body)
            {
            MailMessage mailMessage = new MailMessage(sender, recipient);
            mailMessage.Subject = subject;
            mailMessage.Body = body;
             mailMessage.IsBodyHtml = true;

            SmtpClient smtp = new SmtpClient();
            smtp.Host = "smtp.gmail.com";
            smtp.EnableSsl = true;
            NetworkCredential NetworkCred = new NetworkCredential();
    
            NetworkCred.UserName = "your_email";
            NetworkCred.Password = "password";
            
            smtp.UseDefaultCredentials = true;
            smtp.Credentials = NetworkCred;
            smtp.Port = 587;
            smtp.Send(mailMessage);
            }

        }
    }
