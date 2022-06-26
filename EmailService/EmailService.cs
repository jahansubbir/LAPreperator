using System;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmailService
{
    public class EmailService : IEmailService
    {
        private Outlook.Application application = null;
        Outlook.MailItem mail = null;
        public bool Send(Email email, out Exception ex)
        {

            try
            {
                if (application is null)
                {
                    application = new Outlook.Application();
                }
                mail = application.CreateItem(Outlook.OlItemType.olMailItem);
                string reciepients = "";

                foreach (string recipient in email.Reciepients)
                {
                    var emailAddress = recipient.Replace("<", "").Replace(">", "").Trim();
                    bool isEmail = IsValidEmail(emailAddress);
                    if (isEmail)
                    {
                        reciepients += emailAddress;
                        reciepients += ";";
                    }

                }
                mail.To = email.To;
                if (!string.IsNullOrEmpty(email.Cc))
                {
                    mail.CC = email.Cc;
                }

                if (!string.IsNullOrEmpty(email.Bcc))
                {
                    mail.BCC = email.Bcc;
                }

                if (!string.IsNullOrEmpty(email.Attachment))
                {


                    mail.Attachments.Add(email.Attachment, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    //}
                }
                mail.HTMLBody = email.Body;
                mail.Subject = email.Subject;
                mail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                //mail.Display();
                mail.Send();
                ex = null;
                return true;
            }
            catch (Exception exception)
            {
                ex = exception;


                // throw exception;
                return false;
            }
        }
        bool IsValidEmail(string email)
        {
            var trimmedEmail = email.Trim();

            if (trimmedEmail.EndsWith("."))
            {
                return false; // suggested by @TK-421
            }
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == trimmedEmail;
            }
            catch
            {
                return false;
            }
        }
    }
}
