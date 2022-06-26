using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LAPreparator.Serivices
{
    public class MessageModel
    {
        private string htmlBody;
        private List<object> attachment;

        public MessageModel(string subject, string htmlBody, List<object> attachment)
        {
            this.Subject = subject;
            this.Body = htmlBody;
            this.Attachments = new List<object>();
            //List<string> attachments = new List<string>();
            foreach (var item in attachment)
            {
                var a = item as MsgReader.Outlook.Storage.Attachment;
                this.Attachments.Add(GetFile(a));
            }
            //this.Attachments = attachments ;
        }
        public string UniqueKey { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public List<object> Attachments { get; set; }
        private string GetFile(MsgReader.Outlook.Storage.Attachment attachment)
        {
            string tempFolder = Path.GetTempPath();
            string fileName = $@"{tempFolder}{attachment.FileName}";
            
            try { 
                File.WriteAllBytes($"{fileName}", attachment.Data); 
            }catch(Exception e)
            {

            }
            return fileName;

        }
    }
}
