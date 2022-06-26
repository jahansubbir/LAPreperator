using LAPreparator.Serivices;
using System;
using System.Collections.Generic;
using System.IO;


namespace DataDrivenCustomMailer.MessageReaders
{
    

public class MsgTemplateReader : ITemplateReader
    {
        public MessageModel Read(string fileName)
        {
            MessageModel messageModel = null;
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var msg = new MsgReader.Outlook.Storage.Message(fileName))
                {
                    var subject = msg.Subject;
                    var htmlBody = msg.BodyHtml;
                    var attachment = msg.Attachments;
                    messageModel = new MessageModel(subject, htmlBody, attachment);
                }
            }
            catch (Exception e)
            {

            }

            return messageModel;
        }
    }
}
