using System;
using System.Collections.Generic;
using System.Text;

namespace EmailService
{
    public class Email
    {
        public Email()
        {
            Reciepients = new List<string>();
        }
        
        public string Code { get; set; }
        public string Name { get; set; }

        public List<string> Reciepients { get; set; }
        public string To { get; set; }
        public string Cc { get; set; }

        public string Bcc { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string Attachment { get; set; }

    }
}
