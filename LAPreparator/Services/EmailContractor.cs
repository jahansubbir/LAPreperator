using EmailService;
using LAPreparator.DataAccess;
using LAPreparator.Serivices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace LAPreparator.Services
{
 public   class EmailContractor
    {
        
        private readonly ExcelEmailAddressService emailAddressService;
        private readonly ITemplateReader templateReader;

        public EmailContractor(ExcelEmailAddressService emailAddressService, ITemplateReader templateReader)
        {
            
            this.emailAddressService = emailAddressService;
            this.templateReader = templateReader;
        }

        public List<Email> GetEmailModels()
        {
            string resourceDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Resources");
            string mailBody = templateReader.Read(Path.Combine(resourceDirectory, "template.msg")).Body;
            var emails = emailAddressService.GetEmails();
            foreach (Email email in emails)
            {
                email.Body = mailBody;
            }
            return emails;
        }
    }
}
