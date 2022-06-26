using System;
using System.Collections.Generic;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
namespace EmailService
{
    public interface IEmailService {



        bool Send(Email email,out Exception exception);
        
    }
}
