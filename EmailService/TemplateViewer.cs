using System;
using System.Collections.Generic;
using System.Text;

namespace EmailService
{
  public static  class TemplateViewer
    {
        public static void DisplayMessage(string path)
        {
            if (path.EndsWith(".msg"))
            {
                var app = new Microsoft.Office.Interop.Outlook.Application();
                var mailItem = app.Session.OpenSharedItem(path) as Microsoft.Office.Interop.Outlook.MailItem;
                mailItem.Display();
            }
            else
            {
                throw new Exception("File Extension is not correct! extension should be .msg");
            }
        }
    }
}
