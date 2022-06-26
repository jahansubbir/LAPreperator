using ExcelDataExchanger;
using ExcelWriter;
using LAPreparator.DataAccess;
using LAPreparator.Models;
using LAPreparator.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace LAPreparator.BusinessLogic
{
    public class LaCreator : ILaCreator
    {
        private readonly IExcelDataWriter writer;
        private readonly IExchanger exchanger;
        private readonly EmailContractor emailContractor;
        private readonly EmailService.IEmailService emailService;

        public LaCreator(
            IExcelDataWriter writer,
             IExchanger exchanger,
             EmailContractor emailContractor,
             EmailService.IEmailService emailService
            )
        {
            this.writer = writer;
            this.exchanger = exchanger;
            this.emailContractor = emailContractor;
            this.emailService = emailService;
        }
        public List<LA> CreateLAs(IEnumerable<IGrouping<string, DataRow>> groupedData, string sourceFile, string sourceSheet, string sourceRange/*, string destinationFile, string destinationSheet, string destinationRange*/)
        {
            //Arrange Resources
            List<LA> list = new List<LA>();
            var addresses = emailContractor.GetEmailModels();

            var tempPath = Path.GetTempPath();
            var destinationFileName = Path.GetFileNameWithoutExtension(sourceFile);
            var destinationFile = $"{tempPath}{destinationFileName}.xlsx";
            foreach (var group in groupedData)
            {
                LA la = new LA();
                la.Shipper = group.Key;
                try
                {
                    var address = addresses.Find(a => a.Name?.Trim() == group.Key);

                    if (address != null)
                    {
                        
                        if (address.Reciepients?.Count > 0)
                        {



                            //  string destinationFileName = $"{directoryName}{group.Key}.xlsx";
                            writer.CreateExcel(destinationFile, sourceRange, group.CopyToDataTable());
                            exchanger.Copy(sourceFile, sourceSheet, sourceRange, destinationFile, "Sheet1", "A1:AZ");
                            address.Attachment += destinationFile;
                            address.Subject = Path.GetFileName(destinationFile);
                            bool emailSent = emailService.Send(address, out Exception ex);
                            File.Delete(destinationFile);
                            if (emailSent)
                            {
                                la.LaSent = true;
                                la.Status = "Sent";
                            }
                            else
                            {
                                la.Status = ex.Message;
                                la.LaSent = false;
                            }

                        }
                        else
                        {
                            la.Status = "Address Not Found";
                            la.LaSent = false;
                        }
                    }
                    else
                    {
                        la.Status = "Receipients Address is missing";
                        la.LaSent = false;

                    }
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message);
                    
                    la.LaSent = false;
                    la.Status = exception.Message;
                }
                list.Add(la);
            }
            return list;
        }


    }
}
