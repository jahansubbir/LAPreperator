using LAPreparator.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace LAPreparator.Utilities
{
   public static class ReportGenerator
    {
        public static void CreateReport(List<LA> las)
        {
            string fileName = Path.Combine(Directory.GetCurrentDirectory(),$"LA{DateTime.Now.ToString("yyyyMMddhhmm")}.xls");

            StreamWriter writer = new StreamWriter(fileName);
            writer.WriteLine("SHIPPER\tSTATUS\tLA SENT");
            foreach (LA item in las)
            {
                writer.WriteLine($"{item.Shipper}\t{item.Status}\t{item.LaSent}");
            }
            writer.Flush();
            writer.Close();
            ProcessStartInfo pInfo = new ProcessStartInfo
            {
                FileName = fileName,
                UseShellExecute = true
            };
            Process.Start(pInfo);
            
        }
    }
}
