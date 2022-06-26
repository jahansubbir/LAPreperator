using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataExchanger
{
   public interface IExchanger
    { /// <summary>
      /// Copies excel data from one workbook to another
      /// </summary>
      /// <param name="sourceFile">source excel file path</param>
      /// <param name="sourceSheet">source work sheet</param>
      /// <param name="sourceRange">range to copy</param>
      /// <param name="destinationFile">destination file path</param>
      /// <param name="destinationSheet">destination work sheet</param>
      /// <param name="destinationRange">destionation range</param>
        void Copy(string sourceFile,string sourceSheet, string sourceRange, string destinationFile,string destinationSheet, string destinationRange);
    }
}
