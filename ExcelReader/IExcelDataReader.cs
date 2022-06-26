using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace ExcelReader
{
 public   interface IExcelDataReader
    {
        DataTable GetData(string fileName, string sheetName, string range);
        
    }
}
