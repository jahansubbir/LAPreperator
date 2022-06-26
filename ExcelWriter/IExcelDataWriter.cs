using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace ExcelWriter
{
    public interface IExcelDataWriter
    {
        void CreateExcel(string fileName,string range, DataTable data);
        
    }
}
