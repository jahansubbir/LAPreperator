using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelReader
{
    public class ExcelDataReader : IExcelDataReader
    {
        public ExcelDataReader()
        {

        }
        private IWorkbook workBook;
        private ISheet sheet;

        public DataTable GetData(string fileName, string sheetName, string range)
        {
            try
            {
                using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    workBook = fileName.EndsWith(".xls",StringComparison.InvariantCultureIgnoreCase) ? new HSSFWorkbook(file) : (IWorkbook)new XSSFWorkbook(file);
                }

                sheet = workBook.GetSheet(sheetName);
                var dataTable = new DataTable(sheet.SheetName);

                int startingRowIndex = 0;
                dataTable = GetColumns(dataTable, range, out startingRowIndex);
                dataTable = GetRows(dataTable, startingRowIndex);

                return dataTable;
            }
            catch (IOException ioException)
            {
                return null;
            }
            catch (Exception exception)
            {
                throw exception;
            }
          
            //  GetColumnNames(range);
        }

        private DataTable GetRows(DataTable dataTable, int startingRow)
        {
            for (int i = startingRow; i <= sheet.LastRowNum; i++)
            {
                var sheetRow = sheet.GetRow(i);
                if (sheetRow != null)
                {
                    var dtRow = dataTable.NewRow();
                    dtRow.ItemArray = dataTable.Columns
                        .Cast<DataColumn>()
                        .Select(c => sheetRow.GetCell(c.Ordinal, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString())
                        .ToArray();
                    dataTable.Rows.Add(dtRow);
                }
            }
            
            return dataTable;
        }

        private DataTable GetColumns(DataTable dt, string range, out int startingRowIndex)
        {
            //init list of columnName
            List<string> columns = new List<string>();
            //get rowIndex of data
            var array = range.Split(':');
            var numberString = Regex.Match(array[0], @"\d+").Value;


            startingRowIndex = Int32.Parse(numberString);
            var headerRowIndex = startingRowIndex - 1;
            var lastRowIndex = sheet.PhysicalNumberOfRows;

            IRow headerRow = sheet.GetRow(headerRowIndex);
            if (headerRow != null)
            {
                foreach (var item in headerRow)
                {
                    dt.Columns.Add($"{item?.ToString()}");
                }
            }
            return dt;
            //var lastCellNumber = headerRow.LastCellNum;
            //for (int i = 0; i < lastCellNumber; i++)
            //{
            //    ICell cell=headerRow.GetCell(i);
            //    cell.SetCellType( CellType.String);
            //    columns.Add(cell.StringCellValue);
            //}

        }
    }

}

