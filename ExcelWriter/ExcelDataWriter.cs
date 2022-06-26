using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelWriter
{
    public class ExcelDataWriter : IExcelDataWriter
    {
        IWorkbook workBook;
        ISheet sheet;
        IFont font;
        ICellStyle textCellStyle;
        ICellStyle numberCellStyle;
        ICellStyle dateCellStyle;
        public void CreateExcel(string fileName, string range, DataTable data)
        {
            workBook = new XSSFWorkbook();
            this.sheet = workBook.CreateSheet("Sheet1");
            CreateHeader(range, data, out int dataStartingIndex);
            CreateData(dataStartingIndex, data);

            using (FileStream file = new FileStream(fileName, FileMode.OpenOrCreate))
            {
                workBook.Write(file);
                file.Close();

            }
        }

        private void CreateData(int dataStartingIndex, DataTable data)
        {
            var style = SetDataCellStyle();
            int rowIndex = dataStartingIndex;
            foreach (DataRow dataRow in data.Rows)
            {
                IRow row = sheet.CreateRow(rowIndex);
                int i = 0;
                foreach (string column in dataRow.ItemArray)
                {
                    var cell = row.CreateCell(i);
                    cell.SetCellValue(column.ToString());
                    cell.CellStyle = style;
                    i++;
                }
                rowIndex++;

            }
        }

        private void CreateHeader(string range, DataTable dataTable, out int dataStartingIndex)
        {
            var style = SetHeaderCellStyle();
            if (!range.Contains(':'))
            {
                throw new FormatException("Bad Range Format. Range should be like A1:Z10");
            }
            else
            {
                var array = range.Split(':');
                var numberString = Regex.Match(array[0], @"\d+").Value;


                dataStartingIndex = Int32.Parse(numberString);
                var headerRowIndex = dataStartingIndex - 1;
                var row = sheet.CreateRow(headerRowIndex);
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    var cell = row.CreateCell(i, CellType.String);
                    cell.CellStyle = style;
                    cell.SetCellValue(dataTable.Columns[i].ColumnName.ToString());

                    //  cell.CellStyle = style;
                }

            }
        }

        private ICellStyle SetHeaderCellStyle()
        {
            ICellStyle style = workBook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.BorderBottom = BorderStyle.Medium;
            style.BorderLeft = BorderStyle.Medium;
            style.BorderRight = BorderStyle.Medium;
            style.BorderTop = BorderStyle.Medium;
            style.SetFont(GetFont("Calibri", 12, IndexedColors.DarkBlue, bold: true));
            style.WrapText = false;
            return style;
        }
        private ICellStyle SetDataCellStyle()
        {
            ICellStyle style = workBook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.BorderBottom = BorderStyle.Medium;
            style.BorderLeft = BorderStyle.Medium;
            style.BorderRight = BorderStyle.Medium;
            style.BorderTop = BorderStyle.Medium;
            style.SetFont(GetFont("Calibri", 12, IndexedColors.DarkBlue, bold: false));
            return style;
        }
        private IFont GetFont(string fontName, int size, IndexedColors color, bool bold)
        {
            IFont font = workBook.CreateFont();
            font.FontName = fontName;
            font.Color = color.Index;
            font.FontHeightInPoints = (short)size;
            font.IsBold = bold;
            return font;
        }
    }
}
