using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelDataExchanger
{
    public class ExcelDataExchanger : IExchanger
    {
        IWorkbook sourceBook, destinationBook;
        ISheet sourceSheet, destinationSheet;


        public void Copy(string sourceFile, string sourceSheet, string sourceRange, string destinationFile, string destinationSheet, string destinationRange)
        {
            using (FileStream file = new FileStream(sourceFile, FileMode.Open))
            {
                sourceBook = sourceFile.EndsWith(".xls", StringComparison.InvariantCultureIgnoreCase) ? new HSSFWorkbook(file) : (IWorkbook)new XSSFWorkbook(file);

            }
            this.sourceSheet = sourceBook.GetSheet(sourceSheet);

            using (FileStream file = new FileStream(destinationFile, FileMode.Open))
            {
                this.destinationBook = new XSSFWorkbook(file);
            }
            this.destinationSheet = destinationBook.GetSheet(destinationSheet);
            CopyData(sourceRange, destinationRange);
            if (!destinationFile.EndsWith(".xlsx"))
            {
                destinationFile += ".xlsx";
            }
            using (var fileData = new FileStream(destinationFile, FileMode.OpenOrCreate))
            {

                destinationBook.Write(fileData);
                //  fileData.Close();

                //_workbook = new XSSFWorkbook(fileData);
            }
        }

        private void CopyData(string sourceRange, string destinationRange)
        {
            var array = sourceRange.Split(':');
            var numberString = Regex.Match(array[0], @"\d+").Value;
            var dataStartingIndex = Int32.Parse(numberString);
            IDataFormat dataFormatCustom = destinationBook.CreateDataFormat();

            var destinationStartingIndex = Int32.Parse(Regex.Match(destinationRange.Split(':')[0], @"\d+").Value);
            var destinationStart = destinationStartingIndex - 1;
            var endRow = dataStartingIndex - 1;
            int k = destinationStart;
            for (int i = 0; i < endRow; i++)
            {

                var sourceRow = sourceSheet.GetRow(i);

                var destinationRow = destinationSheet.CreateRow(k);
                if (sourceRow != null)
                {


                    for (int j = 0; j < sourceRow.LastCellNum; j++)
                    {
                        var sourceCell = sourceRow.GetCell(j);
                        if (sourceCell != null)
                        {


                            var destinationCell = destinationRow.CreateCell(j, sourceCell.CellType);
                            destinationCell.CellStyle = ConvertStyle(sourceCell);
                            //destinationCell.SetCellValue(sourceCell.c)
                            switch (sourceCell.CellType)
                            {
                                case CellType.Formula:
                                    destinationCell.CellFormula = sourceCell.CellFormula; break;
                                case CellType.Numeric:
                                    if (DateUtil.IsCellDateFormatted(sourceCell))
                                    {
                                        DateTime date = sourceCell.DateCellValue;
                                        ICellStyle style = sourceCell.CellStyle;
                                        // Excel uses lowercase m for month whereas .Net uses uppercase
                                        string format = style.GetDataFormatString().Replace('m', 'M').Replace(";", "").Replace("@", "");

                                        destinationCell.SetCellValue(date.ToString(format));
                                        //   destinationCell.CellStyle.DataFormat =  dataFormatCustom.GetFormat(format);
                                        break;
                                    }
                                    else
                                    {
                                        destinationCell.SetCellValue(sourceCell.NumericCellValue); break;
                                    }

                                case CellType.String:
                                    destinationCell.SetCellValue(sourceCell.StringCellValue); break;
                                case CellType.Blank:
                                    destinationCell.SetBlank(); break;

                            }
                        }

                    }
                }
                k++;
            }
        }
        private IFont ConvertFont(ICell cell)
        {
            var font = destinationBook.CreateFont();
            var sourceFont = cell.CellStyle.GetFont(sourceBook);
            font.Boldweight = sourceFont.Boldweight;
            font.Charset = sourceFont.Charset;
            font.Color = sourceFont.Color;
            font.FontHeightInPoints = sourceFont.FontHeightInPoints;
            font.FontName = sourceFont.FontName;
            font.IsBold = sourceFont.IsBold;
            return font;

        }
        private ICellStyle ConvertStyle(ICell cell)
        {

            ICellStyle cellStyle=cell.CellStyle;
            


            var convertedStyle = destinationBook.CreateCellStyle() as XSSFCellStyle;
            convertedStyle.DataFormat = cellStyle.DataFormat;
            convertedStyle.Alignment = cellStyle.Alignment;
            convertedStyle.VerticalAlignment = cellStyle.VerticalAlignment;

            convertedStyle.BorderLeft = cellStyle.BorderLeft;
            convertedStyle.BorderTop = cellStyle.BorderTop;
            convertedStyle.BorderRight = cellStyle.BorderRight;
            convertedStyle.BorderBottom = cellStyle.BorderBottom;

            convertedStyle.FillBackgroundColor = cellStyle.FillBackgroundColor;
            //convertedStyle.FillPattern = FillPattern.SolidForeground;
            //   if (cellStyle.FillForegroundColor != 0)

            //cellStyle.FillForegroundColor;
            convertedStyle.FillPattern = cellStyle.FillPattern;

            if (cell.CellStyle.GetType().Equals(typeof(XSSFCellStyle)))
            {
                //cellStyle = cell.CellStyle as XSSFCellStyle;
                convertedStyle.FillForegroundXSSFColor = (cellStyle as XSSFCellStyle).FillForegroundXSSFColor;
            }
            else
            {
                convertedStyle.FillForegroundColor = (cellStyle as HSSFCellStyle).FillForegroundColor;
                //cellStyle = cell.CellStyle as HSSFCellStyle;
            }
            //try
            //{
                
            //    convertedStyle.FillForegroundXSSFColor = (cellStyle as XSSFCellStyle).FillForegroundXSSFColor;
            //}
            //catch (Exception e)
            //{
            //    convertedStyle.FillForegroundColor = cellStyle.FillForegroundColor;
            //    //convertedStyle.FillPattern = cellStyle.FillPattern ;
            //}





            // cellStyle.FillPattern;
            //convertedStyle.WrapText = cellStyle.WrapText;

            convertedStyle.SetFont(ConvertFont(cell));
            //convertedStyle.CloneStyleFrom(cellStyle);
            return convertedStyle;
        }
    }
}
