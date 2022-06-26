using ExcelReader;
using System;
using System.Data;
using System.Diagnostics;
using Xunit;

namespace LAPreperator.Test
{
    public class ExcelDataReaderTest
    {
        [Fact]
        public void GetData()
        {
            //Arrange
            IExcelDataReader excelDataReader = new ExcelDataReader();
            string fileName = @"C:\Users\msj046\Downloads\TESCO LA UK_2022-02-11_BOOKING NO_BAC0309771_X-PRESS NILWALA - 0YJ08S1MA.xls";
            string sheetName = "Sheet1";
            string range = "A15:Z";
            //Act
           var dt= excelDataReader.GetData(fileName, sheetName, range);

            //Assert
            foreach (var column in dt.Columns)
            {
                Debug.Write($"{column}\t");
            }
            foreach (DataRow row in dt.Rows)
            {
                foreach (string  column in row.ItemArray )
                {
                    Debug.WriteLine(column.ToString());
                }
            }
            Assert.True(true);
            //Console.ReadKey();
        }
    }
}
