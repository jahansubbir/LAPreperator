using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EmailService;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace LAPreparator.DataAccess
{
    public class ExcelEmailAddressService 
    {
        IWorkbook workbook;
        ISheet sheet;
        public List<Email> GetEmails()
        {
            string path= $@"{Directory.GetCurrentDirectory()}\Resources\AddressBook.xlsx";
            List<Email> Emailes = new List<Email>();
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
            }
            sheet = workbook.GetSheet("Sheet1");
            //int rowNumber = 1;
            int endRow = sheet.LastRowNum;
            for (int i = 1; i <= endRow; i++)
            {
                IRow row = sheet.GetRow(i);
                Email Email = new Email();
                row.Cells.ForEach(a => a.SetCellType(CellType.String));
                Email.Code = row.GetCell(0)?.StringCellValue;
                Email.Name = row.GetCell(1)?.StringCellValue;
                Email.Reciepients = row.GetCell(2)?.StringCellValue.Split(new[] { ';', ',', ' ' },StringSplitOptions.RemoveEmptyEntries).ToList();
                Email.Cc = row.GetCell(3)?.StringCellValue;
                Email.Bcc = row.GetCell(4)?.StringCellValue;
                Email.To = row.GetCell(2)?.StringCellValue;
                Emailes.Add(Email);
            }
            workbook.Close();
            return Emailes;
        }
    }
}
