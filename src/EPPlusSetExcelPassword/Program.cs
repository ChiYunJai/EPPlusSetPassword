using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSetExcelPassword
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            var file = new System.IO.FileInfo("sample.xlsx");
            if (file.Exists)
            {
                Console.WriteLine("警告，範例檔案已存在，自動刪除");
                file.Delete();
            }

            using (var excel = new ExcelPackage(file))
            {
                excel.Workbook.Worksheets.Add("sheet1");
                excel.Save("123321");
            }
        }
    }
}
