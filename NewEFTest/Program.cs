using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace NewEFTest
{
    internal class Program
    {

        public class ExcelManager
        {
            Application app;
            Workbook wb;
            Worksheet ws;
            Range rg;
            string _pathAndName;

            public ExcelManager(string pathAndName)
            {
                _pathAndName = pathAndName;
                app = new Application();
                wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                ws = (Worksheet)wb.Worksheets[1];
                ws.Name = "Data";
                rg = ws.UsedRange;
            }

            public void FillCell(int row, int column, string value)
            {
                ws.Cells[row, column] = value;
            }

            public void Save()
            {
                wb.SaveAs(_pathAndName);
                wb.Close();
            }
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Путь, по которому должен быть сохранён файл:");
            Console.WriteLine(Path.Combine(Environment.CurrentDirectory));
            ExcelManager em = new ExcelManager(Path.Combine(Environment.CurrentDirectory));
            em.FillCell(1, 2, "Man");
            em.Save();

            Console.ReadLine();
        }
    }
}
