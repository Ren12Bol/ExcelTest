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
            //Создание экземпляра программы
            Application app;
            //Создание экземпляра книги
            Workbook wb;
            //Создание экземпляра листа
            Worksheet ws;
            string _pathAndName;


            //Это метод для непосредственно создания файла
            public ExcelManager(string pathAndName)
            {
                _pathAndName = pathAndName;
                app = new Application();

                /* Workbooks.Add cоздаёт книгу
                 * XlWBATemplate.xlWBATWorksheet это тип книги
                 * Не обязателен, но лучше пусть будет */
                wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                
                //Открывает для программы первый лист книги
                ws = (Worksheet)wb.Worksheets[1];
                ws.Name = "Data";
            }

            // Метод для заполнения листа
            public void FillCell(int row, int column, string value)
            {
                ws.Cells[row, column] = value;
            }

            // Метод для сохранения файла
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

            //В скобках страшное это текущий путь
            ExcelManager em = new ExcelManager(Path.Combine(Environment.CurrentDirectory) + @"\NameOfTheFile.xlsx");
            em.FillCell(1, 2, "Man");
            em.Save();

            Console.ReadLine();
        }
    }
}
