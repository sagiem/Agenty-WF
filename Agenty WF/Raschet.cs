using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Agenty_WF
{
    class Raschet
    {
        public string file;
        string b, c, e, dateYR, aktnYR;
        decimal d;
        decimal sum = 0;
        int z = 14;
        int t = 1;

        List<ExcelOpen> exp = new List<ExcelOpen>();


        public Raschet(string file, string dateYR, string aktnYR)
        {
            this.file = file;
            this.dateYR = dateYR;
            this.aktnYR = aktnYR;
        }
        public void Exelreader()
        {
            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(file); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
            b = ObjWorkSheet.Cells[z, 3].Text.ToString();
            while (b != "")
            {
                b = ObjWorkSheet.Cells[z, 3].Text.ToString();//считываем текст в строку
                c = ObjWorkSheet.Cells[z, 4].Text.ToString();
                if (b.Length == 10)
                {
                    e = b;
                    d = decimal.Parse(c.Replace(" ", ""));
                    ExcelOpen excelOpen = new ExcelOpen(t, e, d);
                    exp.Add(excelOpen);
                    t++;
                    sum = sum + d;
                }
                z++;
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за  собой

        }

        public void ExelOtchet()
        {

            //Создаём новый Word.Application
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //Загружаем документ
            Microsoft.Office.Interop.Word.Document doc = null;

            object fileName = @"C:\Users\kashinmv\Desktop\прога\Agenty WF\files\otchet.docx";
            object falseValue = false;
            object trueValue = true;
            object missing = Type.Missing;

            doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);

            //Указываем таблицу в которую будем помещать данные (таблица должна существовать в шаблоне документа!)
            Microsoft.Office.Interop.Word.Table tbl = app.ActiveDocument.Tables[1];

            //Заполняем в таблицу - 10 записей.
            int i;
            for (i = 1; i <= exp.Count(); i++)
            {
                tbl.Rows.Add(ref missing);//Добавляем в таблицу строку.
                                          //Обычно саздаю только строку с заголовками и одну пустую для данных.
                tbl.Rows[i + 1].Cells[1].Range.Text = ((exp[i - 1]).a).ToString();
                tbl.Rows[i + 1].Cells[2].Range.Text = ((exp[i - 1]).c).ToString();
                tbl.Rows[i + 1].Cells[3].Range.Text = ((exp[i - 1]).d).ToString();

            }

            tbl.Rows.Add(ref missing);//Добавляем в таблицу строку.
                                      //Обычно саздаю только строку с заголовками и одну пустую для данных.
            tbl.Rows[i + 1].Cells[1].Range.Text = "Итого";
            tbl.Rows[i + 1].Cells[2].Range.Text = (i - 1).ToString();
            tbl.Rows[i + 1].Cells[3].Range.Text = (sum).ToString();


            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.

            object findText = "[data]";
            object replaceWith = dateYR;
            object replace = 2;




            app.Selection.Find.Execute(ref findText, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
            ref replace, ref missing, ref missing, ref missing, ref missing);

            object findText1 = "[ссылка акт]";
            object replaceWith1 = "№ "+aktnYR + " от " + dateYR;
            object replace1 = 2;


            app.Selection.Find.Execute(ref findText1, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith1,
            ref replace1, ref missing, ref missing, ref missing, ref missing);



            //Открываем документ для просмотра.
            app.Visible = true;
            //app.Quit(); // выйти из word
            GC.Collect(); // убрать за  собой


            //// Создаём экземпляр нашего приложения
            //Excel.Application excelApp = new Excel.Application();
            //// Создаём экземпляр рабочий книги Excel
            //Excel.Workbook workBook;
            //// Создаём экземпляр листа Excel
            //Excel.Worksheet workSheet;

            //workBook = excelApp.Workbooks.Add();
            //workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);

            //// Заполняем первую строку числами от 1 до 10
            //for (int j = 1; j <= exp.Count(); j++)
            //{
            //    workSheet.Cells[j, 1] = (exp[j-1]).a;
            //    workSheet.Cells[j, 2] = (exp[j - 1]).c;
            //    workSheet.Cells[j, 3] = (exp[j - 1]).d;
            //}

            //// Открываем созданный excel-файл
            //excelApp.Visible = true;
            //excelApp.UserControl = true;



        }

        public void ExelAkt()
        {
            //Создаём новый Word.Application
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //Загружаем документ
            Microsoft.Office.Interop.Word.Document doc = null;

            object fileName = @"C:\Users\kashinmv\Desktop\прога\Agenty WF\files\akt.docx";
            object falseValue = false;
            object trueValue = true;
            object missing = Type.Missing;

            doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);

            //Указываем таблицу в которую будем помещать данные (таблица должна существовать в шаблоне документа!)
            Microsoft.Office.Interop.Word.Table tbl = app.ActiveDocument.Tables[1];

            //Заполняем в таблицу - 10 записей.
            int i;
            for (i = 1; i <= exp.Count(); i++)
            {
                tbl.Rows.Add(ref missing);//Добавляем в таблицу строку.
                                          //Обычно саздаю только строку с заголовками и одну пустую для данных.
                tbl.Rows[i + 1].Cells[1].Range.Text = ((exp[i - 1]).a).ToString();
                tbl.Rows[i + 1].Cells[2].Range.Text = ((exp[i - 1]).c).ToString();
                

            }

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.

            object findText = "[data]";
            object replaceWith = "Директор";
            object replace = 2;

            app.Selection.Find.Execute(ref findText, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
            ref replace, ref missing, ref missing, ref missing, ref missing);

            object findText1 = "p2";
            object replaceWith1 = "Бухгалтер";
            object replace1 = 2;

            app.Selection.Find.Execute(ref findText1, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith1,
            ref replace1, ref missing, ref missing, ref missing, ref missing);



            //Открываем документ для просмотра.
            app.Visible = true;
            //app.Quit(); // выйти из word
            GC.Collect(); // убрать за  собой
        }
    }
}
