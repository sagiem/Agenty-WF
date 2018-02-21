using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SQLite;
using System.Data.Common;

namespace Agenty_WF
{
    class Raschet
    {
        public string file;
        string b, c, e, dateYR, aktnYR, AG;
        string nashOrg, nashvlic, nashOsn, nashPodp;
        decimal d;
        decimal sum = 0;
        int z = 14;
        int t = 1;
        private SQLiteConnection DB;

        List<ExcelOpen> exp = new List<ExcelOpen>();
        Dictionary<string, YRdb> yr = new Dictionary<string, YRdb>();


        public Raschet(string file, string dateYR, string aktnYR, string AG)
        {
            this.file = file;
            this.dateYR = dateYR;
            this.aktnYR = aktnYR;
            this.AG = AG;
        }

        public void Yrread()
        {
            DB = new SQLiteConnection("Data Source=data\\otchet_art.db");
            DB.Open();
            SQLiteCommand command = new SQLiteCommand("select * from Агенты", DB);
            DbDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {

                string Agent = ((string)reader["Агент"]);
                string Vlice = ((string)reader["Влице"]);
                string Osnovanie = ((string)reader["Основание"]);
                string Dogovor = ((string)reader["Договор"]);
                string DataDog = ((string)reader["Дата"]);
                string Podpisant = ((string)reader["Подписант"]);

                YRdb yRdb = new YRdb(Vlice, Osnovanie, Dogovor, Podpisant);
                yr.Add(Agent, yRdb);
            }




            //SQLiteCommand command1 = new SQLiteCommand("select * from Наша", DB);
            //DbDataReader reader1 = command1.ExecuteReader();


            //    nashOrg = ((string)reader1["Организация"]);
            //    nashvlic = ((string)reader1["Влице"]);
            //    nashOsn = ((string)reader1["Основание"]);
            //    nashPodp = ((string)reader1["Подписант"]);
            

            DB.Close();
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

            Yrread();

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            //Задаём параметры замены и выполняем замену.
            object findText = "[акт]";
            object replaceWith = dateYR;
            object replace = 2;

            app.Selection.Find.Execute(ref findText, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
            ref replace, ref missing, ref missing, ref missing, ref missing);


            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            //Задаём параметры замены и выполняем замену.
            object findText1 = "[дата_акта]";
            object replaceWith1 = dateYR;
            object replace1 = 2;

            app.Selection.Find.Execute(ref findText1, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith1,
            ref replace1, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            //Задаём параметры замены и выполняем замену.
            object findText2 = "[data]";
            object replaceWith2 = dateYR;
            object replace2 = 2;

            app.Selection.Find.Execute(ref findText2, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith2,
            ref replace2, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            //Задаём параметры замены и выполняем замену.
            object findText3 = "[договор]";
            object replaceWith3 = (yr[AG]).Dogovor;
            object replace3 = 2;

            app.Selection.Find.Execute(ref findText3, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith3,
            ref replace3, ref missing, ref missing, ref missing, ref missing);



            //Открываем документ для просмотра.
            app.Visible = true;
            GC.Collect(); // убрать за  собой
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
            GC.Collect(); // убрать за  собой
        }
    }
}
