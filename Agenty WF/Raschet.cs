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
        double d;
        double sum = 0;
        int z = 14;
        int t = 1;
        int j;
        private SQLiteConnection DB;
        double oplataYR;
        

        List<ExcelOpen> exp = new List<ExcelOpen>();
        Dictionary<string, YRdb> yr = new Dictionary<string, YRdb>();


        public Raschet(string file, string dateYR, string aktnYR, string AG, string oplataYR)
        {
            this.file = file;
            this.dateYR = dateYR;
            this.aktnYR = aktnYR;
            this.AG = AG;
            this.oplataYR= double.Parse(oplataYR.Replace(".", ","));
            
        }

        public void Yrread()
        {
            DB = new SQLiteConnection("Data Source=C:\\Agenty\\data\\otchet_art.db");
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

            //DB = new SQLiteConnection("Data Source=data\\otchet_art.db");
            //DB.Open();
            SQLiteCommand commandn = new SQLiteCommand("select * from Наша", DB);
            DbDataReader readern = commandn.ExecuteReader();
            while (readern.Read())
            {

                nashOrg = ((string)readern["Организация"]);
                nashvlic = ((string)readern["Влице"]);
                nashOsn = ((string)readern["Основание"]);
                nashPodp = ((string)readern["Подписант"]);

            }

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
                b = (b.ToUpper()).Replace("EEE", "").Replace("ЕЕЕ", "").Replace("MMM", "").Replace("МММ", "");
                if (b.Length == 10)
                {
                    e = b;
                    d = Double.Parse(c.Replace(" ", "").Replace(".", ","));
                    ExcelOpen excelOpen = new ExcelOpen(t, e, d);
                    exp.Add(excelOpen);
                    t++;
                    sum = sum + d;
                }
                z++;
            }

            j = exp.Count();

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

            object fileName = @"C:\Agenty\templates\otchet.docx";
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
            object replaceWith = aktnYR;
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


            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            //Задаём параметры замены и выполняем замену.
            object findText4 = "[организация1]";
            object replaceWith4 = nashOrg;
            object replace4 = 2;

            app.Selection.Find.Execute(ref findText4, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith4,
            ref replace4, ref missing, ref missing, ref missing, ref missing);


            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            //Задаём параметры замены и выполняем замену.
            object findText5 = "[организация2]";
            object replaceWith5 = AG;
            object replace5 = 2;

            app.Selection.Find.Execute(ref findText5, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith5,
            ref replace5, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            //Задаём параметры замены и выполняем замену.
            object findText6 = "[фамилия1]";
            object replaceWith6 = nashPodp;
            object replace6 = 2;

            app.Selection.Find.Execute(ref findText6, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith6,
            ref replace6, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            //Задаём параметры замены и выполняем замену.
            object findText7 = "[фамилия2]";
            object replaceWith7 = (yr[AG]).Podpisant;
            object replace7= 2;

            app.Selection.Find.Execute(ref findText7, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith7,
            ref replace7, ref missing, ref missing, ref missing, ref missing);



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

            object fileName = @"C:\Agenty\templates\akt.docx";
            object falseValue = false;
            object trueValue = true;
            object missing = Type.Missing;

            doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);

            //Указываем таблицу в которую будем помещать данные (таблица должна существовать в шаблоне документа!)
            Microsoft.Office.Interop.Word.Table tbl = app.ActiveDocument.Tables[1];        
            tbl.Rows[2].Cells[1].Range.Text = (j).ToString();
            tbl.Rows[2].Cells[2].Range.Text = (sum).ToString();

            Yrread();

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.

            object findText0 = "[акт]";
            object replaceWith0 = aktnYR;
            object replace0 = 2;

            app.Selection.Find.Execute(ref findText0, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith0,
            ref replace0, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.

            object findText = "[договор]";
            object replaceWith = (yr[AG]).Dogovor;
            object replace = 2;

            app.Selection.Find.Execute(ref findText, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
            ref replace, ref missing, ref missing, ref missing, ref missing);

            object findText1 = "[дата]";
            object replaceWith1 = dateYR;
            object replace1 = 2;

            app.Selection.Find.Execute(ref findText1, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith1,
            ref replace1, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.

            object findText2 = "[организация1]";
            object replaceWith2 = nashOrg;
            object replace2 = 2;

            app.Selection.Find.Execute(ref findText2, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith2,
            ref replace2, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.

            object findText3 = "[лицо1]";
            object replaceWith3 = nashvlic;
            object replace3 = 2;

            app.Selection.Find.Execute(ref findText3, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith3,
            ref replace3, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.

            object findText4 = "[основание1]";
            object replaceWith4 = nashOsn;
            object replace4 = 2;

            app.Selection.Find.Execute(ref findText4, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith4,
            ref replace4, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.


            object findText5 = "[организация2]";
            object replaceWith5 = AG;
            object replace5 = 2;

            app.Selection.Find.Execute(ref findText5, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith5,
            ref replace5, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.


            object findText6 = "[лицо2]";
            object replaceWith6 = (yr[AG]).Vlice;
            object replace6 = 2;

            app.Selection.Find.Execute(ref findText6, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith6,
            ref replace6, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.


            object findText7 = "[основание2]";
            object replaceWith7 = (yr[AG]).Osnovanie;
            object replace7 = 2;

            app.Selection.Find.Execute(ref findText7, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith7,
            ref replace7, ref missing, ref missing, ref missing, ref missing);


            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.


            object findText8 = "[основание2]";
            object replaceWith8 = (yr[AG]).Osnovanie;
            object replace8 = 2;

            app.Selection.Find.Execute(ref findText8, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith8,
            ref replace8, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.


            object findText9 = "[сумма]";
            object replaceWith9 = Math.Floor((sum/100)*oplataYR);
            object replace9 = 2;

            app.Selection.Find.Execute(ref findText9, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith9,
            ref replace9, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.


            object findText10 = "[сумма2]";
            object replaceWith10 = RusNumber.Str(Convert.ToInt32(Math.Floor((sum / 100) * oplataYR)));
            object replace10 = 2;

            app.Selection.Find.Execute(ref findText10, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith10,
            ref replace10, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.


            object findText11 = "[сумма3]";
            object replaceWith11 = RusNumber.Kop(Math.Round(((sum / 100) * oplataYR),2));
            object replace11 = 2;

            app.Selection.Find.Execute(ref findText11, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith11,
            ref replace11, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.


            object findText12 = "[организация1]";
            object replaceWith12 = nashOrg;
            object replace12 = 2;

            app.Selection.Find.Execute(ref findText12, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith12,
            ref replace12, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.


            object findText14 = "[организация2]";
            object replaceWith14 = AG;
            object replace14 = 2;

            app.Selection.Find.Execute(ref findText14, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith14,
            ref replace14, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.


            object findText15 = "[фамилия1]";
            object replaceWith15 = nashPodp;
            object replace15 = 2;

            app.Selection.Find.Execute(ref findText15, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith15,
            ref replace15, ref missing, ref missing, ref missing, ref missing);

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.


            object findText16 = "[фамилия2]";
            object replaceWith16 = (yr[AG]).Podpisant;
            object replace16 = 2;

            app.Selection.Find.Execute(ref findText16, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith16,
            ref replace16, ref missing, ref missing, ref missing, ref missing);

            //Открываем документ для просмотра.
            app.Visible = true;
            GC.Collect(); // убрать за  собой
        }
    }
}
