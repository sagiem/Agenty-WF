﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using System.Data.Common;



namespace Agenty_WF
{
    public partial class Form1 : Form
    {
        string file;
        private SQLiteConnection DB;


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DB = new SQLiteConnection("Data Source=C:\\Users\\user\\Desktop\\прога\\Agenty-WF\\files\\otchet_art.db");
            DB.Open();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void button_openfileYR_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Multiselect = false;
            openfile.DefaultExt = "*.xls;*.xlsx";
            openfile.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            openfile.Title = "Выберите документ Excel";
            openfile.ShowDialog();
            if (openfile.FileName != null)
            {
                this.file = openfile.FileName;
            }
        }

        private void button_otchetYR_Click(object sender, EventArgs e)
        {
            Raschet raschet = new Raschet(file, date_aktYR.Text, textb_aktnYR.Text);
            raschet.Exelreader();
            raschet.ExelOtchet();
        }

        private void button_aktYR_Click(object sender, EventArgs e)
        {
            MessageBox.Show(date_aktYR.Text);
            
        }

        private void Combob_YR_Click(object sender, EventArgs e)
        {
            SQLiteCommand command = new SQLiteCommand("select * from Агенты", DB);
            //DB.Open();
            Combob_YR.Items.Clear();
            DbDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                Combob_YR.Items.Add((string)reader["Агент"]);         //СтолбецТаблицы

            }
        }

       
    }
}
