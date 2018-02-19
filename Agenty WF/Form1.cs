using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Agenty_WF
{
    public partial class Form1 : Form
    {
        string file;
        

        public Form1()
        {
            InitializeComponent();
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
    }
}
