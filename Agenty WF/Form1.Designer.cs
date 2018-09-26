namespace Agenty_WF
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.date_aktYR = new System.Windows.Forms.DateTimePicker();
            this.button_editYR = new System.Windows.Forms.Button();
            this.button_otchetYR = new System.Windows.Forms.Button();
            this.button_aktYR = new System.Windows.Forms.Button();
            this.button_openfileYR = new System.Windows.Forms.Button();
            this.textb_aktnYR = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textb_oplataYR = new System.Windows.Forms.TextBox();
            this.Combob_YR = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(13, 13);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(504, 300);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.date_aktYR);
            this.tabPage1.Controls.Add(this.button_editYR);
            this.tabPage1.Controls.Add(this.button_otchetYR);
            this.tabPage1.Controls.Add(this.button_aktYR);
            this.tabPage1.Controls.Add(this.button_openfileYR);
            this.tabPage1.Controls.Add(this.textb_aktnYR);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.textb_oplataYR);
            this.tabPage1.Controls.Add(this.Combob_YR);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(496, 274);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Юридические лица";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // date_aktYR
            // 
            this.date_aktYR.Location = new System.Drawing.Point(128, 95);
            this.date_aktYR.Name = "date_aktYR";
            this.date_aktYR.Size = new System.Drawing.Size(144, 20);
            this.date_aktYR.TabIndex = 12;
            // 
            // button_editYR
            // 
            this.button_editYR.Location = new System.Drawing.Point(396, 10);
            this.button_editYR.Name = "button_editYR";
            this.button_editYR.Size = new System.Drawing.Size(95, 23);
            this.button_editYR.TabIndex = 11;
            this.button_editYR.Text = "Редактировать";
            this.button_editYR.UseVisualStyleBackColor = true;
            // 
            // button_otchetYR
            // 
            this.button_otchetYR.Location = new System.Drawing.Point(30, 183);
            this.button_otchetYR.Name = "button_otchetYR";
            this.button_otchetYR.Size = new System.Drawing.Size(75, 23);
            this.button_otchetYR.TabIndex = 10;
            this.button_otchetYR.Text = "Отчет";
            this.button_otchetYR.UseVisualStyleBackColor = true;
            this.button_otchetYR.Click += new System.EventHandler(this.button_otchetYR_Click);
            // 
            // button_aktYR
            // 
            this.button_aktYR.Location = new System.Drawing.Point(128, 183);
            this.button_aktYR.Name = "button_aktYR";
            this.button_aktYR.Size = new System.Drawing.Size(75, 23);
            this.button_aktYR.TabIndex = 9;
            this.button_aktYR.Text = "Акт";
            this.button_aktYR.UseVisualStyleBackColor = true;
            this.button_aktYR.Click += new System.EventHandler(this.button_aktYR_Click);
            // 
            // button_openfileYR
            // 
            this.button_openfileYR.Location = new System.Drawing.Point(30, 135);
            this.button_openfileYR.Name = "button_openfileYR";
            this.button_openfileYR.Size = new System.Drawing.Size(103, 23);
            this.button_openfileYR.TabIndex = 8;
            this.button_openfileYR.Text = "Загрузить файл";
            this.button_openfileYR.UseVisualStyleBackColor = true;
            this.button_openfileYR.Click += new System.EventHandler(this.button_openfileYR_Click);
            // 
            // textb_aktnYR
            // 
            this.textb_aktnYR.Location = new System.Drawing.Point(30, 95);
            this.textb_aktnYR.Name = "textb_aktnYR";
            this.textb_aktnYR.Size = new System.Drawing.Size(67, 20);
            this.textb_aktnYR.TabIndex = 7;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label5.Location = new System.Drawing.Point(103, 98);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(19, 13);
            this.label5.TabIndex = 6;
            this.label5.Text = "от";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.Location = new System.Drawing.Point(6, 98);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(20, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "№";
            // 
            // textb_oplataYR
            // 
            this.textb_oplataYR.Location = new System.Drawing.Point(104, 40);
            this.textb_oplataYR.Name = "textb_oplataYR";
            this.textb_oplataYR.Size = new System.Drawing.Size(60, 20);
            this.textb_oplataYR.TabIndex = 4;
            this.textb_oplataYR.Text = "9";
            // 
            // Combob_YR
            // 
            this.Combob_YR.FormattingEnabled = true;
            this.Combob_YR.Location = new System.Drawing.Point(104, 12);
            this.Combob_YR.Name = "Combob_YR";
            this.Combob_YR.Size = new System.Drawing.Size(271, 21);
            this.Combob_YR.TabIndex = 3;
            this.Combob_YR.Click += new System.EventHandler(this.Combob_YR_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(6, 73);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(221, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Акт об оказании услуг по вводу ОСАГО";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(6, 43);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(98, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Оплата по вводу";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(6, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Агент";
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(496, 274);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Физические лица";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(529, 325);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Отчет-акт";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.DateTimePicker date_aktYR;
        private System.Windows.Forms.Button button_editYR;
        private System.Windows.Forms.Button button_otchetYR;
        private System.Windows.Forms.Button button_aktYR;
        private System.Windows.Forms.Button button_openfileYR;
        private System.Windows.Forms.TextBox textb_aktnYR;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textb_oplataYR;
        private System.Windows.Forms.ComboBox Combob_YR;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
    }
}

