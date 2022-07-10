using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Runtime;
using System.Configuration;


namespace emailchecker
{

    partial class Form1
    {
        

        string outputpath = "";
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button5 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.listView1 = new System.Windows.Forms.ListView();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.button3 = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button2 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage2.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Controls.Add(this.groupBox1);
            resources.ApplyResources(this.tabPage2, "tabPage2");
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button5);
            this.groupBox2.Controls.Add(this.button4);
            this.groupBox2.Controls.Add(this.textBox3);
            resources.ApplyResources(this.groupBox2, "groupBox2");
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.TabStop = false;
            // 
            // button5
            // 
            resources.ApplyResources(this.button5, "button5");
            this.button5.Name = "button5";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button4
            // 
            resources.ApplyResources(this.button4, "button4");
            this.button4.Name = "button4";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // textBox3
            // 
            resources.ApplyResources(this.textBox3, "textBox3");
            this.textBox3.Name = "textBox3";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.listView1);
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // listView1
            // 
            this.listView1.HideSelection = false;
            resources.ApplyResources(this.listView1, "listView1");
            this.listView1.Name = "listView1";
            this.listView1.UseCompatibleStateImageBehavior = false;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.groupBox4);
            this.tabPage1.Controls.Add(this.groupBox3);
            resources.ApplyResources(this.tabPage1, "tabPage1");
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.label4);
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Controls.Add(this.progressBar1);
            this.groupBox4.Controls.Add(this.button3);
            resources.ApplyResources(this.groupBox4, "groupBox4");
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.TabStop = false;
            // 
            // label4
            // 
            resources.ApplyResources(this.label4, "label4");
            this.label4.Name = "label4";
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.Name = "label3";
            // 
            // progressBar1
            // 
            resources.ApplyResources(this.progressBar1, "progressBar1");
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Click += new System.EventHandler(this.progressBar1_Click);
            // 
            // button3
            // 
            resources.ApplyResources(this.button3, "button3");
            this.button3.Name = "button3";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.button2);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.textBox2);
            this.groupBox3.Controls.Add(this.textBox1);
            this.groupBox3.Controls.Add(this.button1);
            this.groupBox3.Controls.Add(this.label1);
            resources.ApplyResources(this.groupBox3, "groupBox3");
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.TabStop = false;
            // 
            // button2
            // 
            resources.ApplyResources(this.button2, "button2");
            this.button2.Name = "button2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // textBox2
            // 
            resources.ApplyResources(this.textBox2, "textBox2");
            this.textBox2.Name = "textBox2";
            // 
            // textBox1
            // 
            resources.ApplyResources(this.textBox1, "textBox1");
            this.textBox1.Name = "textBox1";
            // 
            // button1
            // 
            resources.ApplyResources(this.button1, "button1");
            this.button1.Name = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            resources.ApplyResources(this.tabControl1, "tabControl1");
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            // 
            // tabPage3
            // 
            resources.ApplyResources(this.tabPage3, "tabPage3");
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabPage2.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.ResumeLayout(false);

        }



        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public void HelloWorld()
        {

            label3.Text = "Status: Analisando...";
            OpenFile();


            void OpenFile()
            {

                //#A Creating a safeProcess List 
                List<int> safePId = new List<int>();

                Process[] safeProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process p in safeProcess)
                {
                    safePId.Add(p.Id);
                }

                Form1 excel = new Form1(textBox1.Text, 1); // input file name
                Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                //finding head row
                int headRow = 0;
                for (int x = 0; x<1000; x++)
                {
                    if (excel.ReadCell(x, 1).ToString() != "" && excel.ReadCell(x, 1).ToString() != "SA1") //where 2 equals row 3
                    {
                        headRow = x;
                        goto escape1;
                    }
                }
                escape1:
                //finding the right column

                int targetColumn = 0;
                for (int i = 0; i <= excel.LastColumn(); i++)
                {
                    if (excel.ReadCell(headRow, i).Contains("mail") == true) //where 2 equals row 3
                    {
                        targetColumn = i;
                        goto escape2;
                    }
                }
                escape2:

                //passing every listview1.Item to a new string list

                var listItem = new List<string>();

                foreach(ListViewItem lisViewItem in this.listView1.Items)
                {
                    listItem.Add(lisViewItem.Text.ToLower());
                }

                // reading cells to checkup if they match with listitems
                int count = headRow +1;
                progressBar1.Maximum = excel.LastRow();

                for (int y = 0; y < headRow; y++)
                {
                    progressBar1.Value = y;
                    if (progressBar1.Value == headRow - 1)
                    {
                        goto escape3;
                    }
                }
                escape3:
                for (int i = 1; i <= excel.LastRow(); i++)
                {
                    progressBar1.Value = count;
                    if (listItem.Any(s=>excel.ReadCell(count, targetColumn).ToString().Contains(s)) == true || 
                        excel.ReadCell(count, targetColumn).Contains("@") == false || excel.ReadCell(count, targetColumn).ToString() == "")
                    {
                        excel.WriteCell(excel.ReadCell(count, targetColumn), count, targetColumn); //WriteCell(value,line,column)
                                                                             //                                                             //were 1 = 2
                    } 
                    count++;
                    

                    if (count == excel.LastRow()) // why foreach doesn't work here????? I had to put this to run normally
                    {
                        progressBar1.Value++;
                        MessageBox.Show("Processo Finalilzado!", "Finalizado", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        button1.Enabled = true;
                        button2.Enabled = true;
                        button3.Enabled = true;
                        textBox1.Enabled = true;
                        textBox2.Enabled = true;
                        progressBar1.Enabled = false;
                        label3.Enabled = false;
                        goto exit;
                    }
                } 

                exit:


                label3.Text = "Status: Pronto!";
                excel.SaveFile();

                //#A - Killing bad EXCEl processes
                Process[] killProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process p2 in killProcess)
                {
                    int countp = 0;
                    foreach (var i in safePId)
                    {
                        if (p2.Id == i)
                        {
                            countp++;
                        }
                    }

                    if (countp == 0)
                    {
                        p2.Kill();
                    }
                }
            }

        }
        public Form1(string path, int Sheet)
        {


            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];

        }

        public string ReadCell(int i, int j)
        {

            i++;
            j++;


            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2.ToString();
            else
                return "";

        }

        public string WriteCell(string k, int i, int j)
        {
            i++;
            j++;

            ws.Cells[i, j].Value2 = k;
            ws.Cells[i, j].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            return k;



        }

         void SaveFile()
        {
            Configuration configuration =
            ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
           
            wb.SaveAs(@"C:\users\andre\desktop\Resutado",51); //output file name
            wb.Close();
        }

        public int LastRow()
        {
            _Excel.Range last = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            _Excel.Range range = ws.get_Range("A1", last);

            int lastUsedRow = last.Row;
            return lastUsedRow;
        }

        public int LastColumn()
        {
            _Excel.Range last = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            _Excel.Range range = ws.get_Range("B1", last);

            int lastUsedcolumn = last.Column;

            return lastUsedcolumn;
        }
        private static void SetSetting(string key, string value)
        {
            Configuration configuration =
            ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            configuration.AppSettings.Settings[key].Value = value;
            configuration.Save(ConfigurationSaveMode.Full, true);
            ConfigurationManager.RefreshSection("appSettings");
        }

        #endregion

        private TabPage tabPage2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.GroupBox groupBox1;
        private ListView listView1;
        private TabPage tabPage1;
        private ProgressBar progressBar1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private TabControl tabControl1;
        private TabPage tabPage3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Label label4;
    }

}

