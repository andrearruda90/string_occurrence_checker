using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Configuration;
using System.Threading;
using System.Diagnostics;

namespace emailchecker
{
    public partial class Form1 : Form
    {
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public OpenFileDialog ofd = new OpenFileDialog();
        public FolderBrowserDialog fbd = new FolderBrowserDialog();

        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {   

            this.Text = "Verificador de E-mails";

            //disable resize window
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            this.MaximizeBox = false;

            tabPage1.Text = "Analisar"; tabPage2.Text = "Opções"; button1.Text = "Buscar";
            label1.Text = "Arquivo Origem"; label3.Text = "Status: Aguardando";label3.Enabled = false; label4.Text = "";
            button2.Text = "Buscar";label2.Text = "Destino"; button3.Text = "Analisar";button4.Text = "Adicionar";
            button5.Text = "Remover"; groupBox1.Text = "Considerar os itens abaixo:"; groupBox2.Text = "Adicionar/Remover";

            progressBar1.Enabled = false;


            //setting up listView2 items
            OpenLstvwItems();

            // Set to details view.
            listView1.View = View.Details;

            // Add a column with width 20 and left alignment.
            listView1.Columns.Add("Lista", 268, HorizontalAlignment.Left);

            //Removing Header
            listView1.HeaderStyle = ColumnHeaderStyle.None;
            listView1.MultiSelect = false;
        }
        public void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Open_File inputFile = new Open_File();
                Save_File outputFile = new Save_File();

                ofd.Filter = inputFile.filter;
                ofd.ShowDialog();

                inputFile.fileName = Path.GetFileName(ofd.FileName);
                inputFile.fullPath = Path.GetFullPath(ofd.FileName);
                outputFile.directoryPath = inputFile.fullPath.Replace($@"\{inputFile.fileName}", "");

                textBox1.Text = inputFile.fullPath;
                textBox2.Text = outputFile.directoryPath;



            }
            catch (Exception)
            {
                
                //throw;
            }
            
        }
        public void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Save_File outputFile = new Save_File();
                
                fbd.ShowDialog();
                
                outputFile.directoryPath =fbd.SelectedPath.ToString();
                textBox2.Text = outputFile.directoryPath;

            }
            catch (Exception)
            {

                //throw;
            }
            
        }
        private void button3_Click(object sender, EventArgs e)
        {
            //checking if both textboxs have valid path
            if (File.Exists(textBox1.Text) == false || Directory.Exists(textBox2.Text) == false)
            {
                MessageBox.Show("Insira Arquivo e Diretório Válido nos campos acima!","Campo Inválido",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            else
            {
                button1.Enabled = false; button2.Enabled = false; button3.Enabled = false; textBox1.Enabled = false;
                textBox2.Enabled = false; progressBar1.Enabled = true; label3.Enabled = true;
                
            }

            convertInputFile();
            Execution();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox3.Text != "")
            {
                listView1.Items.Add(textBox3.Text);
                SaveLstvwItems();
                textBox3.Text = "";
                
            }
            else
                MessageBox.Show("Digite a palavra-chave primeiro.", "Campo Vazio", MessageBoxButtons.OK,
                                                                                   MessageBoxIcon.Exclamation);
        }
        private void button5_Click(object sender, EventArgs e)
        {
            
            if (listView1.SelectedItems.Count != 0)
            {       
                DialogResult question = MessageBox.Show($"Remover \"{listView1.SelectedItems[0].Text}\" da lista?", "Confirmação", MessageBoxButtons.YesNo,
                                                                                                         MessageBoxIcon.Question);
                if (question == DialogResult.Yes)
                {
                    for (int i = listView1.SelectedItems.Count - 1; i >= 0; i--)
                    {
                        ListViewItem li = listView1.SelectedItems[i];
                        listView1.Items.Remove(li);
                        RemoveLstvwItems();
                    }
                }
                else if (question == DialogResult.No)
                {
                    //do something else
                }

                
            }
            else
            {
                MessageBox.Show("Selecione um item primeiro!", "Sem seleção", MessageBoxButtons.OK,
                                                                              MessageBoxIcon.Exclamation);
            }
            
            
        }
        private void SaveLstvwItems()
        {
            using (var tw = new StreamWriter(Environment.CurrentDirectory.ToString() + @"\Exclude.txt"))

            {
                foreach (ListViewItem item in listView1.Items)
                {
                    tw.WriteLine(item.Text);

                }
                tw.Close();
            }

        }
        private void RemoveLstvwItems()
        {
            File.Delete(Environment.CurrentDirectory.ToString() + @"\Exclude.txt");
            SaveLstvwItems();
        }
        private void OpenLstvwItems()
        {
            var totalrows = File.ReadLines(Environment.CurrentDirectory.ToString() + @"\Exclude.txt").Count();

            string[] strAllLines = System.IO.File.ReadAllLines(Environment.CurrentDirectory.ToString() + @"\Exclude.txt");

            for (int i = 0; i < totalrows; i++)
            {
                listView1.Items.Add(strAllLines[i]);
            }
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
        public void SaveFile()
        {

            Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            wb.SaveAs(textBox2.Text + @"\Resultado",51);
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
        public void getexcelprocesses()
        {


            Process[] safeProcess = Process.GetProcessesByName("EXCEL");
            foreach (Process p in safeProcess)
            {
                process.Add(p.Id);
            }
        }
        public void killExcelProcesses()
        {
            Process[] killProcess = Process.GetProcessesByName("EXCEL");
            foreach (Process p2 in killProcess)
            {
                int countp = 0;
                foreach (var i in process)
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
        public void convertInputFile()
        {
            Convert_File cf = new Convert_File();

            cf.fullPathOriginFile = textBox1.Text;
            cf.fileExtension = cf.fullPathOriginFile.Substring(cf.fullPathOriginFile.Length - 4, 4);
            cf.outputFileExtension = ".xlsx";
            cf.directoryName = "converting";
            cf.convertedFileName = @"\converting";
            cf.fullDirectoryName = Path.GetPathRoot(Environment.SystemDirectory) + cf.directoryName;
            cf.fullPathConvertingFile = $"{cf.fullDirectoryName}{cf.convertedFileName}{cf.fileExtension}";
            cf.fullPathConvertedFile = $"{cf.fullDirectoryName}{cf.convertedFileName}{cf.outputFileExtension}";

            //checking / creating work directory
            if (Directory.Exists(cf.fullDirectoryName))
            {
                Directory.Delete(cf.fullDirectoryName, true);
                Directory.CreateDirectory(cf.fullDirectoryName);
                File.Copy(cf.fullPathOriginFile, cf.fullPathConvertingFile, true);
            }
            else
            {
                Directory.CreateDirectory(cf.fullDirectoryName);
                File.Copy(cf.fullPathOriginFile, cf.fullPathConvertingFile, true);
            }


            // >>>>>>>> run python here <<<<<<<
            string exePath = System.AppDomain.CurrentDomain.BaseDirectory + @"csvconversor.exe";

            Process pro = new Process();
            pro.StartInfo.FileName = exePath;
            pro.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;

            pro.Start();





            //checking if converted file already exists
            int timeElapsed = 0;
        loop1:
            Thread.Sleep(1000);
            foreach (string file in Directory.GetFiles(cf.fullDirectoryName))
            {
                if (file.Contains("xlsx") || file.Contains("xls"))
                {
                    break;
                }
                else
                {
                    if (timeElapsed < 30)
                    {
                        timeElapsed++;
                        goto loop1;
                    }
                    else
                    {
                        MessageBox.Show("Arquivo convertido não existe, o programa será fechado.", "Erro", 
                                                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.Close();

                    }
                }
            }
            try
            {
                pro.CloseMainWindow();
            }
            catch (Exception)
            {
                //throw;
            }
                          
        }
        public Form1(string path, int Sheet)
        {
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }
        public void Execution()
        {
            
            label3.Text = "Status: Analisando...";

            getexcelprocesses();
            
            // getting input file name
            Form1 excel = new Form1(textBox1.Text, 1);
            
            Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            //finding head row
            int headRow = 0;
            for (int x = 0; x < 1000; x++)
            {
                if (excel.ReadCell(x, 1).ToString() != "" && excel.ReadCell(x, 1).ToString() != "SA1") //where 2 equals row 3
                {
                    headRow = x;
                    goto escape1;
                }
            }
        escape1:

            //finding the e-mail column
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

            foreach (ListViewItem lisViewItem in this.listView1.Items)
            {
                listItem.Add(lisViewItem.Text.ToLower());
            }

            int count = headRow + 1;
            progressBar1.Maximum = excel.LastRow();

            int pb = count;
            
            //filling fields
            while (count <= excel.LastRow())
            {

                if (listItem.Any(s => excel.ReadCell(count, targetColumn).ToString().Contains(s)) == true ||
                    excel.ReadCell(count, targetColumn).Contains("@") == false || excel.ReadCell(count, targetColumn).ToString() == "")
                {
                    excel.WriteCell(excel.ReadCell(count, targetColumn), count, targetColumn); //WriteCell(value,line,column)
                }
                count++;

                if (progressBar1.Value == progressBar1.Maximum)
                {
                    break;
                }
                else
                {
                    pb++;
                    progressBar1.Value = pb;
                }

            }
            //End Message
            MessageBox.Show("Processo Finalilzado!", "Finalizado", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            button1.Enabled = true; button2.Enabled = true; button3.Enabled = true; textBox1.Enabled = true;
            textBox2.Enabled = true; progressBar1.Enabled = false; label3.Enabled = false;
            label3.Text = "Status: Pronto!";

            excel.wb.SaveAs(textBox2.Text + @"\Resultado", 51);
            excel.wb.Close();
            //SaveFile();

            killExcelProcesses();
        }

    }
}

