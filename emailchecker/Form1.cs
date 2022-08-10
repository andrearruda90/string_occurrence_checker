using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using OfficeOpenXml;



namespace emailchecker
{
    public partial class Form1 : Form
    {
        public OpenFileDialog ofd = new OpenFileDialog();
        public FolderBrowserDialog fbd = new FolderBrowserDialog();

        public Form1() //metodo construtor
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
            label1.Text = "Arquivo Origem"; label3.Text = "Status: Aguardando"; label3.Enabled = false; label4.Text = "";
            button2.Text = "Buscar"; label2.Text = "Destino"; button3.Text = "Analisar"; button4.Text = "Adicionar";
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

                if (ofd.FileName.EndsWith(".csv") == false)
                {
                    textBox2.Enabled = false;
                }
                else
                {
                    textBox2.Enabled = true;
                }



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

                outputFile.directoryPath = fbd.SelectedPath.ToString();
                textBox2.Text = outputFile.directoryPath;

            }
            catch (Exception)
            {

                //throw;
            }

        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                //checking if both textboxs have valid path
                if (File.Exists(textBox1.Text) == false || Directory.Exists(textBox2.Text) == false)
                {
                    MessageBox.Show("Insira Arquivo e Diretório Válido nos campos acima!", "Campo Inválido", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                else
                {
                    button1.Enabled = false; button2.Enabled = false; button3.Enabled = false; textBox1.Enabled = false;
                    textBox2.Enabled = false; progressBar1.Enabled = true; label3.Enabled = true;
                    Execution();
                }
            }
            catch (Exception)
            {
                return;
                // throw;
            }

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


        public void Execution()
        {
            //// according to the Polyform Noncommercial license: (Needed)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            label3.Text = "Status: Analisando...";

            //passing every listview1.Item to a new string list

            var listItem = new List<string>();

            foreach (ListViewItem lisViewItem in this.listView1.Items)
            {
                listItem.Add(lisViewItem.Text.ToLower());
            }

            //_-_-_-_-_-_-_-_-_-_-_-_-_-_-__-_-_-_-_-_-_-_-_-_-_-_-_-_-__-_-_-_-_-_-_-_-_-_-_-_-_-_-_
            // path to your excel file

            string path = null;
            if (textBox1.Text.EndsWith(".csv"))
            {
                convertInputFile();
                path = Path.GetPathRoot(Environment.SystemDirectory) + @"converting\converted.xlsx";
            }
            else
            {
                path = textBox1.Text;
            }
            FileInfo fileInfo = new FileInfo(path);

            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            // get number of rows in the sheet
            int rows = worksheet.Dimension.Rows;
            progressBar1.Maximum = rows;


            //defining headRow
            //_-_-_-_-_-_-_-_-_-_-_-_-_-_-__-_-_-_-_-_-_-_-_-_-_-_-_-_-__-_-_-_-_-_-_-_-_-_-_-_-_-_-_
            int headRow = 0;
            for (int row = 1; row <= rows; row++)
            {
                ExcelRange cel = worksheet.Cells[row, 1];
                string celValue = cel.Value == null ? string.Empty : cel.Value.ToString();
                if (celValue != "SA1" && celValue != "")
                {
                    headRow = row;
                    goto exit1;
                }
            }
        exit1:
            //_-_-_-_-_-_-_-_-_-_-_-_-_-_-__-_-_-_-_-_-_-_-_-_-_-_-_-_-__-_-_-_-_-_-_-_-_-_-_-_-_-_-_

            //targetColumn aqui
            //_-_-_-_-_-_-_-_-_-_-_-_-_-_-__-_-_-_-_-_-_-_-_-_-_-_-_-_-__-_-_-_-_-_-_-_-_-_-_-_-_-_-_
            int targertColumn = 0;
            int columns = worksheet.Dimension.Columns;

            for (int col = 1; col < columns; col++)
            {
                ExcelRange cel2 = worksheet.Cells[headRow, col];
                string celValue2 = cel2.Value == null ? string.Empty : cel2.Value.ToString();
                if (celValue2.Contains("mail"))
                {
                    targertColumn = col;
                    goto exit2;
                }
            }
        exit2:
            //_-_-_-_-_-_-_-_-_-_-_-_-_-_-__-_-_-_-_-_-_-_-_-_-_-_-_-_-__-_-_-_-_-_-_-_-_-_-_-_-_-_-_


            // loop through the worksheet rows
            for (int row = headRow + 1; row <= rows; row++)
            {
                int column = targertColumn;
                ExcelRange cel = worksheet.Cells[row, column];
                string celValue = cel.Value == null ? string.Empty : cel.Value.ToString();

                foreach (string argument in listItem)
                {
                    if (listItem.Any(s => celValue.Contains(s)) == true || celValue.Contains("@") == false || celValue == "")
                    {
                        cel.Style.Font.Color.SetColor(0, 255, 0, 0); // see rgb table online
                        progressBar1.Value = row;
                    }
                }
            }

            try
            {
                // save changes
                if (textBox1.Text.EndsWith(".csv"))
                {
                    package.SaveAs(textBox2.Text);
                }
                else
                {
                    package.Save();
                }

            }
            catch (Exception)
            {
                //throw;
            }
            //_-_-_-_-_-_-_-_-_-_-_-_-_-_-__-_-_-_-_-_-_-_-_-_-_-_-_-_-__-_-_-_-_-_-_-_-_-_-_-_-_-_-_


            //End Message
            label3.Text = "Completo!";
            MessageBox.Show("Processo Finalilzado!", "Finalizado", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            button1.Enabled = true; button2.Enabled = true; button3.Enabled = true; textBox1.Enabled = true;
            textBox2.Enabled = true; progressBar1.Enabled = false; label3.Enabled = false;


        }
        public void convertInputFile()
        {
            string fullPathOriginFile = textBox1.Text;
            string fileExtension = fullPathOriginFile.Substring(fullPathOriginFile.Length - 4, 4);
            string outputFileExtension = ".xlsx";
            string directoryName = "converting";
            string convertedFileName = @"\converting";
            string fullDirectoryName = Path.GetPathRoot(Environment.SystemDirectory) + directoryName;
            string fullPathConvertingFile = $"{fullDirectoryName}{convertedFileName}{fileExtension}";
            string fullPathConvertedFile = $"{fullDirectoryName}{convertedFileName}{outputFileExtension}";




            //checking / creating work directory
            if (Directory.Exists(fullDirectoryName))
            {
                Directory.Delete(fullDirectoryName, true);
                Directory.CreateDirectory(fullDirectoryName);
                File.Copy(fullPathOriginFile, fullPathConvertingFile, true);
                Thread.Sleep(5000);
            }
            else
            {
                Directory.CreateDirectory(fullDirectoryName);
                File.Copy(fullPathOriginFile, fullPathConvertingFile, true);
                Thread.Sleep(5000);
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
            foreach (string file in Directory.GetFiles(fullDirectoryName))
            {
                if (file.EndsWith(".xlsx") || file.EndsWith(".xls"))
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
    }
}

