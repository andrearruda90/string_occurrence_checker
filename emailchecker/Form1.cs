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
using System.Configuration;

namespace emailchecker
{

    public partial class Form1 : Form
    {


        private OpenFileDialog ofd = new OpenFileDialog();
        private FolderBrowserDialog fbd = new FolderBrowserDialog();
        public Form1()
        {


            //#S
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "Verificador de E-mails";

            //disable resize window
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            this.MaximizeBox = false;

            tabPage1.Text = "Analisar";
            tabPage2.Text = "Opções";
            tabPage3.Text = "Sobre";
            button1.Text = "Buscar";
            label1.Text = "Arquivo Origem";
            label3.Text = "Status: Aguardando";
            label3.Enabled = false; 
            label4.Text = "";
            button2.Text = "Buscar";
            label2.Text = "Destino";
            button3.Text = "Analisar";
            button4.Text = "Adicionar";
            button5.Text = "Remover";
            groupBox1.Text = "Considerar os itens abaixo:";
            groupBox2.Text = "Adicionar/Remover";

            progressBar1.Enabled = false;


            //setting up listView2 items
            OpenLstvwItems();

            //SetSetting("outputPath", "Empty");
            //SetSetting("outputFilename", "Empty");
            // Set to details view.
            listView1.View = View.Details;
            // Add a column with width 20 and left alignment.
            listView1.Columns.Add("Lista", 268, HorizontalAlignment.Left);
            //Removing Header
            listView1.HeaderStyle = ColumnHeaderStyle.None;
            listView1.MultiSelect = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Configuration configuration =
            ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            ofd.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)" +
                    "|*.xls|csv files (*.csv)|*.csv";//|All files (*.*)|*.*";
            ofd.ShowDialog();
            textBox1.Text = ofd.FileName;

            string fileName = Path.GetFileName(textBox1.Text);
            string fileNameWExt = Path.GetFileNameWithoutExtension(textBox1.Text);

            textBox2.Text = textBox1.Text.Replace($@"\{fileName}", "");

            SetSetting("outputFilename", $@"\{fileNameWExt} - Analisado{fileName.Replace(fileName, "")}");
            SetSetting("outputPath", $@"{textBox2.Text}{configuration.AppSettings.Settings
                                        ["outputFilename"].Value.ToString()}".ToString());
            SetSetting("inputName", textBox1.Text.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Configuration configuration =
            ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            fbd.ShowDialog();
            
            outputpath = textBox2.Text;

            string userpath = fbd.SelectedPath.ToString();
            string fileName = Path.GetFileName(textBox1.Text);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(textBox1.Text);

            textBox2.Text = userpath;

            SetSetting("outputPath", $@"{userpath}\{fileNameWithoutExtension} - Analisado{fileName.Replace(fileNameWithoutExtension, "")}");
            SetSetting("inputName", textBox1.Text.ToString());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //checking if both textboxs have valid path
            if (File.Exists(textBox1.Text) == false || Directory.Exists(textBox2.Text) == false)
            {
                MessageBox.Show("Insira Arquivo e Diretório Válido nos campos acima!","Campo vazio",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            else
            {
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                textBox1.Enabled = false;
                textBox2.Enabled = false;
                progressBar1.Enabled = true;
                label3.Enabled = true;
                HelloWorld();
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

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
 
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
           
        }
    }
}

