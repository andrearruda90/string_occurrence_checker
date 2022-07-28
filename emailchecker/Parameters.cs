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
using emailchecker;


namespace emailchecker
{
    public class Open_File
    {
        public string filter = "csv files (*.csv)|*.csv|xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls";
        public string fileName;
        public string fileExtension;
        public string directoryPath;
        public string fullPath;
    }
    public class Save_File
    {
        public string fileName;
        public string directoryPath;
        public int fileFormat = 51;
    }

}
