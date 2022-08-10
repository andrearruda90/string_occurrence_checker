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
        public class File_Action
        {
            public string fileName { get; set; }
            public string fileExtension { get; set; }
            public string directoryPath { get; set; }
            public string fullPath { get; set; }
        }
        public class Open_File:File_Action
        {
            public string filter { get; set; } = "csv files (*.csv)|*.csv|xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls";
        }
        public class Save_File:File_Action
        {
            public int fileFormat { get; set; } = 51;
        }

    public class Convert_File : File_Action
    {
        public string convertedFileExtension { get; set; }
        public string convertedFileName { get; set; }
        public string fullPathOriginFile { get;set;}
        public string fullPathConvertedFile { get; set; }
        public string fullPathConvertingFile { get; set; }
        public string directoryName { get; set; }
        public string fullDirectoryName { get; set; }
    }
}
