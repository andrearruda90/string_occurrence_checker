namespace mbox
{

    public class myclass
    {
        public static void Main()
        {

            string fullPathOriginFile = @"c:\users\andre\desktop\myFile.csv";
            string fileExtension = fullPathOriginFile.Substring(fullPathOriginFile.Length - 4, 4);
            string outputFileExtension = ".xlsx";
            string directoryName = "converting";
            string convertedFileName = @"\converting";
            string fullDirectoryName = Path.GetPathRoot(Environment.SystemDirectory) + directoryName;
            string fullPathConvertingFile = $"{fullDirectoryName}{convertedFileName}{fileExtension}";
            string fullPathConvertedFile = $"{fullDirectoryName}{convertedFileName}{outputFileExtension}";

            string exePath = System.AppDomain.CurrentDomain.BaseDirectory + @"converting.exe";
            Console.WriteLine(fullPathConvertedFile);
        }
    }
}