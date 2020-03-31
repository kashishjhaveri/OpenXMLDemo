using System;
using System.Collections.Generic;
using System.Text;


namespace OpenXMLDemo.Models
{
    class Constants
    {
        public class Locations
        {
            public readonly static string data_url = @"https://raw.githubusercontent.com/kashishjhaveri/OpenXMLDemo/master/books.csv";

            //string filePath = @"C:\Users\kashi\Documents\Big_Data_Analytics\Semester_1\1001-Information_Encoding_Standards\info.csv";
            public readonly static string DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            public readonly static string ExePath = Environment.CurrentDirectory;

            public readonly static string ContentsFolder = $"{ExePath}\\..\\..\\..\\Content";
            public readonly static string ImagesFolder = $"{ContentsFolder}\\Images";
            public readonly static string DataFolder = $"{ContentsFolder}\\Data";
            public const string InfoFile = "info.csv";
            public const string ImageFile = "myimage.jpg";

            //string filePath = $@"{desktopPath}\info.csv";
            public readonly static string FilePath = $@"{DataFolder}\info.csv";
        }
        
        public class FTP
        {
            public const string UserName = @"bdat100119f\bdat1001";
            public const string Password = "bdat1001";

            public const string BaseUrl = "ftp://waws-prod-dm1-127.ftp.azurewebsites.windows.net/bdat1001-20914/200448232 Kashish Jhaveri";
            public const string ImageUrl = BaseUrl + "/myimage.jpg";
            public const string DocFile = "Goodreads-Books.docx";
            public const string ExcelFile = "Goodreads-Books.xlsx";
            public const string PptFile = "Goodreads-Books.pptx";

            public const int OperationPauseTime = 0000;
        }
    }
}
