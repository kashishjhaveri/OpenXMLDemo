using OpenXMLDemo.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

namespace OpenXMLDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            List<List<string>> DataLines = new List<List<string>>();
            List<string> CsvData = new List<string>();
            string DocFilepath = $@"{Constants.Locations.DesktopPath}\{Constants.FTP.DocFile}";
            string ExcelFilepath = $@"{Constants.Locations.DesktopPath}\{Constants.FTP.ExcelFile}";
            string PresentationFilepath = $@"{Constants.Locations.DesktopPath}\{Constants.FTP.PptFile}";

            var request = (HttpWebRequest)WebRequest.Create(Constants.Locations.data_url);
            var response = (HttpWebResponse)request.GetResponse();
            string responseString;
            using (var stream = response.GetResponseStream())
            {
                using (var reader = new StreamReader(stream))
                {
                    while (!reader.EndOfStream)
                    {
                        responseString = reader.ReadLine();
                        CsvData = responseString.Split(',').ToList();
                        DataLines.Add(CsvData);
                    }
                }
            }

            List<string> HeaderData = DataLines.First();
            DataLines.RemoveAt(0);

            Console.WriteLine("Creating Word File...");
            Models.OpenXMLUtilites.WordDoc.CreateWordprocessingDocument(HeaderData, DataLines);
            Console.WriteLine("Word File Created!");

            Console.WriteLine("Creating Presentation...");
            Models.OpenXMLUtilites.Presentation.CreatePresentation();
            Models.OpenXMLUtilites.Presentation.AddImage(PresentationFilepath, $@"{Constants.Locations.DesktopPath}\{Constants.Locations.ImageFile}");
            Models.OpenXMLUtilites.Presentation.InsertNewSlide(PresentationFilepath, HeaderData, DataLines, "Data");
            Console.WriteLine("Presentation Created!");

            Console.WriteLine("Creating Excel...");
            Models.OpenXMLUtilites.Excel.CreateSpreadsheetWorkbook();
            Models.OpenXMLUtilites.Excel.InsertData(HeaderData, DataLines);
            Console.WriteLine("Excel Created!");

            FTP.uploadFile(Constants.FTP.DocFile, File.ReadAllBytes(DocFilepath));
            FTP.uploadFile(Constants.FTP.ExcelFile, File.ReadAllBytes(ExcelFilepath));
            FTP.uploadFile(Constants.FTP.PptFile, File.ReadAllBytes(PresentationFilepath));

        }
    }
}
