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

            Models.OpenXMLUtilites.WordDoc.CreateWordprocessingDocument(HeaderData, DataLines);

            Models.OpenXMLUtilites.Excel.CreateSpreadsheetWorkbook();
            Models.OpenXMLUtilites.Excel.InsertData(HeaderData, DataLines);

            //OpenXMLUtilites.Presentation.CreatePresentation();

            string DocFilepath = $@"{Constants.Locations.DesktopPath}\{Constants.FTP.DocFile}";
            string ExcelFilepath = $@"{Constants.Locations.DesktopPath}\{Constants.FTP.ExcelFile}";
            
            FTP.uploadFile(Constants.FTP.DocFile, File.ReadAllBytes(DocFilepath));
            FTP.uploadFile(Constants.FTP.ExcelFile, File.ReadAllBytes(ExcelFilepath));

            //filepath = $@"{Constants.Locations.DesktopPath}\{Constants.FTP.PptFile}";
            //using (StreamReader sourceStream = new StreamReader(filepath))
            //{
            //    fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
            //}
            //FTP.uploadFile(Constants.FTP.PptFile, fileContents);
        }
    }
}
    