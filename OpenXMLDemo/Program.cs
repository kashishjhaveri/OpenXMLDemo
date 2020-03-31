using OpenXMLDemo.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;

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

            //OpenXMLUtilites.WordDoc.CreateWordprocessingDocument(DataLines);
            OpenXMLUtilites.Excel.CreateSpreadsheetWorkbook();
            OpenXMLUtilites.Excel.InsertData(DataLines);
        }
    }
}
    