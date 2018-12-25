using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MyLibrary;
using System.IO;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TestConsole
{
    class WordTableTester
    {
        private static string CurrDir = System.Environment.CurrentDirectory;//Debug路徑
        static DirectoryInfo dir = new DirectoryInfo(CurrDir);
        static string BaseDir = dir.Parent.Parent.FullName;//專案路徑

        static string templateFile = BaseDir + @"\testDocument\英文Invoice_test.docx";
        static string outputFile = BaseDir + @"\OutReport\m1583-英文Invoice.docx";

        static void Main(string[] args)
        {
            OpenXmlHelper docx = new OpenXmlHelper();
            
            Dictionary<string, string> _tplFile = new Dictionary<string, string>();
            _tplFile.Add("invoice", templateFile);
            docx.CloneFromFile(_tplFile, true);

            docx.CopyBlock("b_title");
            docx.CopyBlock("b_item");
            docx.ReplaceBookmark("e_arcase", "xxxxxxx:");
            docx.CopyTable(2);
            docx.ReplaceBookmark("t_title", "Total:");
            docx.ReplaceBookmark("t_curr", "NTD");
            docx.ReplaceBookmark("t_total", "23,000.00");
            docx.CopyTable(2);
            docx.ReplaceBookmark("t_title", "");
            docx.ReplaceBookmark("t_curr", "USD");
            docx.ReplaceBookmark("t_total", "120.00");

            docx.CopyBlock("b_total3");

            docx.CopyPageHeader("invoice");//複製頁首
            docx.CopyPageFoot("invoice", false);//複製頁尾/邊界

            docx.SaveTo(outputFile);

            Process.Start(outputFile);
            Console.ReadLine();
        }
    }
}
