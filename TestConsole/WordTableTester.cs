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

        static string templateFile = BaseDir + @"\testDocument\催延展定稿_imt2e3_g2_form.docx";
        static string outputFile = BaseDir + @"\OutReport\催延展.docx";

        static void Main(string[] args)
        {
            OpenXmlHelper docx = new OpenXmlHelper();
            
            Dictionary<string, string> _tplFile = new Dictionary<string, string>();
            _tplFile.Add("invoice", templateFile);
            docx.CloneFromFile(_tplFile, true);

            docx.CopyBlock("b_all");
			//TableRow sTr = docx.GetTemplateTable("invoice", 0).GetTable(0).GetRow(1);
			docx.GetTemplateTable("invoice", 0).GetTable(0).NewRow();

			//docx.AppendRow("invoice", 0);

			docx.CopyPageFoot("invoice", false);//複製頁尾/邊界

            docx.SaveTo(outputFile);

            Process.Start(outputFile);
            Console.ReadLine();
        }
    }
}
