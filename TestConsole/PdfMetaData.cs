using System;
using System.Collections.Generic;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Diagnostics;

namespace TestConsole
{
	class PdfMetaData
    {
        private static string CurrDir = System.Environment.CurrentDirectory;//Debug路徑
        static DirectoryInfo dir = new DirectoryInfo(CurrDir);
        static string BaseDir = dir.Parent.Parent.FullName;//專案路徑

        static string templateFile = BaseDir + @"\testDocument\PI-66686-JA-SPCL-20180508_DESCRIPTION.pdf";
		static string txtFile = BaseDir + @"\testDocument\XmlData.txt";

		static void Main(string[] args) {
			PdfReader reader = new PdfReader(templateFile);
			string s = reader.Info["XmlData"];
			s = s.Replace("><", ">\r\n<");

			using (FileStream fs = new FileStream(txtFile, FileMode.Create)) {
				StreamWriter sw = new StreamWriter(fs);
				sw.Write(s);
				sw.Flush();
				sw.Close();
			}

			Process.Start(txtFile);
		}
	}
}
