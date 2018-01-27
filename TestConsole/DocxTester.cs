using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TestConsole {
	class DocxTester {
		private static string CurrDir = System.Environment.CurrentDirectory;//Debug路徑
		static DirectoryInfo dir = new DirectoryInfo(CurrDir);
		static string BaseDir = dir.Parent.Parent.FullName;//專案路徑

		static string templateFile = BaseDir + @"\testDocument\01發明專利申請書.docx";
		static string outputFile = BaseDir + @"\testDocument\01發明專利申請書-NT66824.docx";

		static void Main(string[] args) {
			Dictionary<string, string> _TemplateFileList = new Dictionary<string, string>();
			_TemplateFileList.Add("apply", BaseDir + @"\testDocument\01發明專利申請書.docx");
			_TemplateFileList.Add("base", BaseDir + @"\testDocument\00基本資料表.docx");
			OpenXmlHelper docxRpt = new OpenXmlHelper();
			docxRpt.CloneFromFile(_TemplateFileList, true);

			Dictionary<int, Paragraph> attach = docxRpt.CopyBlockDict("b_attach");
			foreach(var line in attach ){
				Console.WriteLine(line.Key + "→" + line.Value.InnerText);
			}
			attach.Remove(13);
			foreach (var line in attach) {
				docxRpt.AddParagraph(line.Value);
				Console.WriteLine(line.Key + "→" + line.Value.InnerText);
			}

			docxRpt.SaveTo(BaseDir + @"\OutReport\01發明專利申請書" + DateTime.Now.ToString("yyyyMMdd") + ".docx");
			Console.ReadLine();
		}
	}
}
