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
	class WordOpenXml
	{
		private static string CurrDir = System.Environment.CurrentDirectory;//Debug路徑
		static DirectoryInfo dir = new DirectoryInfo(CurrDir);
		static string BaseDir = dir.Parent.Parent.FullName;//專案路徑

		static string templateFile = BaseDir + @"\testDocument\[團體標章註冊申請書].docx";
		static string baseFile = BaseDir + @"\testDocument\00基本資料表.docx";
		static string outputFile = BaseDir + @"\OutReport\[團體標章註冊申請書]-NT66824.docx";

		static void Main(string[] args) {
			OpenXmlHelper docx = new OpenXmlHelper();
			Dictionary<string, string> _TemplateFileList = new Dictionary<string, string>();
			_TemplateFileList.Add("apply", templateFile);
			_TemplateFileList.Add("base", baseFile);
			docx.CloneFromFile(_TemplateFileList, true);

			docx.CopyBlock("titl");
			docx.CopyBlock("block1");
			docx.CopyBlock("b_apcust");
			docx.CopyBlock("b_agent");
			docx.CopyBlock("b_content");
			docx.CopyBlock("b_fees");
			docx.CopyBlock("b_attach");
			docx.CopyBlock("b_sign");
			docx.CopyPageFoot("apply", true);
			docx.CopyPageFoot("base", false);
			docx.SetPageSize((decimal)21,(decimal)29.7);

			docx.SaveTo(outputFile);

			Process.Start(outputFile);
			Console.ReadLine();
		}
	}
}
