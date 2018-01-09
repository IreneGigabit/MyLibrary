using System;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TestConsole {
	class WordXml {
		private static string CurrDir = System.Environment.CurrentDirectory;
		static string templateFile = CurrDir + @"\商標註冊申請書.docx";
		static string outputFile = CurrDir + @"\new.xml";

		static void Main(string[] args) {
			//createXML();
			writeDOCX();
		}

		#region 建立word(xml)
		public static void createXML() {
			// 建立 WordprocessingDocument 類別，透過 WordprocessingDocument 類別中的 Create 方法建立 Word 文件
			using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(outputFile, WordprocessingDocumentType.Document)) {
				// 建立 MainDocumentPart 類別物件 mainPart，加入主文件部分
				MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
				// 實例化 Document(w) 部分
				mainPart.Document = new Document();
				// 建立 Body 類別物件，於加入 Doucment(w) 中加入 Body 內文
				Body body = mainPart.Document.AppendChild(new Body());
				// 建立 Paragraph 類別物件，於 Body 本文中加入段落 Paragraph(p)
				Paragraph paragraph = body.AppendChild(new Paragraph());
				// 建立 Run 類別物件，於 段落 Paragraph(p) 中加入文字屬性 Run(r) 範圍
				Run run = paragraph.AppendChild(new Run());
				// 在文字屬性 Run(r) 範圍中加入文字內容
				run.AppendChild(new Text("在 body 本文內容產生 text 文字"));
			}

			//Process.Start(outputFile);
		}
		#endregion

		#region 寫入word(docx)
		public static void writeDOCX() {
			WordprocessingDocument tempDoc = WordprocessingDocument.Open(templateFile, false);
			Paragraph p = tempDoc.MainDocumentPart.Document.Body.Elements<Paragraph>().First();
			OpenXmlElement cc=p.CloneNode(true);

			using (WordprocessingDocument document = WordprocessingDocument.Open(outputFile, true)) {
				// Assign a reference to the existing document body.
				Body body = document.MainDocumentPart.Document.Body;
    
				// Add new text.
				Paragraph para = body.AppendChild(new Paragraph());
				Run run = para.AppendChild(new Run());
				run.AppendChild(new Text("測試" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")));

				body.AppendChild(cc);

				document.Close();
			}
			tempDoc.Dispose();
		}
		#endregion

	}

}
