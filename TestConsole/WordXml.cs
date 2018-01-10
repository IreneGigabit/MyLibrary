using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TestConsole {
	class WordXml {
		private static string CurrDir = System.Environment.CurrentDirectory;
		static string templateFile = CurrDir + @"\商標註冊申請書.docx";
		static string outputFile = CurrDir + @"\商標註冊申請書 - 複製.docx";

		static void Main(string[] args) {
			//createXML();
			//writeDOCX();
			//readTag();
			cloneDoc();
			Console.ReadLine();
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

		private static void cloneDoc() {
			System.IO.File.Copy(templateFile, outputFile, true);

			WordprocessingDocument tempDoc = WordprocessingDocument.Open(templateFile, false);

			using (WordprocessingDocument document = WordprocessingDocument.Open(outputFile, true)) {

				Paragraph foot1 =new Paragraph();
				Paragraph foot2 = new Paragraph();
				SectionProperties[] foot = document.MainDocumentPart.RootElement.Descendants<SectionProperties>().ToArray();
				foot1.AppendChild(foot[0]);
				//foot2.Append(foot[1]);

				IEnumerable<SectionProperties> sectPrs = document.MainDocumentPart.RootElement.Descendants<SectionProperties>();
				foreach (SectionProperties sectPr in sectPrs) {
					Console.WriteLine("sectPr ..!!");
				}

				document.MainDocumentPart.Document.Body.RemoveAllChildren<SdtElement>();
				document.MainDocumentPart.Document.Body.RemoveAllChildren<Paragraph>();
				//document.MainDocumentPart.Document.Body.RemoveAllChildren<SectionProperties>();

				Body body = document.MainDocumentPart.Document.Body;

				string tagName = "prior";

				Tag elementTag = tempDoc.MainDocumentPart.RootElement.Descendants<Tag>()
				.Where(
					element => element.Val == tagName
				).SingleOrDefault();

				Console.WriteLine("start find " + tagName + "..");
				if (elementTag != null) {
					Console.WriteLine("find " + tagName + "!!");

					SdtElement block = (SdtElement)elementTag.Parent.Parent;
					IEnumerable<Paragraph> tagRuns = block.Descendants<Paragraph>();
					foreach (Paragraph tagRun in tagRuns) {
						Console.WriteLine("find Paragraph(" + tagName + ")!!");
						body.AppendChild(tagRun.CloneNode(true));
					}

					foreach (Paragraph tagRun in tagRuns) {
						Console.WriteLine("find Paragraph(" + tagName + ")!!");
						body.AppendChild(tagRun.CloneNode(true));
					}

				}
			}
		}

		private static void readTag() {
			WordprocessingDocument tempDoc = WordprocessingDocument.Open(templateFile, false);

			using (WordprocessingDocument document = WordprocessingDocument.Create(outputFile, WordprocessingDocumentType.Document)) {
				// 建立 MainDocumentPart 類別物件 mainPart，加入主文件部分 
				MainDocumentPart mainPart = document.AddMainDocumentPart();
				// 實例化 Document(w) 部分
				mainPart.Document = new Document();
				//Part
				mainPart.AddPart(tempDoc.MainDocumentPart.NumberingDefinitionsPart);


				Body body = mainPart.Document.AppendChild(new Body());
				string tagName = "prior";
	
				Tag elementTag = tempDoc.MainDocumentPart.RootElement.Descendants<Tag>()
				.Where(
					element => element.Val == tagName
				).SingleOrDefault();

				Console.WriteLine("start find " + tagName + "..");
				if (elementTag != null) {
					Console.WriteLine("find " + tagName + "!!");

					SdtElement block = (SdtElement)elementTag.Parent.Parent;
					IEnumerable<Paragraph> tagRuns = block.Descendants<Paragraph>();
					foreach (Paragraph tagRun in tagRuns) {
						Console.WriteLine("find Paragraph(" + tagName + ")!!");
						body.AppendChild(tagRun.CloneNode(true));
					}

					foreach (Paragraph tagRun in tagRuns) {
						Console.WriteLine("find Paragraph(" + tagName + ")!!");
						body.AppendChild(tagRun.CloneNode(true));
					}

				}

				document.Close();
			}
			tempDoc.Dispose();
		}

		private static void PasteTagText(MainDocumentPart documentPart, string tagName, string text) {
			Tag elementTag = documentPart.RootElement.Descendants<Tag>()
			.Where(
				element => element.Val == tagName
			).SingleOrDefault();

			if (elementTag != null) {
				SdtElement block = (SdtElement)elementTag.Parent.Parent;
				//Run tagRun = block.Descendants<Run>().FirstOrDefault();//.SingleOrDefault();
				IEnumerable<Run> tagRuns = block.Descendants<Run>();
				foreach (Run tagRun in tagRuns) {
					if (tagRun.GetFirstChild<Text>() != null) {
						string[] txtArr = text.Split('\n');
						for (int i = 0; i < txtArr.Length; i++) {
							if (i == 0) {
								tagRun.GetFirstChild<Text>().Text = txtArr[i];
							} else {
								tagRun.Append(new Break());
								tagRun.Append(new Text(txtArr[i]));
							}
						}
						break;
					}
				}
			}
		}
	}

}
