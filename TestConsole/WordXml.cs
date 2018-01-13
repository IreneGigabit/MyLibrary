using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections;
using System.IO;

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

			using (WordprocessingDocument outDoc = WordprocessingDocument.Open(outputFile, true)) {

				SectionProperties foot1 = new SectionProperties();
				SectionProperties foot2 = new SectionProperties();
				SectionProperties[] foot = outDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().ToArray();
				foot1 = (SectionProperties)foot[0].CloneNode(true);
				foot2 = (SectionProperties)foot[1].CloneNode(true);
				Paragraph pfoot1 = (Paragraph)foot[0].Parent.Parent.CloneNode(true);

				IEnumerable<SectionProperties> sectPrs = outDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>();
				foreach (SectionProperties sectPr in sectPrs) {
					Console.WriteLine("sectPr ..!!");
				}

				outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SdtElement>();
				outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<Paragraph>();
				outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SectionProperties>();

				Body body = outDoc.MainDocumentPart.Document.Body;

				body.Append(copyTag2(tempDoc, "title"));
				body.AppendChild(new Paragraph());//空白行
				body.Append(copyTag2(tempDoc, "Block1"));
				PasteBookmarkText(outDoc.MainDocumentPart, "seq_no", "NT-332838");
				body.AppendChild(new Paragraph());//空白行
				//body.AppendChild(copyTag(tempDoc,outDoc, "title"));
				//body.AppendChild(new Paragraph());//空白行
				//body.AppendChild(copyTag(tempDoc, outDoc, "block1"));
				//body.AppendChild(new Paragraph());//空白行


				//body.AppendChild(new Paragraph( new Run( new LastRenderedPageBreak(), new Text("Last text on the page"))));//?
				//body.AppendChild(new Paragraph(new Run(new LastRenderedPageBreak(), foot1)));//?
				//body.AppendChild(new Paragraph(new ParagraphProperties(foot1)));//頁尾+換頁
				body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));//換頁
				body.AppendChild(foot1);//頁尾
				//body.AppendChild(foot2);
			}
		}

		private static Paragraph[] copyTag2(WordprocessingDocument doc, string tagName) {
			List<Paragraph> arrElement = new List<Paragraph>();
			Tag elementTag = doc.MainDocumentPart.RootElement.Descendants<Tag>()
			.Where(
				element => element.Val.Value.ToLower()==tagName.ToLower()
			).SingleOrDefault();

			Console.WriteLine("start find " + tagName + "..");
			if (elementTag != null) {
				Console.WriteLine("find " + tagName + "!!");

				SdtElement block = (SdtElement)elementTag.Parent.Parent;
				IEnumerable<Paragraph> tagRuns = block.Descendants<Paragraph>();
				foreach (Paragraph tagRun in tagRuns) {
					arrElement.Add((Paragraph)tagRun.CloneNode(true));
					//return tagRun.CloneNode(true);
				}
			}
			return arrElement.ToArray();
		}

		private static void PasteBookmarkText(MainDocumentPart documentPart, string bookmarkName, string text) {
			IEnumerable<BookmarkEnd> bookMarkEnds = documentPart.RootElement.Descendants<BookmarkEnd>();
			foreach (BookmarkStart bookmarkStart in documentPart.RootElement.Descendants<BookmarkStart>()) {
				if (bookmarkStart.Name.Value.ToLower() == bookmarkName.ToLower()) {
					Console.WriteLine("find bookmark(" + bookmarkName + ")!!");
					string id = bookmarkStart.Id.Value;
					BookmarkEnd bookmarkEnd = bookMarkEnds.Where(i => i.Id.Value == id).First();

					//var bookmarkText = bookmarkEnd.NextSibling();
					Run bookmarkRun = bookmarkStart.NextSibling<Run>();
					if (bookmarkRun != null) {
						string[] txtArr = text.Split('\n');
						for (int i = 0; i < txtArr.Length; i++) {
							if (i == 0) {
								Console.WriteLine("insert single!!");
								bookmarkRun.GetFirstChild<Text>().Text = txtArr[i];
							} else {
								Console.WriteLine("insert multi!!");
								bookmarkRun.Append(new Break());
								bookmarkRun.Append(new Text(txtArr[i]));
							}
						}
						//bookmarkRun.GetFirstChild<Text>().Text = text;
						//bookmarkRun.Append(new Break());
						//bookmarkRun.Append(new Text("換行"));
					}
				}
			}
		}

		private static OpenXmlElement copyTag(WordprocessingDocument doc, WordprocessingDocument outdoc, string tagName) {
			Body body = outdoc.MainDocumentPart.Document.Body;
			Tag elementTag = doc.MainDocumentPart.RootElement.Descendants<Tag>()
			.Where(
				element => element.Val == tagName
			).SingleOrDefault();

			Console.WriteLine("start find " + tagName + "..");
			if (elementTag != null) {
				Console.WriteLine("find " + tagName + "!!");

				SdtElement block = (SdtElement)elementTag.Parent.Parent;
				IEnumerable<Paragraph> tagRuns = block.Descendants<Paragraph>();
				foreach (Paragraph tagRun in tagRuns) {
					body.AppendChild(tagRun.CloneNode(true));
					//return tagRun.CloneNode(true);
				}
			}
			return new Text();
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
