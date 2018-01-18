using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Diagnostics;
using System.IO;
using MyLibrary;

namespace TestConsole {
	class WordXml {
		private static string CurrDir = System.Environment.CurrentDirectory;
		static string templateFile = CurrDir + @"\[團體標章註冊申請書].docx";
		static string outputFile = CurrDir + @"\[團體標章註冊申請書]-NT66824.docx";
		//static string templateFile = CurrDir + @"\img.docx";
		//static string outputFile = CurrDir + @"\img-new.doc";
		static string imgFile = CurrDir + @"\66824.jpg";

		static void Main(string[] args) {
			//createXML();
			//writeDOCX();
			//readTag();
			cloneDoc();
			//imageDoc();
			Process.Start(outputFile);
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

		#region word插入圖片
		private static void imageDoc() {
			File.Copy(templateFile, outputFile,true);
			using (WordprocessingDocument document = WordprocessingDocument.Open(outputFile, true)) {
				Body body = document.MainDocumentPart.Document.Body;
				body.AppendChild(new Paragraph(GenerateImageRun(document, new ImageFile(CurrDir + @"\66824.jpg"))));
				var cat2Img = new ImageFile(CurrDir + @"\66824.jpg")
				{
					Width = 8,
					Height = 8
				};
				var imgRun = GenerateImageRun(document, cat2Img);
				body.AppendChild(new Paragraph(imgRun));
			}
		}
		#endregion

		#region 複製範本產生新檔
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
				Console.WriteLine(pfoot1.ToString());
				IEnumerable<SectionProperties> sectPrs = outDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>();
				foreach (SectionProperties sectPr in sectPrs) {
					Console.WriteLine("sectPr ..!!");
				}

				outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SdtElement>();
				outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<Paragraph>();
				outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SectionProperties>();

				Body body = outDoc.MainDocumentPart.Document.Body;

				body.Append(copyTag2(tempDoc, "title"));
				//body.AppendChild(new Paragraph());//空白行
				copyTag3(tempDoc,outDoc,  "Block1");
				//body.Append(copyTag2(tempDoc, "Block1"));
				PasteBookmarkText(outDoc.MainDocumentPart, "seq_no", "NT66824(20180111)");
				PasteBookmarkText(outDoc.MainDocumentPart, "appl_name", "FE9測試");
				PasteBookmarkText(outDoc.MainDocumentPart, "color", "彩色");
				body.AppendChild(new Paragraph(GenerateImageRun(outDoc, new ImageFile(imgFile))));
				body.Append(copyTag2(tempDoc, "b_apcust"));//申請人
				PasteBookmarkText(outDoc.MainDocumentPart, "apcust_country", "TW中華民國");
				PasteBookmarkText(outDoc.MainDocumentPart, "apcust_cname", "英業達股份有限公司");
				PasteBookmarkText(outDoc.MainDocumentPart, "apcust_ename", "INVENTEC CORPORATION");
				body.Append(copyTag2(tempDoc, "b_agent"));//代理人
				PasteBookmarkText(outDoc.MainDocumentPart, "agt1_name", "高,玉駿");
				PasteBookmarkText(outDoc.MainDocumentPart, "agt2_name", "楊,祺雄");
				body.Append(copyTag2(tempDoc, "b_content"));//表彰內容
				PasteBookmarkText(outDoc.MainDocumentPart, "good_name", "英業達股份有限公司");
				body.Append(copyTag2(tempDoc, "b_fees"));//繳費資訊
				PasteBookmarkText(outDoc.MainDocumentPart, "pay_fees", "4700");
				PasteBookmarkText(outDoc.MainDocumentPart, "rectitle_name", "英業達股份有限公司");
				body.Append(copyTag2(tempDoc, "b_attach"));//附送書件
				body.Append(copyTag2(tempDoc, "b_statment"));//聲明內容
				body.AppendChild(new Paragraph(new ParagraphProperties(foot1)));//頁尾+換頁

				//基本資料表
				body.Append(copyTag2(tempDoc, "base_title"));//抬頭
				body.Append(copyTag2(tempDoc, "base_apcust"));//申請人
				PasteBookmarkText(outDoc.MainDocumentPart, "apcountry", "TW中華民國");
				PasteBookmarkText(outDoc.MainDocumentPart, "apclass", "法人公司機關學校");
				PasteBookmarkText(outDoc.MainDocumentPart, "apcust_no", "04322046");
				PasteBookmarkText(outDoc.MainDocumentPart, "cname_title", "中文名稱");
				PasteBookmarkText(outDoc.MainDocumentPart, "ap_cname", "英業達股份有限公司");
				PasteBookmarkText(outDoc.MainDocumentPart, "ename_title", "英文名稱");
				PasteBookmarkText(outDoc.MainDocumentPart, "ap_ename", "INVENTEC CORPORATION");
				PasteBookmarkText(outDoc.MainDocumentPart, "ap_live_country", "TW中華民國");
				PasteBookmarkText(outDoc.MainDocumentPart, "ap_zip", "840");
				PasteBookmarkText(outDoc.MainDocumentPart, "ap_addr", "高雄市大樹區學城路1段9、13、15、17、19、21、23號");
				PasteBookmarkText(outDoc.MainDocumentPart, "ap_crep", "堃峯");
				PasteBookmarkText(outDoc.MainDocumentPart, "ap_erep", "Lee &Richard");
				body.Append(copyTag4(tempDoc, "base_agent"));//代理人
				PasteBookmarkText(outDoc.MainDocumentPart, "agt_id1", "B100379440");
				PasteBookmarkText(outDoc.MainDocumentPart, "agt_name1", "高,玉駿");
				PasteBookmarkText(outDoc.MainDocumentPart, "agt_zip1", "105");
				PasteBookmarkText(outDoc.MainDocumentPart, "agt_addr1", "高,玉駿");
				PasteBookmarkText(outDoc.MainDocumentPart, "agatt_tel1", "02-77028299#261");
				PasteBookmarkText(outDoc.MainDocumentPart, "agatt_fax1", "02-77028289");
				PasteBookmarkText(outDoc.MainDocumentPart, "agt_id2", "M120741174");
				PasteBookmarkText(outDoc.MainDocumentPart, "agt_name2", "楊,祺雄");
				PasteBookmarkText(outDoc.MainDocumentPart, "agt_zip2", "105");
				PasteBookmarkText(outDoc.MainDocumentPart, "agt_addr2", "高,玉駿");
				PasteBookmarkText(outDoc.MainDocumentPart, "agatt_tel2", "02-77028299#261");
				PasteBookmarkText(outDoc.MainDocumentPart, "agatt_fax2", "02-77028289");

				//body.AppendChild(copyTag(tempDoc,outDoc, "title"));
				//body.AppendChild(new Paragraph());//空白行
				//body.AppendChild(copyTag(tempDoc, outDoc, "block1"));
				//body.AppendChild(new Paragraph());//空白行


				//body.AppendChild(new Paragraph( new Run( new LastRenderedPageBreak(), new Text("Last text on the page"))));//?
				//body.AppendChild(new Paragraph(new Run(new LastRenderedPageBreak(), foot1)));//?
				//body.AppendChild(new Paragraph(new ParagraphProperties(foot1)));//頁尾+換頁
				//body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));//換頁
				body.AppendChild(foot2);//頁尾
				//body.AppendChild(foot2);
			}
		}
		#endregion

		private static void copyTag3(WordprocessingDocument tmpDoc, WordprocessingDocument outDoc, string tagName) {
			Body body = outDoc.MainDocumentPart.Document.Body;
			Tag elementTag = tmpDoc.MainDocumentPart.RootElement.Descendants<Tag>()
			.Where(
				element => element.Val.Value.ToLower() == tagName.ToLower()
			).SingleOrDefault();

			if (elementTag != null) {
				Console.WriteLine("start find " + tagName + "..");
				var tagRuns = elementTag.Parent.Parent.Descendants<Paragraph>().ToArray();
				foreach (var tagRun in tagRuns) {
					body.AppendChild(tagRun.CloneNode(true));
				}
			}
		}
		private static Paragraph[] copyTag4(WordprocessingDocument doc, string tagName) {
			List<Paragraph> arrElement = new List<Paragraph>();
			Tag elementTag = doc.MainDocumentPart.RootElement.Descendants<Tag>()
			.Where(
				element => element.Val.Value.ToLower() == tagName.ToLower()
			).SingleOrDefault();

			Console.WriteLine("start find " + tagName + "..");
			if (elementTag != null) {
				Console.WriteLine("find " + tagName + "!!");

				SdtElement block = (SdtElement)elementTag.Parent.Parent;
				IEnumerable<Paragraph> tagRuns = block.Descendants<Paragraph>();
				foreach (Paragraph tagRun in tagRuns) {
					arrElement.Add(new Paragraph(new Run(new Text(tagRun.InnerText))));
				}
			}
			return arrElement.ToArray();
		}
		private static Paragraph[] copyTag2(WordprocessingDocument doc, string tagName) {
			List<Paragraph> arrElement = new List<Paragraph>();
			Tag elementTag = doc.MainDocumentPart.RootElement.Descendants<Tag>()
			.Where(
				element => element.Val.Value.ToLower() == tagName.ToLower()
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
					bookmarkStart.Remove();
					bookmarkEnd.Remove();
				}
			}
		}

		public static Run GenerateImageRun(WordprocessingDocument wordDoc, ImageFile img) {
			MainDocumentPart mainPart = wordDoc.MainDocumentPart;

			ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
			var relationshipId = mainPart.GetIdOfPart(imagePart);
			imagePart.FeedData(img.getDataStream());

			// Define the reference of the image.
			var element =
				 new Drawing(
					 new DW.Inline(
						 //Size of image, unit = EMU(English Metric Unit)
						 //1 cm = 360000 EMUs
						 new DW.Extent() { Cx = img.GetWidthInEMU(), Cy = img.GetHeightInEMU() },
						 new DW.EffectExtent()
						 {
							 LeftEdge = 0L,
							 TopEdge = 0L,
							 RightEdge = 0L,
							 BottomEdge = 0L
						 },
						 new DW.DocProperties()
						 {
							 Id = (UInt32Value)1U,
							 Name = img.ImageName
						 },
						 new DW.NonVisualGraphicFrameDrawingProperties(
							 new A.GraphicFrameLocks() { NoChangeAspect = true }),
						 new A.Graphic(
							 new A.GraphicData(
								 new PIC.Picture(
									 new PIC.NonVisualPictureProperties(
										 new PIC.NonVisualDrawingProperties()
										 {
											 Id = (UInt32Value)0U,
											 Name = img.FileName
										 },
										 new PIC.NonVisualPictureDrawingProperties()),
									 new PIC.BlipFill(
										 new A.Blip(
											 new A.BlipExtensionList(
												 new A.BlipExtension()
												 {
													 Uri =
														"{28A0092B-C50C-407E-A947-70E740481C1C}"
												 })
										 )
										 {
											 Embed = relationshipId,
											 CompressionState =
											 A.BlipCompressionValues.Print
										 },
										 new A.Stretch(
											 new A.FillRectangle())),
									 new PIC.ShapeProperties(
										 new A.Transform2D(
											 new A.Offset() { X = 0L, Y = 0L },
											 new A.Extents()
											 {
												 Cx = img.GetWidthInEMU(),
												 Cy = img.GetHeightInEMU()
											 }),
										 new A.PresetGeometry(
											 new A.AdjustValueList()
										 )
										 { Preset = A.ShapeTypeValues.Rectangle }))
							 )
							 { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
					 )
					 {
						 DistanceFromTop = (UInt32Value)0U,
						 DistanceFromBottom = (UInt32Value)0U,
						 DistanceFromLeft = (UInt32Value)0U,
						 DistanceFromRight = (UInt32Value)0U,
						// EditId = "50D07946"
					 });
			return new Run(element);
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

	}

}
