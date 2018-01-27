using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Web;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Drawing;

/// <summary>
/// Docx 操作類別(use OpenXml SDK)
/// </summary>
public class OpenXmlHelperOld {
	//protected SectionProperties[] footer = null;
	public WordprocessingDocument tempDoc = null;
	public WordprocessingDocument baseDoc = null;
	public WordprocessingDocument outDoc = null;
	protected MemoryStream tempMem = new MemoryStream();
	protected MemoryStream baseMem = new MemoryStream();
	protected MemoryStream outMem = new MemoryStream();
	protected Body outBody = null;

	public OpenXmlHelperOld() {
	}

	#region 關閉
	/// <summary>
	/// 關閉
	/// </summary>
	public void Dispose() {
		if (this.tempDoc != null) tempDoc.Dispose();
		if (this.outDoc != null) outDoc.Dispose();
		if (this.tempMem != null) tempMem.Close();
		if (this.outMem != null) outMem.Close();
		HttpContext.Current.Response.End();
	}
	#endregion

	#region 複製範本檔
	/// <summary>
	/// 複製範本檔
	/// </summary>
	/// <param name="templateFile">申請書範本檔名(實體路徑)</param>
	/// <param name="baseFile">基本資料表範本檔名(實體路徑)</param>
	/// <param name="cleanFlag">是否清空內容(只保留版面配置)</param>
	public void CloneFromFile(string templateFile, string baseFile, bool cleanFlag) {

		byte[] tempArray = File.ReadAllBytes(templateFile);
		byte[] baseArray = File.ReadAllBytes(baseFile);
		byte[] outArray = File.ReadAllBytes(templateFile);

		tempMem.Write(tempArray, 0, (int)tempArray.Length);
		baseMem.Write(baseArray, 0, (int)baseArray.Length);
		outMem.Write(outArray, 0, (int)outArray.Length);

		tempDoc = WordprocessingDocument.Open(tempMem, false);
		baseDoc = WordprocessingDocument.Open(baseMem, false);
		outDoc = WordprocessingDocument.Open(outMem, true);

		//footer = outDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().ToArray();

		//清空內容
		if (cleanFlag) {
			//outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SdtElement>();
			//outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<Paragraph>();
			//outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SectionProperties>();
			outDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
		}

		outBody = outDoc.MainDocumentPart.Document.Body;
	}
	#endregion

	#region 輸出檔案
	/// <summary>
	/// 輸出檔案
	/// </summary>
	public void Flush(string outputName) {
		outDoc.MainDocumentPart.Document.Save();
		outDoc.Close();
		//byte[] byteArray = outMem.ToArray();
		HttpContext.Current.Response.Clear();
		HttpContext.Current.Response.HeaderEncoding = System.Text.Encoding.GetEncoding("big5");
		HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment; filename=\"" + outputName + "\"");
		HttpContext.Current.Response.ContentType = "application/octet-stream";
		HttpContext.Current.Response.AddHeader("Content-Length", outMem.Length.ToString());
		HttpContext.Current.Response.BinaryWrite(outMem.ToArray());
		this.Dispose();
	}
	#endregion

	#region 增加段落文字
	/// <summary>
	/// 增加段落文字
	/// </summary>
	public void AddParagraph(string text) {
		outDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(new Text(text))));
	}
	#endregion

	#region 複製範本Block
	/// <summary>
	/// 複製範本Block
	/// </summary>
	public void CopyBlock(string blockName) {
		CopyBlock(tempDoc, blockName);
	}
	#endregion

	#region 複製範本Block(指定文件)
	/// <summary>
	/// 複製範本Block(指定文件)
	/// </summary>
	public void CopyBlock(WordprocessingDocument srcDoc, string blockName) {
		Tag elementTag = srcDoc.MainDocumentPart.RootElement.Descendants<Tag>()
		.Where(
			element => element.Val == blockName
		).SingleOrDefault();

		if (elementTag != null) {
			SdtElement block = (SdtElement)elementTag.Parent.Parent;
			IEnumerable<Paragraph> tagRuns = block.Descendants<Paragraph>();
			foreach (Paragraph tagRun in tagRuns) {
				outBody.Append((OpenXmlElement)tagRun.CloneNode(true));
			}
		}
	}
	#endregion

	#region 複製範本Block,回傳List
	/// <summary>
	/// 複製範本Block,回傳List
	/// </summary>
	private List<Paragraph> CopyBlockList(string blockName) {
		try {
			List<Paragraph> arrElement = new List<Paragraph>();
			Tag elementTag = tempDoc.MainDocumentPart.RootElement.Descendants<Tag>()
			.Where(
				element => element.Val.Value.ToLower() == blockName.ToLower()
			).SingleOrDefault();

			if (elementTag != null) {
				SdtElement block = (SdtElement)elementTag.Parent.Parent;
				IEnumerable<Paragraph> tagPars = block.Descendants<Paragraph>();
				foreach (Paragraph tagPar in tagPars) {
					arrElement.Add((Paragraph)tagPar.CloneNode(true));
				}
			}
			return arrElement;
		}
		catch (Exception ex) {
			throw new Exception("複製範本Block!!(" + blockName + ")", ex);
		}
	}
	#endregion

	#region 複製範本Block,並取代文字
	/// <summary>
	/// 複製範本Block,並取代文字
	/// </summary>
	public void CloneReplaceBlock(string blockName, string searchStr, string newStr) {
		try {
			List<Paragraph> pars = CopyBlockList(blockName);
			for (int i = 0; i < pars.Count; i++) {
				pars[i] = (new Paragraph(new Run(new Text(pars[i].InnerText.Replace(searchStr, newStr)))));
			}
			outBody.Append(pars.ToArray());
		}
		catch (Exception ex) {
			throw new Exception("複製範本Block錯誤!!(" + blockName + ")", ex);
		}
	}
	#endregion

	#region 取代書籤
	/// <summary>
	/// 取代書籤
	/// </summary>
	public void ReplaceBookmark(string bookmarkName, string text) {
		try {
			MainDocumentPart mainPart = outDoc.MainDocumentPart;
			IEnumerable<BookmarkEnd> bookMarkEnds = mainPart.RootElement.Descendants<BookmarkEnd>();
			foreach (BookmarkStart bookmarkStart in mainPart.RootElement.Descendants<BookmarkStart>()) {
				if (bookmarkStart.Name.Value.ToLower() == bookmarkName.ToLower()) {
					string id = bookmarkStart.Id.Value;
					//BookmarkEnd bookmarkEnd = bookMarkEnds.Where(i => i.Id.Value == id).First();
					BookmarkEnd bookmarkEnd = bookMarkEnds.Where(i => i.Id.Value == id).FirstOrDefault();

					////var bookmarkText = bookmarkEnd.NextSibling();
					//Run bookmarkRun = bookmarkStart.NextSibling<Run>();
					//if (bookmarkRun != null) {
					//	string[] txtArr = text.Split('\n');
					//	for (int i = 0; i < txtArr.Length; i++) {
					//		if (i == 0) {
					//			bookmarkRun.GetFirstChild<Text>().Text = txtArr[i];
					//		} else {
					//			bookmarkRun.Append(new Break());
					//			bookmarkRun.Append(new Text(txtArr[i]));
					//		}
					//	}
					//}
					Run bookmarkRun = bookmarkStart.NextSibling<Run>();
					if (bookmarkRun != null) {
						Run tempRun = bookmarkRun;
						string[] txtArr = text.Split('\n');
						for (int i = 0; i < txtArr.Length; i++) {
							if (i == 0) {
								bookmarkRun.GetFirstChild<Text>().Text = txtArr[i];
							} else {
								bookmarkRun.Append(new Break());
								bookmarkRun.Append(new Text(txtArr[i]));
							}
						}
						int j = 0;
						while (tempRun.NextSibling() != null && tempRun.NextSibling().GetType() != typeof(BookmarkEnd)) {
							j++;
							tempRun.NextSibling().Remove();
							if (j >= 20)
								break;
						}
					}
					bookmarkStart.Remove();
					if (bookmarkEnd != null) bookmarkEnd.Remove();
				}
			}
		}
		catch (Exception ex) {
			throw new Exception("取代書籤錯誤!!(" + bookmarkName + ")", ex);
		}
	}
	#endregion

	#region 複製範本頁尾
	/// <summary>
	/// 複製範本頁尾
	/// </summary>
	/// <param name="sourceDoc">複製來源</param>
	/// <param name="haveBreak">是否帶分節符號(新頁)</param>
	public void CopyPageFoot(WordprocessingDocument sourceDoc, bool haveBreak) {
		int index = 0;//取消index參數,只抓第1個

		string newRefId = string.Format("foot_{0}", Guid.NewGuid().ToString().Substring(0, 8));

		FooterReference[] footerSections = sourceDoc.MainDocumentPart.RootElement.Descendants<FooterReference>().ToArray();
		string srcRefId = footerSections[index].Id;
		footerSections[index].Id = newRefId;

		FooterPart elementFoot = sourceDoc.MainDocumentPart.FooterParts
		.Where(
			element => sourceDoc.MainDocumentPart.GetIdOfPart(element) == srcRefId
		).SingleOrDefault();
		outDoc.MainDocumentPart.AddPart(elementFoot, newRefId);

		if (haveBreak)
			outBody.AppendChild(new Paragraph(new ParagraphProperties(footerSections[index].Parent.CloneNode(true))));//頁尾+分節符號
		else
			outBody.AppendChild(footerSections[index].Parent.CloneNode(true));//頁尾
	}

	//private void CopyPageFoot(WordprocessingDocument sourceDoc, int index,bool haveBreak) {
	//	string refId0 = string.Format("foot_{0}", Guid.NewGuid().ToString().Substring(0, 8));
	//
	//	//IEnumerable<SectionProperties> sections0 = tempDoc.MainDocumentPart.Document.Body.Elements<SectionProperties>();
	//	SectionProperties[] elementSections = sourceDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().ToArray();
	//	SectionProperties bfoot0 = (SectionProperties)elementSections[index].CloneNode(true);
	//	FooterReference bid0 = bfoot0.GetFirstChild<FooterReference>();
	//	string srcRefId = bid0.Id;
	//	bid0.Id = refId0;
	//
	//	FooterPart elementFoot = sourceDoc.MainDocumentPart.FooterParts
	//	.Where(
	//		element => sourceDoc.MainDocumentPart.GetIdOfPart(element) == srcRefId
	//	).SingleOrDefault();
	//
	//	outDoc.MainDocumentPart.AddPart(elementFoot, refId0);
	//
	//	if (haveBreak)
	//		outBody.AppendChild(new Paragraph(new ParagraphProperties(bfoot0)));//頁尾+分節符號
	//	else
	//		outBody.AppendChild(bfoot0);//頁尾
	//}
	#endregion

	#region 插入圖片
	/// <summary>
	/// 插入圖片
	/// </summary>
	public void AppendImage(string imgFilePath) {
		ImageFile img = new ImageFile(imgFilePath);

		ImagePart imagePart = outDoc.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);
		string relationshipId = outDoc.MainDocumentPart.GetIdOfPart(imagePart);
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
									 ) { Preset = A.ShapeTypeValues.Rectangle }))
						 ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
				 )
				 {
					 DistanceFromTop = (UInt32Value)0U,
					 DistanceFromBottom = (UInt32Value)0U,
					 DistanceFromLeft = (UInt32Value)0U,
					 DistanceFromRight = (UInt32Value)0U,
					 //EditId = "50D07946"
				 });

		outDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
	}
	#endregion
}
