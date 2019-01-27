using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Web;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Word = Microsoft.Office.Interop.Word;

//using System.Drawing;

/// <summary>
/// Docx 操作類別(use OpenXml SDK)
/// TODO:
/// 紙張邊界
/// </summary>
public class OpenXmlHelper
{
	protected WordprocessingDocument outDoc = null;
	protected MemoryStream outMem = new MemoryStream();
	protected Body outBody = null;
	Dictionary<string, WordprocessingDocument> tplDoc = new Dictionary<string, WordprocessingDocument>();
	protected string defTplDocName = "";
	Dictionary<string, MemoryStream> tplMem = new Dictionary<string, MemoryStream>();

	public OpenXmlHelper() {
	}

	#region 關閉 +void Dispose()
	/// <summary>
	/// 關閉
	/// </summary>
	public void Dispose() {
		if (this.outDoc != null) outDoc.Dispose();
		if (this.outMem != null) outMem.Close();

		foreach (var item in tplDoc) {
			item.Value.Dispose();
		}
		foreach (var item in tplMem) {
			item.Value.Close();
			item.Value.Dispose();
		}
		//微軟KB 312629https://support.microsoft.com/en-us/help/312629/prb-threadabortexception-occurs-if-you-use-response-end--response-redi
		//Response.End、Server.Transfer、Response.Redirect被呼叫時，會觸發ThreadAbortException，因此要改用CompleteRequest()
		//HttpContext.Current.Response.End();
		if (HttpContext.Current != null) HttpContext.Current.ApplicationInstance.CompleteRequest();
	}
	#endregion

	#region 建立空白檔案 +void Create()
	/// <summary>
	/// 建立空白檔案
	/// </summary>
	public void Create() {
		outMem = new MemoryStream();
		outDoc = WordprocessingDocument.Create(outMem, WordprocessingDocumentType.Document);
		MainDocumentPart mainPart = outDoc.AddMainDocumentPart();
		mainPart.Document = new Document();
		outBody = mainPart.Document.AppendChild(new Body());
	}
	#endregion

	#region 複製範本檔 +void CloneFromFile(Dictionary<string, string> templateList, bool cleanFlag)
	/// <summary>
	/// 複製範本檔
	/// </summary>
	/// <param name="templateList">範本＜別名,檔名(實體路徑)＞</param>
	public void CloneFromFile(Dictionary<string, string> templateList, bool cleanFlag) {
		foreach (var x in templateList.Select((Entry, Index) => new { Entry, Index })) {
			if (x.Index == 0) {
				byte[] outArray = File.ReadAllBytes(x.Entry.Value);
				outMem.Write(outArray, 0, (int)outArray.Length);
				outDoc = WordprocessingDocument.Open(outMem, true);
				defTplDocName = x.Entry.Key;
			}

			byte[] tplArray = File.ReadAllBytes(x.Entry.Value);
			tplMem.Add(x.Entry.Key, new MemoryStream());
			tplMem[x.Entry.Key].Write(tplArray, 0, (int)tplArray.Length);
			tplDoc.Add(x.Entry.Key, WordprocessingDocument.Open(tplMem[x.Entry.Key], false));
		}

		//清空輸出檔內容
		if (cleanFlag) {
			//outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SdtElement>();
			//outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<Paragraph>();
			//outDoc.MainDocumentPart.Document.Body.RemoveAllChildren<SectionProperties>();
			outDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
			//List<OpenXmlElement> child = outDoc.MainDocumentPart.Document.Body.ChildElements.TakeWhile(d => d.GetType() != typeof(SectionProperties)).ToList();
			//child.ForEach(c => c.Remove());
		}

		outBody = outDoc.MainDocumentPart.Document.Body;
	}
	#endregion

	#region 將指定檔案轉為pdf +static void ToPdf(string sourceFile, string outputFile)
	/// <summary>
	/// 將指定檔案轉為pdf
	/// </summary>
	/// <param name="sourceFile">來源檔</param>
	/// <param name="outputFile">目的檔</param>
	public static void ConvertToPDF(string sourceFile, string outputFile) {
		object srcFile = sourceFile;
		object destFile = outputFile;

		//word用的常數值==
		object wdFormatPDF = Word.WdSaveFormat.wdFormatPDF;
		object oFalse = false;
		object oTrue = true;
		object oMissing = System.Reflection.Missing.Value;
		//===============

		Word._Application wordApp = new Word.Application();

		Word._Document myDoc = wordApp.Documents.Open(ref srcFile, ref oMissing, ref oTrue, ref oMissing,
							ref oMissing, ref oMissing, ref oMissing, ref oMissing,
							ref oMissing, ref oMissing, ref oMissing, ref oFalse,
							ref oMissing, ref oMissing, ref oMissing, ref oMissing);
		myDoc.Activate();
		try {
			wordApp.ActiveDocument.SaveAs(ref destFile, ref wdFormatPDF, ref oMissing, ref oMissing,
				ref oMissing, ref oMissing, ref oMissing, ref oMissing,
				ref oMissing, ref oMissing, ref oMissing, ref oMissing,
				ref oMissing, ref oMissing, ref oMissing, ref oMissing);
		}
		finally {
			wordApp.ActiveDocument.Close(ref oMissing, ref oMissing, ref oMissing);
			wordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
			if (myDoc != null)
				System.Runtime.InteropServices.Marshal.ReleaseComObject(myDoc);
			if (wordApp != null)
				System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
			myDoc = null;
			wordApp = null;
			GC.Collect();
		}
	}
	#endregion

	#region 合併檔案後存至指定路徑 +static void MergeFile(List<Source> sourceFile, string outputFile)
	/// <summary>
	/// 合併檔案後存至指定路徑
	/// </summary>
	/// <param name="sourceFile">要合併的檔案清單(以第一個檔為母檔)</param>
	/// <param name="outputFile">合併後的輸出檔案</param>
	public static void MergeFile(List<Docx> sourceFile, string outputFile) {
		byte[] byteArray = File.ReadAllBytes(sourceFile[0].FileName);
		using (MemoryStream stream = new MemoryStream()) {
			stream.Write(byteArray, 0, (int)byteArray.Length);
			using (WordprocessingDocument myDoc = WordprocessingDocument.Open(stream, true)) {
				MainDocumentPart mainPart = myDoc.MainDocumentPart;
				/*
				PageMargin pgMar = mainPart.Document.Descendants<PageMargin>().FirstOrDefault();
				Int32Value tMargin = 0;
				Int32Value bMargin = 0;
				UInt32Value lMargin = 0;
				UInt32Value rMargin = 0;
				if (pgMar != null) {
					tMargin = pgMar.Top;
					bMargin = pgMar.Bottom;
					lMargin = pgMar.Left;
					rMargin = pgMar.Right;
				}*/

				if (sourceFile.Count > 1) {
					for (var i = 1; i < sourceFile.Count; i++) {
						
						AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML);
						string altChunkId = mainPart.GetIdOfPart(chunk);
						
						using (FileStream fileStream = new FileStream(sourceFile[i].FileName, FileMode.Open)) {
							chunk.FeedData(fileStream);
						}
						AltChunk altChunk = new AltChunk();
						altChunk.Id = altChunkId;

						//mainPart.Document.Body.InsertAfter(altChunk, mainPart.Document.Body.Elements<Paragraph>().First());
						//mainPart.Document.Body.InsertAfter(new Paragraph(altChunk), mainPart.Document.Body.Elements<Paragraph>().First());
						//mainPart.Document.Body.Append(altChunk);
						mainPart.Document.Body.AppendChild(altChunk);

						/*
						using (WordprocessingDocument srcDoc = WordprocessingDocument.Open(sourceFile[i].FileName, false)) {
							//List<OpenXmlElement> contents = srcDoc.MainDocumentPart.Document.Body.ChildElements.Where(x => x.GetType() != typeof(SectionProperties)).ToList();
							List<OpenXmlElement> contents = srcDoc.MainDocumentPart.Document.Body.ChildElements.ToList();
							contents.ForEach(t =>
							{
								mainPart.Document.Body.InsertAfter(t.CloneNode(true), mainPart.Document.Body.Elements<Paragraph>().Last());
								//mainPart.Document.Body.InsertBefore(t.CloneNode(true), mainPart.Document.Body.Elements<Paragraph>().First());
								//mainPart.Document.Body.Elements<SectionProperties>().First().PrependChild(t.CloneNode(true));
							});
						}*/

						if (sourceFile[i].BeforeBreak) {
							//mainPart.Document.Body.InsertAfter(new Break() { Type = BreakValues.Page }, mainPart.Document.Body.Elements<Paragraph>().First());
						}
					}
				}

				mainPart.Document.Save();
			}
			File.WriteAllBytes(outputFile, stream.ToArray());
		}
	}
	#endregion

	#region 輸出檔案(memory) +void Flush(string outputName)
	/// <summary>
	/// 輸出檔案(memory)
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
		//this.Dispose();
	}
	#endregion

	#region 另存檔案 +void SaveTo(string outputPath)
	/// <summary>
	/// 另存檔案
	/// </summary>
	public void SaveTo(string outputPath) {
		outDoc.MainDocumentPart.Document.Save();
		outDoc.Close();
		using (FileStream fileStream = new FileStream(outputPath, FileMode.Create)) {
			outMem.Position = 0;
			outMem.WriteTo(fileStream);
		}
		this.Dispose();
	}
	#endregion

	#region 複製範本Block,回傳List +List<OpenXmlElement> CopyBlockList(string blockName)
	/// <summary>
	/// 複製範本Block,回傳List
	/// </summary>
	public List<OpenXmlElement> CopyBlockList(string blockName) {
		return CopyBlockList(defTplDocName, blockName);
	}
	#endregion

	#region 複製範本Block,回傳List +List<OpenXmlElement> CopyBlockList(string srcDocName, string blockName)
	/// <summary>
	/// 複製範本Block,回傳List
	/// </summary>
	/// <param name="srcDocName">來源範本別名</param>
	public List<OpenXmlElement> CopyBlockList(string srcDocName, string blockName) {
		try {
			WordprocessingDocument srcDoc = tplDoc[srcDocName];
			List<OpenXmlElement> arrElement = new List<OpenXmlElement>();
			Tag elementTag = srcDoc.MainDocumentPart.RootElement.Descendants<Tag>()
			.Where(
				element => element.Val.Value.ToLower() == blockName.ToLower()
			).SingleOrDefault();

			if (elementTag != null) {
				SdtElement block = (SdtElement)elementTag.Parent.Parent;
				SdtContentBlock blockCont = block.Descendants<SdtContentBlock>().FirstOrDefault();
				if (blockCont != null) {
					IEnumerable<OpenXmlElement> childs = blockCont.ChildElements;
					//IEnumerable<OpenXmlElement> childs = block.ChildElements.Where(i => i.GetType() == typeof(SdtContentBlock));//SdtEndCharProperties //SdtProperties
					//IEnumerable<OpenXmlElement> childs = (OpenXmlElement)block.Descendants<SdtContentBlock>();
					foreach (var item in childs) {
						arrElement.Add(item.CloneNode(true));
					}
				}
			}
			return arrElement;
		}
		catch (Exception ex) {
			throw new Exception("複製範本Block!!(" + blockName + ")", ex);
		}
	}
	#endregion

	#region 複製範本Block +void CopyBlock(string blockName)
	/// <summary>
	/// 複製範本Block
	/// </summary>
	public void CopyBlock(string blockName) {
		foreach (var par in CopyBlockList(blockName)) {
			outBody.Append(par.CloneNode(true));
		}
	}
	#endregion

	#region 複製範本Block +void CopyBlock(string srcDocName, string blockName)
	/// <summary>
	/// 複製範本Block(指定來源)
	/// </summary>
	/// <param name="srcDocName">來源範本別名</param>
	public void CopyBlock(string srcDocName, string blockName) {
		foreach (var par in CopyBlockList(srcDocName, blockName)) {
			outBody.Append(par.CloneNode(true));
		}
	}
	#endregion

	#region 複製範本Block,回傳Dictionary +Dictionary<int, OpenXmlElement> CopyBlockDict(string blockName)
	/// <summary>
	/// 複製範本Block,回傳Dictionary
	/// </summary>
	public Dictionary<int, OpenXmlElement> CopyBlockDict(string blockName) {
		return CopyBlockDict(defTplDocName, blockName);
	}
	#endregion

	#region 複製範本Block,回傳Dictionary +Dictionary<int, OpenXmlElement> CopyBlockDict(string srcDocName, string blockName)
	/// <summary>
	/// 複製範本Block,回傳Dictionary
	/// </summary>
	/// <param name="srcDocName">來源範本別名</param>
	public Dictionary<int, OpenXmlElement> CopyBlockDict(string srcDocName, string blockName) {
		try {
			WordprocessingDocument srcDoc = tplDoc[srcDocName];
			Dictionary<int, OpenXmlElement> dictElement = new Dictionary<int, OpenXmlElement>();

			foreach (var x in CopyBlockList(srcDocName, blockName).Select((Entry, Index) => new { Entry, Index })) {
				dictElement.Add(x.Index + 1, x.Entry);
			}
			return dictElement;
		}
		catch (Exception ex) {
			throw new Exception("複製範本Block!!(" + blockName + ")", ex);
		}
	}
	#endregion

	#region 複製範本Block,並取代文字後再貼上 +void CopyReplaceBlock(string blockName, string searchStr, string newStr)
	/// <summary>
	/// 複製範本Block,並取代文字後再貼上
	/// </summary>
	public void CopyReplaceBlock(string blockName, string searchStr, string newStr) {
		CopyReplaceBlock(defTplDocName, blockName, new Dictionary<string, string>() { { searchStr, newStr } });
	}
	#endregion

	#region 複製範本Block,並取代文字後再貼上 +void CopyReplaceBlock(string srcDocName, string blockName, string searchStr, string newStr)
	/// <summary>
	/// 複製範本Block,並取代文字後再貼上(指定來源)
	/// </summary>
	/// <param name="srcDocName">來源範本別名</param>
	public void CopyReplaceBlock(string srcDocName, string blockName, string searchStr, string newStr) {
		CopyReplaceBlock(srcDocName, blockName, new Dictionary<string, string>() { { searchStr, newStr } });
	}
	#endregion

	#region 複製範本Block,並取代文字後再貼上 +void CopyReplaceBlock(string blockName, Dictionary<string, string> mappingDic)
	/// <summary>
	/// 複製範本Block,並取代文字後再貼上
	/// </summary>
	public void CopyReplaceBlock(string blockName, Dictionary<string, string> mappingDic) {
		CopyReplaceBlock(defTplDocName, blockName, mappingDic);
	}
	#endregion

	#region 複製範本Block,並取代文字後再貼上 +void CopyReplaceBlock(string srcDocName, string blockName, Dictionary<string, string> mappingDic)
	/// <summary>
	/// 複製範本Block,並取代文字後再貼上(指定來源)
	/// </summary>
	/// <param name="srcDocName">來源範本別名</param>
	public void CopyReplaceBlock(string srcDocName, string blockName, Dictionary<string, string> mappingDic) {
		int i = 0;
		try {
			List<OpenXmlElement> pars = CopyBlockList(srcDocName, blockName);
			for (i = 0; i < pars.Count; i++) {
				string oldInnerText = pars[i].InnerText;
				string tmpInnerText = pars[i].InnerText;
				foreach (var item in mappingDic) {
					tmpInnerText = tmpInnerText.Replace(item.Key, item.Value);
				}
				if (oldInnerText != tmpInnerText) {
					Run parRun = pars[i].Descendants<Run>().FirstOrDefault();
					pars[i].RemoveAllChildren<Run>();
					if (parRun != null) {
						parRun.RemoveAllChildren<Text>();
						//parRun.Append(new Text(tmpInnerText));
						parRun.Append(new Text { Text = tmpInnerText, Space = SpaceProcessingModeValues.Preserve });
						pars[i].Append(parRun.CloneNode(true));
					}
				}
			}
			outBody.Append(pars.ToArray());
		}
		catch (Exception ex) {
			throw new Exception("複製範本Block錯誤!!(" + blockName + "," + i + ")", ex);
		}
	}
	#endregion

	#region 取代輸出檔文字 +void ReplaceText(string searchStr, string newStr)
	/// <summary>
	/// 取代輸出檔文字
	/// </summary>
	public void ReplaceText(string searchStr, string newStr) {
		List<Paragraph> pars = outBody.Descendants<Paragraph>().ToList();
		for (int i = 0; i < pars.Count; i++) {
			string tmpInnerText = pars[i].InnerText;
			if (tmpInnerText.IndexOf(searchStr) > -1) {
				tmpInnerText = tmpInnerText.Replace(searchStr, newStr);
				Run parRun = pars[i].Descendants<Run>().FirstOrDefault();
				pars[i].RemoveAllChildren<Run>();
				if (parRun != null) {
					parRun.RemoveAllChildren<Text>();
					parRun.Append(new Text(tmpInnerText));
					parRun.Append(new Text { Text = tmpInnerText, Space = SpaceProcessingModeValues.Preserve });
					pars[i].Append(parRun.CloneNode(true));
				}
			}
		}
	}
	#endregion

	#region 取代書籤 +void ReplaceBookmark(string bookmarkName, string text, string ptext)
	/// <summary>
	/// 取代書籤
	/// </summary>
	/// <param name="bookmarkName">書籤名稱</param>
	/// <param name="text">取代的值</param>
	/// <param name="ptext">若取代的值為空,則用此字串代替</param>
	public void ReplaceBookmark(string bookmarkName, string text, string ptext) {
		if (text == "")
			ReplaceBookmark(bookmarkName, ptext, false, System.Drawing.Color.Empty);
		else
			ReplaceBookmark(bookmarkName, text, false, System.Drawing.Color.Empty);
	}
	#endregion

	#region 取代書籤 +void ReplaceBookmark(string bookmarkName, string text, string ptext, System.Drawing.Color color)
	/// <summary>
	/// 取代書籤
	/// </summary>
	/// <param name="bookmarkName">書籤名稱</param>
	/// <param name="text">取代的值</param>
	/// <param name="ptext">若取代的值為空,則用此字串代替</param>
	/// <param name="color">若取代的值為空,套用此顏色</param>
	public void ReplaceBookmark(string bookmarkName, string text, string ptext, System.Drawing.Color color) {
		if (text == "")
			ReplaceBookmark(bookmarkName, ptext, false, color);
		else
			ReplaceBookmark(bookmarkName, text, false, System.Drawing.Color.Empty);
	}
	#endregion

	#region 取代書籤 +void ReplaceBookmark(string bookmarkName, string text)
	/// <summary>
	/// 取代書籤
	/// </summary>
	/// <param name="bookmarkName">書籤名稱</param>
	/// <param name="text">取代的值</param>
	public void ReplaceBookmark(string bookmarkName, string text) {
		ReplaceBookmark(bookmarkName, text, false, System.Drawing.Color.Empty);
	}
	#endregion

	#region 取代書籤 +void ReplaceBookmark(string bookmarkName, string text, System.Drawing.Color color)
	/// <summary>
	/// 取代書籤
	/// </summary>
	/// <param name="bookmarkName">書籤名稱</param>
	/// <param name="text">取代的值</param>
	/// <param name="color">套用此顏色</param>
	public void ReplaceBookmark(string bookmarkName, string text, System.Drawing.Color color) {
		ReplaceBookmark(bookmarkName, text, false, color);
	}
	#endregion

	#region 取代書籤 +void ReplaceBookmark(string bookmarkName, string text, bool delFlag)
	/// <summary>
	/// 取代書籤
	/// </summary>
	/// <param name="bookmarkName">書籤名稱</param>
	/// <param name="text">取代的值</param>
	/// <param name="delFlag">若取代值為空,是否刪除整個段落</param>
	public void ReplaceBookmark(string bookmarkName, string text, bool delFlag) {
		ReplaceBookmark(bookmarkName, text, delFlag, System.Drawing.Color.Empty);
	}
	#endregion

	#region 取代書籤 +void ReplaceBookmark(string bookmarkName, string text, bool delFlag, System.Drawing.Color color)
	/// <summary>
	/// 取代書籤
	/// </summary>
	/// <param name="bookmarkName">書籤名稱</param>
	/// <param name="text">取代的值</param>
	/// <param name="delFlag">若取代值為空,是否刪除整個段落</param>
	/// <param name="color">套用此顏色</param>
	public void ReplaceBookmark(string bookmarkName, string text, bool delFlag, System.Drawing.Color color) {
		try {
			MainDocumentPart mainPart = outDoc.MainDocumentPart;
			//IEnumerable<BookmarkEnd> bookMarkEnds = mainPart.RootElement.Descendants<BookmarkEnd>();
			BookmarkStart bookmarkStart = mainPart.RootElement.Descendants<BookmarkStart>().Where(i => i.Name.Value.ToLower() == bookmarkName.ToLower()).FirstOrDefault();
			if (bookmarkStart != null) {
				//foreach (BookmarkStart bookmarkStart in mainPart.RootElement.Descendants<BookmarkStart>()) {
				if (bookmarkStart.Name.Value.ToLower() == bookmarkName.ToLower()) {
					string id = bookmarkStart.Id.Value;

					//如果是空值,且要刪除整個段落
					if (text == "" && delFlag) {
						bookmarkStart.Parent.Remove();
					} else {
						BookmarkEnd bookmarkEnd = bookmarkStart.Parent.Descendants<BookmarkEnd>().Where(i => i.Id.Value == id).FirstOrDefault();

						if (bookmarkStart.NextSibling() != null && bookmarkStart.NextSibling().GetType() == typeof(BookmarkEnd) && ((BookmarkEnd)bookmarkStart.NextSibling()).Id == id) {
							//Console.WriteLine(bookmarkName + "=11111");
							Run LastRun = bookmarkStart.PreviousSibling<Run>();
							if (LastRun == null) {
								LastRun = bookmarkStart.Ancestors<Paragraph>().FirstOrDefault().AppendChild(new Run());
								//LastRun.AppendChild(new RunProperties(new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" }));
								RunFonts f = bookmarkStart.Ancestors<Paragraph>().FirstOrDefault().Descendants<RunFonts>().FirstOrDefault();
								LastRun.AppendChild(new RunProperties(f.CloneNode(true)));
							}
							string[] txtArr = text.Split('\n');
							for (int i = 0; i < txtArr.Length; i++) {
								if (i != 0) {
									LastRun.Append(new Break());
								}
								//LastRun.Append(new Text(txtArr[i]));
								LastRun.Append(new Text { Text = txtArr[i], Space = SpaceProcessingModeValues.Preserve });
							}
							//bookmarkStart.InsertBeforeSelf(LastRun);
						} else {
							//Console.WriteLine(bookmarkName + "=22222");
							//留第一個run其他run刪除,從BookmarkStart刪到BookmarkEnd為止
							OpenXmlElement[] bookmarkItems = bookmarkStart.Parent.ChildElements.ToArray();
							//HttpContext.Current.Response.Write(bookmarkItems.Count());
							//HttpContext.Current.Response.End();

							bool canRemove = false;
							int bIndex = 0;

							foreach (OpenXmlElement item in bookmarkItems) {
								if (item.GetType() == typeof(BookmarkEnd) && bookmarkEnd != null && bookmarkEnd.Id == id) {
									break;
								}
								if (canRemove && item.GetType() == typeof(Run)) {
									if (bIndex == 0) {
										string[] txtArr = text.Split('\n');
										for (int i = 0; i < txtArr.Length; i++) {
											if (i == 0) {
												if (color != System.Drawing.Color.Empty) {
													RunProperties FirstRunProp = item.Descendants<RunProperties>().FirstOrDefault();
													if (FirstRunProp == null) {
														FirstRunProp = new RunProperties();
													}
													Color RunColor = new Color() { Val = toHtmlHexColor(color) };
													FirstRunProp.Append(RunColor);
												}
												item.GetFirstChild<Text>().Text = txtArr[i];
											} else {
												item.Append(new Break());
												//item.Append(new Text(txtArr[i]));
												item.Append(new Text { Text = txtArr[i], Space = SpaceProcessingModeValues.Preserve });
											}
										}
									} else {
										item.Remove();
									}
									bIndex++;
								}
								//if (item.GetType() == typeof(BookmarkStart)) {
								if (item.Equals(bookmarkStart)) {
									canRemove = true;
								}
							}
						}

						bookmarkStart.Remove();
						if (bookmarkEnd != null) bookmarkEnd.Remove();
					}
				}
			}
		}
		catch (Exception ex) {
			throw new Exception("取代書籤錯誤!!(" + bookmarkName + ")", ex);
		}
	}
	#endregion

	#region 複製範本頁尾 +void CopyPageFoot(string srcDocName, bool isNewChapter)
	/// <summary>
	/// 複製範本頁尾
	/// </summary>
	/// <param name="srcDocName">來源範本別名</param>
	public void CopyPageFoot(string srcDocName) {
		CopyPageFoot(srcDocName, false);
	}
	#endregion

	#region 複製範本頁尾 +void CopyPageFoot(string srcDocName, bool isNewChapter)
	/// <summary>
	/// 複製範本頁尾
	/// </summary>
	/// <param name="srcDocName">來源範本別名</param>
	/// <param name="isNewChapter">是否帶分節符號(新章節)</param>
	public void CopyPageFoot(string srcDocName, bool isNewChapter) {
		WordprocessingDocument sourceDoc = tplDoc[srcDocName];
		int index = 0;//取消index參數,只抓第1個
		FooterReference[] footer = sourceDoc.MainDocumentPart.RootElement.Descendants<FooterReference>().Where(x => x.Type == HeaderFooterValues.Default).ToArray();
		string newRefId = string.Format("foot_{0}", Guid.NewGuid().ToString().Substring(0, 8));
		if (srcDocName != defTplDocName) {
			string srcRefId = footer[index].Id;
			footer[index].Id = newRefId;
			FooterPart elementFoot = sourceDoc.MainDocumentPart.FooterParts
			.Where(
				element => sourceDoc.MainDocumentPart.GetIdOfPart(element) == srcRefId
			).SingleOrDefault();
			outDoc.MainDocumentPart.AddPart(elementFoot, newRefId);

			/*
			FooterPart fp = outDoc.MainDocumentPart.FooterParts
			.Where(
				element => outDoc.MainDocumentPart.GetIdOfPart(element) == newRefId
			).SingleOrDefault();
			ParagraphStyleId pStyle = fp.Footer.Descendants<ParagraphStyleId>().First();
			string oldStyleId = pStyle.Val;
			string newStyleId = string.Format("fs_{0}", Guid.NewGuid().ToString().Substring(0, 8));
			pStyle.Val = newStyleId;

			StylesPart srcPart = sourceDoc.MainDocumentPart.StyleDefinitionsPart;
			Style st = srcPart.Styles.Descendants<Style>()
			.Where(
				element => element.StyleId == oldStyleId
			).SingleOrDefault();
			st.StyleId = newStyleId;

			StylesPart outPart = outDoc.MainDocumentPart.StyleDefinitionsPart;
			outPart.Styles.Append(st.CloneNode(true));
             * */
		}
		/*
        if (isNewChapter) {
            //outBody.AppendChild(new Paragraph(new ParagraphProperties(footer[index].Parent.CloneNode(true))));//頁尾+分節符號
            outBody.Append(new Paragraph(new ParagraphProperties(footer[index].Parent.CloneNode(true))));//頁尾+分節符號
        } else {
            SectionProperties sp = (SectionProperties)footer[index].Parent.CloneNode(true);
            OpenXmlElement[] spItemsNode = sp.ChildElements.ToArray();
            foreach (OpenXmlElement item in spItemsNode) {
                if (item.GetType() == typeof(HeaderReference)) {
                    item.Remove();
                } else if (item.GetType() == typeof(FooterReference)) {
                    if (((FooterReference)item).Id != newRefId)
                        item.Remove();
                }
            }
            outBody.AppendChild(sp);//頁尾
        }
        */
		if (footer.Length > 0) {
			SectionProperties sp = (SectionProperties)footer[index].Parent.CloneNode(true);
			OpenXmlElement[] spItemsNode = sp.ChildElements.ToArray();
			foreach (OpenXmlElement item in spItemsNode) {
				if (item.GetType() == typeof(HeaderReference)) {
					//item.Remove();
				} else if (item.GetType() == typeof(FooterReference)) {
					if (((FooterReference)item).Id != newRefId && srcDocName != defTplDocName)
						item.Remove();
				}
			}
			if (isNewChapter) {
				//outBody.AppendChild(new Paragraph(new ParagraphProperties(footer[index].Parent.CloneNode(true))));//頁尾+分節符號
				//outBody.Append(new Paragraph(new ParagraphProperties(footer[index].Parent.CloneNode(true))));//頁尾+分節符號
				outBody.Append(new Paragraph(new ParagraphProperties(sp)));//頁尾+分節符號
			} else {
				outBody.AppendChild(sp);//頁尾
			}
		} else {
			if (isNewChapter) {
				outBody.Append(new Paragraph(new ParagraphProperties(sourceDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().FirstOrDefault().CloneNode(true))));//沒有頁尾則copy邊界設定
			} else {
				outBody.AppendChild(sourceDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().FirstOrDefault().CloneNode(true));//沒有頁尾則copy邊界設定
			}
		}

		/*
        if (isNewChapter) {
            if (footer.Length > 0) {
                //outBody.AppendChild(new Paragraph(new ParagraphProperties(footer[index].Parent.CloneNode(true))));//頁尾+分節符號
                //outBody.Append(new Paragraph(new ParagraphProperties(footer[index].Parent.CloneNode(true))));//頁尾+分節符號
                outBody.Append(new Paragraph(new ParagraphProperties(sp)));//頁尾+分節符號
            } else {
                outBody.Append(new Paragraph(new ParagraphProperties(sourceDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().FirstOrDefault().CloneNode(true))));//沒有頁尾則copy邊界設定
            }
        } else {
            if (footer.Length > 0) {
                outBody.AppendChild(sp);//頁尾
            } else {
                outBody.AppendChild(sourceDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().FirstOrDefault().CloneNode(true));//沒有頁尾則copy邊界設定
            }
        }*/

	}
	#endregion

	#region 複製範本頁首 +void CopyPageHeader(string srcDocName)
	/// <summary>
	/// 複製範本頁首
	/// </summary>
	/// <param name="srcDocName">來源範本別名</param>
	public void CopyPageHeader(string srcDocName) {
		WordprocessingDocument sourceDoc = tplDoc[srcDocName];
		int index = 0;//取消index參數,只抓第1個
		HeaderReference[] header = sourceDoc.MainDocumentPart.RootElement.Descendants<HeaderReference>().Where(x => x.Type == HeaderFooterValues.Default).ToArray();
		string newRefId = string.Format("head_{0}", Guid.NewGuid().ToString().Substring(0, 8));
		if (srcDocName != defTplDocName) {
			string srcRefId = header[index].Id;
			header[index].Id = newRefId;
			HeaderPart elementHeader = sourceDoc.MainDocumentPart.HeaderParts
			.Where(
				element => sourceDoc.MainDocumentPart.GetIdOfPart(element) == srcRefId
			).SingleOrDefault();
			outDoc.MainDocumentPart.AddPart(elementHeader, newRefId);
		}

		SectionProperties sp = (SectionProperties)header[index].Parent.CloneNode(true);
		OpenXmlElement[] spItemsNode = sp.ChildElements.ToArray();
		foreach (OpenXmlElement item in spItemsNode) {
			if (item.GetType() == typeof(FooterReference)) {
				//item.Remove();
			} else if (item.GetType() == typeof(HeaderReference)) {
				if (((HeaderReference)item).Id != newRefId && srcDocName != defTplDocName)
					item.Remove();
			}
		}

		if (header.Length > 0) {
			outBody.AppendChild(sp);//頁尾
		} else {
			outBody.AppendChild(sourceDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().FirstOrDefault().CloneNode(true));//沒有頁尾則copy邊界設定
		}

	}
	#endregion

	#region 增加段落 +OpenXmlHelper AddParagraph(Paragraph par)
	/// <summary>
	/// 增加段落
	/// </summary>
	public OpenXmlHelper AddParagraph(Paragraph par) {
		//outDoc.MainDocumentPart.Document.Body.Append(par.CloneNode(true));
		outBody.Append(par.CloneNode(true));
		return this;
	}
	#endregion

	#region 增加段落 +OpenXmlHelper AddParagraph()
	/// <summary>
	/// 增加段落
	/// </summary>
	public OpenXmlHelper AddParagraph() {
		Paragraph NewPar = new Paragraph();
		ParagraphProperties LastParProp = outDoc.MainDocumentPart.RootElement.Descendants<ParagraphProperties>().LastOrDefault();
		if (LastParProp != null) {
			NewPar.Append(LastParProp.CloneNode(true));
		}

		Run LastRun = new Run();
		RunProperties LastRunProp = outDoc.MainDocumentPart.RootElement.Descendants<RunProperties>().LastOrDefault();
		if (LastRunProp != null) {
			LastRun.Append(LastRunProp.CloneNode(true));
		}
		NewPar.Append(LastRun);
		outBody.Append(NewPar);
		//outBody.Append(new Paragraph(new Run()));
		return this;
	}
	#endregion

	#region 在文件最後的段落加上文字 +OpenXmlHelper AddText(string text)
	/// <summary>
	/// 在文件最後的段落加上文字
	/// </summary>
	public OpenXmlHelper AddText(string text) {
		return AddText(text, System.Drawing.Color.Empty);
	}
	#endregion

	#region 在文件最後的段落加上文字 +OpenXmlHelper AddText(string text, System.Drawing.Color color)
	/// <summary>
	/// 在文件最後的段落加上文字
	/// </summary>
	public OpenXmlHelper AddText(string text, System.Drawing.Color color) {
		Paragraph LastPar = outDoc.MainDocumentPart.RootElement.Descendants<Paragraph>().LastOrDefault();
		if (LastPar == null) {
			LastPar = new Paragraph();
		}

		RunProperties LastRunProp = (RunProperties)outDoc.MainDocumentPart.RootElement.Descendants<RunProperties>().LastOrDefault().CloneNode(true);
		if (LastRunProp == null) {
			LastRunProp = new RunProperties();
		}
		Run LastRun = new Run();
		if (color != System.Drawing.Color.Empty) {
			Color RunColor = new Color() { Val = toHtmlHexColor(color) };
			LastRunProp.Append(RunColor);
			LastRun.Append(LastRunProp.CloneNode(true));
		}
		string[] txtArr = text.Split('\n');
		for (int i = 0; i < txtArr.Length; i++) {
			if (i != 0) {
				LastRun.Append(new Break());
			}
			//LastRun.Append(new Text(txtArr[i]));
			LastRun.Append(new Text { Text = txtArr[i], Space = SpaceProcessingModeValues.Preserve });
		}
		LastPar.Append(LastRun);

		return this;
	}
	#endregion

	#region 插入換行符號(Shift-Enter) +OpenXmlHelper NewLine()
	/// <summary>
	/// 插入換行符號(Shift-Enter)
	/// </summary>
	public OpenXmlHelper NewLine() {
		Run LastRun = outDoc.MainDocumentPart.RootElement.Descendants<Run>().LastOrDefault();
		if (LastRun != null) {
			LastRun.Append(new Break());
		}
		return this;
	}
	#endregion

	#region 插入分頁符號(Ctrl-Enter) +OpenXmlHelper NewPage()
	/// <summary>
	/// 插入分頁符號(Ctrl-Enter)
	/// </summary>
	public OpenXmlHelper NewPage() {
		outBody.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));//分頁符號
		return this;
	}
	#endregion

	#region 設定紙張大小 +OpenXmlHelper SetPageSize(double widthCM, double heightCM)
	/// <summary>
	/// 設定紙張大小
	/// </summary>
	/// <param name="widthCM">寬(公分)</param>
	/// <param name="heightCM">高(公分)</param>
	public OpenXmlHelper SetPageSize(double widthCM, double heightCM) {
		//SectionProperties sections0 = outDoc.MainDocumentPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();
		//SectionProperties sections0 = outDoc.MainDocumentPart.RootElement.Descendants<SectionProperties>().FirstOrDefault();
		//SectionProperties sectPr = outDoc.MainDocumentPart.Document.Descendants<SectionProperties>().FirstOrDefault();
		if (outDoc.MainDocumentPart.Document.Descendants<SectionProperties>().FirstOrDefault() == null) {
			outBody.Append(new SectionProperties());
		}

		var sections = outDoc.MainDocumentPart.Document.Descendants<SectionProperties>();
		foreach (SectionProperties sectPr in sections) {
			//PageSize pageSize = sections0.GetFirstChild<PageSize>();
			PageSize pgSz = sectPr.Descendants<PageSize>().FirstOrDefault();
			if (pgSz == null) {
				//pageSize = new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U };
				//pageSize = new PageSize() { Width = 11906, Height = 16838, Orient = PageOrientationValues.Portrait };//21、29.7//直向
				//pageSize = new PageSize() { Width = (UInt32Value)16838U, Height = (UInt32Value)11906U, Orient = PageOrientationValues.Landscape };
				//pageSize = new PageSize() { Width = 16838, Height = 11906, Orient = PageOrientationValues.Landscape };//橫向
				pgSz = new PageSize();
				sectPr.Append(pgSz);
			}
			pgSz.Height = Convert.ToUInt32(Math.Round((decimal)heightCM * (decimal)566.9523, 0));
			pgSz.Width = Convert.ToUInt32(Math.Round((decimal)widthCM * (decimal)566.9523, 0));
		}

		return this;
	}
	#endregion

	#region 設為直向 +OpenXmlHelper SetPagePortrait()
	/// <summary>
	/// 設為直向
	/// </summary>
	public OpenXmlHelper SetPagePortrait() {
		return SetPageOrientation(PageOrientationValues.Portrait);
	}
	#endregion

	#region 設為橫向 +OpenXmlHelper SetPageLandscape()
	/// <summary>
	/// 設為橫向
	/// </summary>
	public OpenXmlHelper SetPageLandscape() {
		return SetPageOrientation(PageOrientationValues.Landscape);
	}
	#endregion

	#region 設定紙張方向 #OpenXmlHelper SetPageOrientation(PageOrientationValues newOrientation)
	/// <summary>
	/// 設定紙張方向,需先設定紙張大小,否則無作用
	/// </summary>
	/// <param name="newOrientation">Landscape:橫向;Portrait:直向</param>
	protected OpenXmlHelper SetPageOrientation(PageOrientationValues newOrientation) {
		var sections = outDoc.MainDocumentPart.Document.Descendants<SectionProperties>();
		foreach (SectionProperties sectPr in sections) {
			bool pageOrientationChanged = false;

			PageSize pgSz = sectPr.Descendants<PageSize>().FirstOrDefault();
			if (pgSz != null) {
				if (pgSz.Orient == null) {
					if (newOrientation != PageOrientationValues.Portrait) {
						pgSz.Orient = new EnumValue<PageOrientationValues>(newOrientation);
						pageOrientationChanged = true;
					}
				} else {
					if (pgSz.Orient.Value != newOrientation) {
						pgSz.Orient.Value = newOrientation;
						pageOrientationChanged = true;
					}
				}
				if (pageOrientationChanged) {
					var width = pgSz.Width;
					var height = pgSz.Height;
					pgSz.Width = height;
					pgSz.Height = width;

					PageMargin pgMar = sectPr.Descendants<PageMargin>().FirstOrDefault();
					if (pgMar != null) {
						var top = pgMar.Top.Value;
						var bottom = pgMar.Bottom.Value;
						var left = pgMar.Left.Value;
						var right = pgMar.Right.Value;

						pgMar.Top = new Int32Value((int)left);
						pgMar.Bottom = new Int32Value((int)right);
						pgMar.Left = new UInt32Value((uint)Math.Max(0, bottom));
						pgMar.Right = new UInt32Value((uint)Math.Max(0, top));
					}
				}
			}
		}
		return this;
	}
	#endregion

	#region 複製表格 +void CopyTable(int index)
	public void CopyTable(int index) {
		CopyTable(defTplDocName, index);
	}
	#endregion

	#region 複製表格 +void CopyTable(string srcDocName, int index)
	public void CopyTable(string srcDocName, int index) {
		try {
			WordprocessingDocument srcDoc = tplDoc[srcDocName];
			List<Table> tblElement = srcDoc.MainDocumentPart.RootElement.Descendants<Table>().ToList();

			if (tblElement != null) {
				//tmpTable = tblElement[index];
				outBody.AppendChild(tblElement[index].CloneNode(true));
			}
		}
		catch (Exception ex) {
			throw new Exception("複製Table錯誤!!(" + index + ")", ex);
		}
	}
	#endregion

	#region 轉換顏色 +static string toHtmlHexColor(System.Drawing.Color color)
	/// <summary>
	/// Color轉換成html色碼
	/// </summary>
	public static string toHtmlHexColor(System.Drawing.Color color) {
		return String.Format("#{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B);
	}
	#endregion

	#region 刪除圖片 +void RemoveImage(string imgText)
	/// <summary>
	/// 刪除圖片
	/// </summary>
	/// <param name="imgText">圖片的替代文字</param>
	public void RemoveImage(string imgText) {
		MainDocumentPart mainPart = outDoc.MainDocumentPart;

		DW.DocProperties picProp = mainPart.RootElement.Descendants<DW.DocProperties>().Where(i => i.Description == imgText).FirstOrDefault();
		if (picProp != null) {
			if (picProp.Parent.Parent.GetType() == typeof(Drawing)) {
				picProp.Parent.Parent.Remove();
			}
		}
	}
	#endregion

	#region 插入圖片 +void AppendImage(ImageFile img)
	/// <summary>
	/// 插入圖片
	/// </summary>
	//public void AppendImage(string imgStr, bool isBase64, decimal scale) {
	//	ImageData img= new ImageData(imgStr, isBase64, scale);
	public void AppendImage(ImageFile img) {
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
		outDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph());
	}
	#endregion

}

#region Docx
public class Docx
{
	public string FileName { get; set; }
	public string Alias { get; set; }
	public bool BeforeBreak { get; set; }

	public Docx() { }
	public Docx(string fileName, string alias, bool beforeBreak) {
		FileName = fileName;
		Alias = alias;
		BeforeBreak = beforeBreak;
	}

	public Docx(string fileName, bool beforeBreak) : this(fileName, "", beforeBreak) { }
}
#endregion

#region ImageFile
public class ImageFile {
	public string FileName = string.Empty;

	public byte[] BinaryData;

	public Stream getDataStream() {
		//Stream DataStream = new MemoryStream(BinaryData);
		return new MemoryStream(BinaryData);
	}

	public ImagePartType ImageType {
		get {
			var ext = Path.GetExtension(FileName).TrimStart('.').ToLower();
			switch (ext) {
				case "jpg":
					return ImagePartType.Jpeg;
				case "png":
					return ImagePartType.Png;
				case "bmp":
					return ImagePartType.Bmp;
			}
			throw new ApplicationException(string.Format("不支援的格式:{0}", ext));
		}
	}

	//public int SourceWidth;
	//public int SourceHeight;
	public decimal Width;
	public decimal Height;

	//public long WidthInEMU => Convert.ToInt64(Width * CM_TO_EMU);
	//private long WidthInEMU = 0;
	public long GetWidthInEMU() {
		//WidthInEMU = Convert.ToInt64(Width * CM_TO_EMU);
		return Convert.ToInt64(Width * CM_TO_EMU);
	}

	//public long HeightInEMU => Convert.ToInt64(Height * CM_TO_EMU);
	//private long HeightInEMU = 0;
	public long GetHeightInEMU() {
		//HeightInEMU = Convert.ToInt64(Height * CM_TO_EMU);
		return Convert.ToInt64(Height * CM_TO_EMU);
	}

	private const decimal INCH_TO_CM = 2.54M;
	private const decimal CM_TO_EMU = 360000M;
	public string ImageName;

	public ImageFile(string fileName, byte[] data, decimal scale) {
		if (fileName == "") {
			FileName = string.Format("IMG_{0}", Guid.NewGuid().ToString().Substring(0, 8));
			ImageName = FileName;
		} else {
			FileName = fileName;
			ImageName = string.Format("IMG_{0}", Guid.NewGuid().ToString().Substring(0, 8));
		}

		BinaryData = data;
		System.Drawing.Bitmap img = new System.Drawing.Bitmap(new MemoryStream(data));
		int dpi = 300;
		//int dpi = (int)(img.VerticalResolution < 300 ? 300 : img.VerticalResolution);
		int SourceWidth = img.Width;
		int SourceHeight = img.Height;
		//throw new Exception(string.Format("w:{0},h:{1},dpi:{2}", SourceWidth, SourceHeight, dpi));
		Width = ((decimal)SourceWidth) / dpi * scale * INCH_TO_CM;
		Height = ((decimal)SourceHeight) / dpi * scale * INCH_TO_CM;
	}

	public ImageFile(byte[] data) :
		this("", data, 1) {
	}

	public ImageFile(byte[] data, decimal scale) :
		this("", data, scale) {
	}

	public ImageFile(string fileName) :
		this(fileName, File.ReadAllBytes(fileName), 1) {
	}

	public ImageFile(string fileName, decimal scale) :
		this(fileName, File.ReadAllBytes(fileName), scale) {
	}
}
#endregion
