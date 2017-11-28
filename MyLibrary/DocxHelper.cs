using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
//https://raikkonenlee0528.blogspot.tw/2016/04/open-xml-sdk-word.html
//http://www.bkjia.com/Asp_Netjc/1083236.html
//http://old-kai.blogspot.tw/2009/10/word-2.html
//http://blog.darkthread.net/post-2013-05-31-word-to-pdf.aspx


namespace MyLibrary {
	/// <summary>
	/// ref:https://coolong124220.nidbox.com/diary/read/8045524
	/// Ver 1.0 By Jeffrey Lee, 2009-07-29
	/// </summary>
	public class DocxHelper {
		/// <summary>
		/// Replace the parser tags in docx document
		/// </summary>
		/// <param name="oxp">OpenXmlPart object</param>
		/// <param name="dct">Dictionary contains parser tags to replace</param>
		private static void parse(OpenXmlPart oxp, Dictionary<string, string> dct) {
			string xmlString = null;
			using (StreamReader sr = new StreamReader(oxp.GetStream())) { xmlString = sr.ReadToEnd(); }
			foreach (string key in dct.Keys)
				xmlString = xmlString.Replace("[$" + key + "$]", dct[key]);
			using (StreamWriter sw = new StreamWriter(oxp.GetStream(FileMode.Create))) { sw.Write(xmlString); }
		}

		/// <summary>
		/// Parse template file and replace all parser tags, return the binary content of
		/// new docx file.
		/// </summary>
		/// <param name="templateFile">template file path</param>
		/// <param name="dct">a Dictionary containing parser tags and values</param>
		/// <returns></returns>
		public static byte[] MakeDocx(string templateFile, Dictionary<string, string> dct) {
			string tempFile = Path.GetTempPath() + ".docx";
			File.Copy(templateFile, tempFile);

			using (WordprocessingDocument wd = WordprocessingDocument.Open(tempFile, true)) {
				//Replace document body
				parse(wd.MainDocumentPart, dct);
				foreach (HeaderPart hp in wd.MainDocumentPart.HeaderParts)
					parse(hp, dct);
				foreach (FooterPart fp in wd.MainDocumentPart.FooterParts)
					parse(fp, dct);
			}
			byte[] buff = File.ReadAllBytes(tempFile);
			File.Delete(tempFile);
			return buff;
		}

		/// <summary>
		/// 複製範本並取出內容的部分
		/// </summary>
		/// 
		/// <param name="fromWordFile">template file path</param>
		/// <param name="toWordFile">output file path</param>
		/// 
		private void CopyWordFile(string fromWordFile, string toWordFile, out string templateText) {
			if (File.Exists(toWordFile)) {
				File.Delete(toWordFile);
			}
			//
			File.Copy(fromWordFile, toWordFile);
			//
			using (WordprocessingDocument fromDoc = WordprocessingDocument.Open(toWordFile, true)) {
				templateText = fromDoc.MainDocumentPart.Document.Body.InnerXml;
				fromDoc.MainDocumentPart.Document.RemoveAllChildren();
			}
		}

		/// 
		/// 設定WORD檔需要的參數
		/// 
		/// 
		/// 
		/// 
		private void SetWordDictionary(string templateText, out string docText) {
			Dictionary<string, string> keyWordDict = new Dictionary<string, string>();
			keyWordDict.Add("SEQTXT", "1");
			keyWordDict.Add("YEARXX", "105");
			keyWordDict.Add("DEPTNM", "部門");
			keyWordDict.Add("PLANNM", "計畫");
			keyWordDict.Add("DATEXX", "日期");
			keyWordDict.Add("BGTWTT", "金額");
			//
			ReplaceTemplateString(keyWordDict, templateText, out docText);
		}

		/// 
		/// 將WORD檔所設定的參數取代掉
		/// 
		/// 
		/// 
		/// 
		/// 
		private void ReplaceTemplateString(Dictionary<string, string> keyWordDict, string templateText, out string docText) {
			foreach (KeyValuePair<string, string> item in keyWordDict) {
				Regex regex = new Regex(item.Key);
				templateText = regex.Replace(templateText, item.Value);
			}
			//
			docText = templateText;
		}

		/// 
		/// 將XML的文字寫入到WORD檔
		/// 
		/// 
		/// 
		/// 
		private void SetWordFile(string toFileName, string docText) {
			using (WordprocessingDocument toDoc = WordprocessingDocument.Open(toFileName, true)) {
				MainDocumentPart mainPart = toDoc.MainDocumentPart;
				Body insertBody = mainPart.Document.AppendChild(new Body());
				insertBody.InnerXml = docText;
			}
		}
	}
}