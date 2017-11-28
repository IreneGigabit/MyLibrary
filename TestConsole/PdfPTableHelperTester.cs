using System;
using System.Collections.Generic;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Diagnostics;
using MyLibrary;
using System.Data;
using System.Data.SqlClient;

namespace TestConsole
{
    class PdfPTableHelperTester
    {
        private static string CurrDir = System.Environment.CurrentDirectory;
        static string templateFile = CurrDir + @"\Credit_letter_apply_one.pdf";
        //static string templateFile = CurrDir + @"\pdf29.pdf";
        static string outputFile = CurrDir + @"\new.pdf";


        //static string fontPath = Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\..\Fonts\kaiu.ttf";//標楷體
        //static string fontPath = Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\..\Fonts\mingliu.ttc,0";//細明體
        static string fontPath = Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\..\Fonts\msjh.ttc,0";//正黑體
        //static string fontPath = Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\..\Fonts\msjh.ttf";//正黑體

        static void Main(string[] args)
        {
            Console.WriteLine("Template File:" + templateFile);
            Console.WriteLine("Output File:" + outputFile);

			NorthWind();
            //multiplePage();
            //tablePage();
            //memorystream();
            //filestream();

            Process.Start(outputFile);

            Console.WriteLine("請按任一鍵離開..");
            Console.ReadKey();
        }

        #region multiplePage - 套印表單欄位範本
        public static void multiplePage()
        {
            List<Byte[]> result = new List<byte[]>();

            for (int i = 1; i <= 2; i++)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    PdfReader pdfReader = new PdfReader(templateFile);
                    PdfStamper pdfStamper = new PdfStamper(pdfReader, ms);

                    AcroFields pdfFormFields = pdfStamper.AcroFields;
                    List<String> keys = new List<String>(pdfFormFields.Fields.Keys);
                    foreach (var key in keys)
                    {
                        // rename the fields
                        pdfFormFields.RenameField(key, String.Format("{0}_{1}", key, i));
                    }
                    pdfFormFields.AddSubstitutionFont(BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED));

                    pdfFormFields.SetField("Radio Button1_" + i, "1");
                    pdfFormFields.SetField("Check Box101_" + i, "是");
                    pdfFormFields.SetField("Currency_" + i, "是");
                    //String[] values = pdfFormFields.GetAppearanceStates("Check Box101");
                    //foreach (String value in values) {
                    //	Console.WriteLine("values=" + value);
                    //}
                    pdfFormFields.SetField("Text21_" + i, "(" + i + ")聖島國際專利商標聯合事務所");
                    pdfFormFields.SetField("Text74_" + i, "台北市松山區南京東路三段248號11樓之1");

                    pdfStamper.Writer.CloseStream = false;
                    //pdfStamper.FormFlattening = true;//平面化

                    pdfStamper.Close();
                    pdfReader.Close();

                    result.Add(ms.ToArray());
                }
            }

            MemoryStream outStream = new MemoryStream();
            outStream = MergePdfForms(result);

            System.IO.File.WriteAllBytes(outputFile, outStream.ToArray());
        }
        #endregion

        #region tablePage - 輸出表格
        public static void tablePage()
        {
            Document doc = new Document(PageSize.A4.Rotate(), 15, 5, 10, 20); //A4橫式,Marginleft,Marginright,Margintop,Marginbottom
            MemoryStream Memory = new MemoryStream();
            PdfWriter PdfWriter = PdfWriter.GetInstance(doc, Memory);

            //字型設定(不使用BaseFont防止PDF內嵌字型導致檔案過大)
            BaseFont bfChinese = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            Font FontMid = new Font(bfChinese, 12);
            Font FontSml = new Font(bfChinese, 10);
            doc.Open();

            //表格
            //PdfPTable table = new PdfPTable(new float[] { 65f, 73f, 124f, 50f, 70f, 105f, 63f, 50f, 40f, 40f, 80f });
            PdfPTable table = new PdfPTable(11);
            PdfPTableHelper tbl = new PdfPTableHelper(table, FontSml);

            table.TotalWidth = 780f;
            table.LockedWidth = true;
            table.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
            table.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            table.DefaultCell.MinimumHeight = 30f;
            table.AddCell(new Paragraph("發貨日", FontSml));
            table.AddCell(new Paragraph("發票號碼", FontSml));
            table.AddCell(new Paragraph("出貨品項", FontSml));
            table.AddCell(new Paragraph("數量", FontSml));
            table.AddCell(new Paragraph("總金額", FontSml));
            table.AddCell(new Paragraph("批號", FontSml));
            table.AddCell(new Paragraph("效期", FontSml));
            table.AddCell(new Paragraph("退(換)貨\n原因", FontSml));
            table.AddCell(new Paragraph("同意", FontSml));
            table.AddCell(new Paragraph("不同意", FontSml));
            table.AddCell(new Paragraph("原廠確認簽名", FontSml));

            tbl.Cell("C00)發票作廢、退貨").Colspan(2).Add();
            tbl.Cell("C05)滯銷").Colspan(3).Add();
            tbl.Cell("C10)結束營業").Colspan(4).Add();
            tbl.Cell("").Colspan(2).Add();

            tbl.Cell("C01)業代誤訂").Colspan(2).Add();
            tbl.Cell("C06)壞損-於配送途中").Colspan(3).Add();
            tbl.Cell("C99)其他").Colspan(4).Add();
            tbl.Cell("").Colspan(2).Add();

            PdfPTable table1 = new PdfPTable(new float[] { 40f, 60f });
            table1.AddCell("C02)客戶誤訂");
            table1.AddCell("C07)壞損-客戶收貨後");
            table1.AddCell("C03)久裕誤寄");
            table1.AddCell("C08)過期");

            tbl.AddTable(table1, 5);
            tbl.Cell("RTN_REMARK").Colspan(6).Add();

            tbl.Cell("C04)配送公司誤送").Colspan(2).Add();
            tbl.Cell("C09)短效期").Colspan(3).Add();

            tbl.Cell("列印日期：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")).Colspan(11).Add();

            table.WriteSelectedRows(0, -1, 30, 425, PdfWriter.DirectContent);

            doc.Close();

            using (FileStream fs = File.Create(outputFile))
            {
                fs.Write(Memory.GetBuffer(), 0, Memory.GetBuffer().Length);
            }

        }
        #endregion

        #region MergePdfForms - 合併pdf(multiplePage使用)
        public static MemoryStream MergePdfForms(List<byte[]> files)
        {
            if (files.Count > 1)
            {
                PdfReader pdfFile;
                Document doc;
                PdfWriter pCopy;
                MemoryStream msOutput = new MemoryStream();
                pdfFile = new PdfReader(files[0]);
                doc = new Document();
                pCopy = new PdfCopy(doc, msOutput);
                doc.Open();
                for (int k = 0; k < files.Count; k++)
                {
                    pdfFile = new PdfReader(files[k]);
                    for (int i = 1; i < pdfFile.NumberOfPages + 1; i++)
                    {
                        ((PdfCopy)pCopy).AddPage(pCopy.GetImportedPage(pdfFile, i));
                    }
                    pCopy.FreeReader(pdfFile);
                }
                pdfFile.Close();
                pCopy.Close();
                doc.Close();
                return msOutput;
            }
            else if (files.Count == 1)
            {
                return new MemoryStream(files[0]);
            }
            return null;
        }
        #endregion

        #region memorystream - 使用MemoryStream產生pdf
        public static void memorystream()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                PdfReader pdfReader = new PdfReader(templateFile, null);
                PdfStamper pdfStamper = new PdfStamper(pdfReader, ms);

                AcroFields pdfFormFields = pdfStamper.AcroFields;
                pdfFormFields.AddSubstitutionFont(BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED));

                pdfFormFields.SetField("Radio Button1", "1");
                pdfFormFields.SetField("Check Box101", "是");
                pdfFormFields.SetField("Currency", "是");
                //String[] values = pdfFormFields.GetAppearanceStates("Check Box101");
                //foreach (String value in values) {
                //	Console.WriteLine("values=" + value);
                //}
                pdfFormFields.SetField("Text21", "()聖島國際專利商標聯合事務所");
                pdfFormFields.SetField("Text74", "台北市松山區南京東路三段248號11樓之1");

                //pdfStamper.Writer.CloseStream = false;
                //pdfStamper.FormFlattening = true;//平面化
                pdfStamper.Close();
                pdfReader.Close();

                System.IO.File.WriteAllBytes(outputFile, ms.ToArray());
            }
        }
        #endregion

        #region 使用FileStream產生pdf
        public static void filestream()
        {
            PdfReader pdfReader = new PdfReader(templateFile, null);
            PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(outputFile, FileMode.Create));

            AcroFields pdfFormFields = pdfStamper.AcroFields;
            pdfFormFields.AddSubstitutionFont(BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED));

            pdfFormFields.SetField("Radio Button1", "1");
            pdfFormFields.SetField("Check Box101", "是");
            pdfFormFields.SetField("Currency", "是");
            //String[] values = pdfFormFields.GetAppearanceStates("Check Box101");
            //foreach (String value in values) {
            //	Console.WriteLine("values=" + value);
            //}
            pdfFormFields.SetField("Text21", "聖島國際專利商標聯合事務所");
            pdfFormFields.SetField("Text74", "台北市松山區南京東路三段248號11樓之1");
            //pdfStamper.Writer.CloseStream = false;
            //pdfStamper.FormFlattening = true;//平面化
            pdfStamper.Close();
            pdfReader.Close();
        }
        #endregion

		#region 使用北風資料庫產生master/detail換頁報表
		public static void NorthWind() {
			//字型設定(不使用BaseFont防止PDF內嵌字型導致檔案過大)
			BaseFont bfChinese = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
			Font FontMid = new Font(bfChinese, 12);
			Font FontSml = new Font(bfChinese, 10);

			//定義紙張
			Document doc = new Document(PageSize.A4, 15, 5, 30, 30); //A4直式,MarginLeft,MarginRight,MarginTop,MarginBottom
			MemoryStream MS = new MemoryStream();
			PdfWriter PdfWriter = PdfWriter.GetInstance(doc, MS);
			doc.Open();

			//定義表格
			PdfPCell cell = new PdfPCell();
			//定義空白行
			Paragraph br = new Paragraph(" ", FontSml);
			br.Alignment = PdfPCell.ALIGN_CENTER;

			DataTable dt = new DataTable();
			using (SqlConnection cnQry = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=""D:\SQL Server 2005 Sample Databases\NORTHWND.MDF"";Integrated Security=True;")) {
				cnQry.Open();
				string SQL = "select top 20 * from Orders o inner join Customers c on o.CustomerID=c.CustomerID";
				new SqlDataAdapter(SQL, cnQry).Fill(dt);

				Console.WriteLine("筆數==>" + dt.Rows.Count);
				for (int x = 0; x < dt.Rows.Count; x++) {
					if (x != 0) {
						doc.NewPage();
					}

					//表頭
					PdfPTable headerTB = new PdfPTable(1);
					headerTB.TotalWidth = PageSize.A4.Width - 50;
					headerTB.LockedWidth = true;
					headerTB.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
					headerTB.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
					headerTB.DefaultCell.MinimumHeight = 35f;
					string title = "國內專利請款單確認明細報表                     所內專用，不得提供客戶                列印日期：" + DateTime.Now.ToString("yyyy/MM/dd");
					headerTB.AddCell(cell.Init(title, FontMid, PdfPCell.ALIGN_CENTER).Colspan(9));
					title = "收據營洽：m1583-王大明                             2016/11/1～2017/11/25";
					headerTB.AddCell(cell.Init(title, FontMid, PdfPCell.ALIGN_LEFT).Colspan(9));
					doc.Add(headerTB);

					doc.Add(br);

					//一筆一個table
					PdfPTable table = new PdfPTable(new float[] { 14, 10, 12, 16, 10, 8, 10, 10, 10 });
					table.TotalWidth = PageSize.A4.Width - 50;
					table.LockedWidth = true;
					table.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
					table.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
					table.DefaultCell.MinimumHeight = 25f;
					table.SplitLate = true;
					table.SplitRows = true;

					table.AddCell(new Paragraph("*請款單號:", FontSml));
					table.AddCell(new Paragraph(dt.Rows[x]["OrderID"].ToString(), FontSml));
					table.AddCell(new Paragraph("開立日期:", FontSml));
					table.AddCell(cell.Init(Convert.ToDateTime(dt.Rows[x]["OrderDate"]).ToString("yyyy/MM/dd"), FontSml).Colspan(2));
					table.AddCell(new Paragraph("請款日期:", FontSml));
					table.AddCell(cell.Init(Convert.ToDateTime(dt.Rows[x]["RequiredDate"]).ToString("yyyy/MM/dd"), FontSml).Colspan(5));

					table.AddCell(new Paragraph("*請款客戶:", FontSml));
					table.AddCell(cell.Init(dt.Rows[x]["CustomerID"].ToString() + dt.Rows[x]["CompanyName"].ToString(), FontSml).Colspan(4));
					table.AddCell(new Paragraph("聯絡人:", FontSml));
					table.AddCell(cell.Init(dt.Rows[x]["ContactName"].ToString() + "\n" + dt.Rows[x]["Phone"].ToString(), FontSml).Colspan(5));

					table.AddCell(new Paragraph("*收據抬頭:", FontSml));
					table.AddCell(cell.Init(dt.Rows[x]["CustomerID"].ToString() + dt.Rows[x]["CompanyName"].ToString(), FontSml).Colspan(4));
					table.AddCell(new Paragraph("收據種類:", FontSml));
					table.AddCell(cell.Init(dt.Rows[x]["ShipVia"].ToString(), FontSml).Colspan(5));


					table.AddCell(new Paragraph("*檢附文件:", FontSml));
					table.AddCell(cell.Init("檢附間接委辦單", FontSml).Colspan(4));
					table.AddCell(new Paragraph("開立種類:", FontSml));
					table.AddCell(cell.Init(dt.Rows[x]["City"].ToString(), FontSml).Colspan(5));

					//明細
					DataTable dt2 = new DataTable();
					SQL = "select * from [Order Details] d where d.OrderID='" + dt.Rows[x]["OrderID"] + "'";
					new SqlDataAdapter(SQL, cnQry).Fill(dt2);

					Console.WriteLine("明細筆數==>" + dt2.Rows.Count);
					if (dt2.Rows.Count > 0) {
						table.AddCell(new Paragraph("交辦單號(註記)", FontSml));
						table.AddCell(new Paragraph("契約號碼", FontSml));
						table.AddCell(new Paragraph("本所編號", FontSml));
						table.AddCell(new Paragraph("案性", FontSml));
						table.AddCell(new Paragraph("請款\n服務費", FontSml));
						table.AddCell(new Paragraph("請款\n規費", FontSml));
						table.AddCell(new Paragraph("入帳\n服務費", FontSml));
						table.AddCell(new Paragraph("入帳\n規費", FontSml));
						table.AddCell(new Paragraph("轉帳\n費用", FontSml));
					}
					
					for (int z = 0; z < dt2.Rows.Count; z++) {
						table.AddCell(new Paragraph(dt2.Rows[z]["ProductID"].ToString(), FontSml));
						table.AddCell(new Paragraph(dt2.Rows[z]["UnitPrice"].ToString(), FontSml));
						table.AddCell(cell.Init(dt2.Rows[z]["Quantity"].ToString(), FontSml).Colspan(3));
						table.AddCell(cell.Init(dt2.Rows[z]["Discount"].ToString(), FontSml).Colspan(4));
					}
					doc.Add(table);

					dt2.Dispose();
				}

				dt.Dispose();
			}

			doc.Close();
			
			using (FileStream fs = File.Create(outputFile)) {
				fs.Write(MS.GetBuffer(), 0, MS.GetBuffer().Length);
			}
			
		}
		#endregion

	}

}

