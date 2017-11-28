using System;
using System.Collections.Generic;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Diagnostics;
using MyLibrary;

namespace TestConsole
{
    class PdfPCellExtTester
    {
        private static string CurrDir = System.Environment.CurrentDirectory;
        static string templateFile = CurrDir + @"\Credit_letter_apply_one.pdf";
        //static string templateFile = CurrDir + @"\pdf29.pdf";
        static string outputFile = CurrDir + @"\new.pdf";


        //static string fontPath = Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\..\Fonts\kaiu.ttf";//標楷體
        //static string fontPath = Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\..\Fonts\mingliu.ttc,0";//細明體
        //static string fontPath = Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\..\Fonts\msjh.ttc,0";//正黑體
        static string fontPath = Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\..\Fonts\msjh.ttf";//正黑體

        static void Main(string[] args)
        {
            Console.WriteLine("Template File:" + templateFile);
            Console.WriteLine("Output File:" + outputFile);

            //multiplePage();
            tablePage();
            //memorystream();
            //filestream();

            Process.Start(outputFile);

            //Console.WriteLine("請按任一鍵離開..");
            //Console.ReadKey();
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

            PdfPCell cell = new PdfPCell();
            table.AddCell(cell.Init("C00)發票作廢、退貨", FontSml).Colspan(2));
            table.AddCell(cell.Init("C05)滯銷", FontSml).Colspan(3));
            table.AddCell(cell.Init("C10)結束營業", FontSml).Colspan(4));
            table.AddCell(cell.Init("", FontSml).Colspan(2));

            table.AddCell(cell.Init("C01)業代誤訂", FontSml).Colspan(2));
            table.AddCell(cell.Init("C06)壞損-於配送途中", FontSml).Colspan(3));
            table.AddCell(cell.Init("C99)其他", FontSml).Colspan(4));
            table.AddCell(cell.Init("", FontSml).Colspan(2));

            PdfPTable table1 = new PdfPTable(new float[] { 40f, 60f });
            table1.AddCell("C02)客戶誤訂");
            table1.AddCell("C07)壞損-客戶收貨後");
            table1.AddCell("C03)久裕誤寄");
            table1.AddCell("C08)過期");
            table.AddCell(cell.Init(table1).Colspan(5));
            table.AddCell(cell.Init("RTN_REMARK", FontSml, PdfPCell.ALIGN_CENTER, PdfPCell.ALIGN_MIDDLE).Colspan(6));

            table.AddCell(cell.Init("C04)配送公司誤送", FontSml).Colspan(2));
            table.AddCell(cell.Init("C09)短效期", FontSml).Colspan(9).Border(0));

            table.AddCell(cell.Init("列印日期：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), FontSml).Colspan(11));

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

    }


}
