using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MyLibrary
{
    public static class NpoiHelper
    {
        /// <summary>
        /// 創造 Excel 檔
        /// </summary>
        public static void ExportExcel<T>(NpoiParam<T> p)
        {
            //依情況決定要建新的 Sheet 或是用舊的 (即來自範本)
            ISheet sht = GetSheet(p);

            if (p.Data.Any())
            {
                //有資料塞格子
                SetSheetValue(ref p, ref sht);
            }
            else
            {
                //若沒資料在起點寫入No Data !
                CreateNewRowOrNot(ref sht, p.RowStartFrom, p.ColumnStartFrom);
                sht.GetRow(p.RowStartFrom).CreateCell(p.ColumnStartFrom).SetCellValue("No Data !");
            }

            Flush(p.Workbook, p.FileFullName);
        }

        private static ISheet GetSheet<T>(NpoiParam<T> param)
        {
            //在 workbook 中以 sheet name 尋找 是否找得到sheet
            if (param.Workbook.GetSheet(param.SheetName) == null)
            {
                //找不到建一張新的sheet
                ISheet sht = param.Workbook.CreateSheet(param.SheetName);
                sht = CreateColumn(param);

                return sht;
            }
            else
            {
                //找得到即為要塞值的目標
                return param.Workbook.GetSheet(param.SheetName);
            }
        }

        private static ISheet CreateColumn<T>(NpoiParam<T> p)
        {
            var sht = p.Workbook.GetSheet(p.SheetName);
            sht.CreateRow(0);
            ICellStyle columnStyle = GetBaseCellStyle(p.Workbook, p.FontStyle);

            if (p.ShowHeader)
            {
                for (int i = 0; i < p.ColumnMapping.Count; i++)
                {
                    var offset = i + p.ColumnStartFrom;

                    sht.GetRow(0).CreateCell(offset);                                                //先創建格子
                    sht.GetRow(0).GetCell(offset).CellStyle = columnStyle;                           //綁定基本格式
                    sht.GetRow(0).GetCell(offset).SetCellValue(p.ColumnMapping[i].ExcelColumnName);  //給值
                }
            }

            return sht;
        }

        private static void SetSheetValue<T>(ref NpoiParam<T> p, ref ISheet sht)
        {
            //要從哪一行開始塞資料 (有可能自定範本 可能你原本範本內就有好幾行表頭 2行 3行...)
            int line = p.RowStartFrom;

            //有可能前面幾欄是自訂好的 得跳過幾個欄位再開始塞
            int columnOffset = p.ColumnStartFrom;

            //根據標題列先處理所有Style (對Npoi來說 '創建'Style在workbook中是很慢的操作 作越少次越好 絕對不要foreach在塞每行列實際資料時重覆作 只通通在標題列做一次就好)
            ICellStyle[] cellStyleArr = InitialColumnStyle(p.Workbook, p.ColumnMapping, p.FontStyle);

            foreach (var item in p.Data)
            {
                //如果 x 軸有偏移值 則表示這行他已經自己建了某幾欄的資料 我們只負責塞後面幾欄 所以並非每次都create new row
                CreateNewRowOrNot(ref sht, line, columnOffset);

                for (int i = 0; i < p.ColumnMapping.Count; i++)
                {
                    //建立格子 (需考量 x 軸有偏移值)
                    var cell = sht.GetRow(line).CreateCell(i + columnOffset);

                    //綁定style (記得 綁定是不慢的 但建新style是慢的 不要在迴圈裡無意義的反覆建style 只在標題處理一次即可)
                    cell.CellStyle = cellStyleArr[i];

                    //給值
                    string value = GetValue(item, p.ColumnMapping, i);               //reflection取值
                    SetCellValue(value, ref cell, p.ColumnMapping[i].DataType);      //幫cell填值
                }

                line++;
            }

            //處理AutoFit (必定是在最後做的 因為你得把所有格子都塞完以後才知道每欄多寬是你需要的)
            if (p.IsAutoFit)
            {
                for (int i = 0; i < p.ColumnMapping.Count; i++)
                {
                    sht.AutoSizeColumn(i);
                }
            }
        }

        private static ICellStyle[] InitialColumnStyle(IWorkbook wb, List<ColumnMapping> columnMapping, FontStyle fontStyle)
        {
            ICellStyle[] styleArr = new ICellStyle[columnMapping.Count];

            for (int i = 0; i < columnMapping.Count; i++)
            {
                //取通用格式
                ICellStyle cellStyle = GetBaseCellStyle(wb, fontStyle);

                //處理格式輸出
                if (!String.IsNullOrEmpty(columnMapping[i].Format))
                {
                    cellStyle.DataFormat = GetCellFormat(wb, columnMapping[i].Format);
                }

                styleArr[i] = cellStyle;
            }

            return styleArr;
        }

        private static void CreateNewRowOrNot(ref ISheet sht, int line, int columnOffset)
        {
            //如果是從自定範本來則不能重畫格子 例如他給我範本 只要我畫後面三格 前兩格他自己做好了 如果我整行重畫 他自己畫的兩格也會消失
            if (columnOffset == 0 || line > sht.LastRowNum)
            {
                sht.CreateRow(line);
            }
        }

        private static void SetCellValue(string value, ref ICell cell, NpoiDataType type)
        {
            switch (type)
            {
                //字串沒有格式
                case NpoiDataType.String:
                    if (!String.IsNullOrEmpty(value)) cell.SetCellValue(value);
                    break;

                //轉日期
                case NpoiDataType.Date:
                    if (!String.IsNullOrEmpty(value)) cell.SetCellValue(Convert.ToDateTime(value));
                    break;

                //轉數字
                case NpoiDataType.Number:
                    if (!String.IsNullOrEmpty(value)) cell.SetCellValue(Convert.ToDouble(value));
                    break;

                //不會發生;
                default:
                    break;
            }
        }

        private static ICellStyle GetBaseCellStyle(IWorkbook wb, FontStyle fontStyle)
        {
            //畫線
            ICellStyle cellStyle = wb.CreateCellStyle();
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;

            //預設字型大小
            IFont font1 = wb.CreateFont();
            font1.FontName = (fontStyle.FontName == null) ? "Arial" : fontStyle.FontName;
            font1.FontHeightInPoints = (fontStyle.FontHeightInPoints == null) ? (short)10 : fontStyle.FontHeightInPoints.Value;
            cellStyle.SetFont(font1);

            return cellStyle;
        }

        private static short GetCellFormat(IWorkbook wb, string formatStr)
        {
            IDataFormat dataFormat = wb.CreateDataFormat();
            return dataFormat.GetFormat(formatStr);
        }

        private static string GetValue<T>(T obj, List<ColumnMapping> columnMapping, int order)
        {
            var fieldName = columnMapping[order].ModelFieldName;
            var prop = typeof(T).GetProperties().Where(q => q.Name == fieldName).First();
            var value = prop.GetValue(obj, null);

            return (value == null) ? "" : value.ToString();
        }

        private static void Flush(IWorkbook wb, string fullName)
        {
            //若檔案已存在先刪除
            if (File.Exists(fullName)) File.Delete(fullName);

            using (FileStream targetFs = File.Create(fullName))
            {
                wb.Write(targetFs);
                wb = null;
            }
        }
    }
}
