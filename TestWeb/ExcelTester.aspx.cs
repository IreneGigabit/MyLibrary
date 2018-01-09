using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using MyLibrary;

namespace TestWeb {
	public partial class ExcelTester : System.Web.UI.Page {
		protected void Page_Load(object sender, EventArgs e) {
			if (IsPostBack) {
				//Response.CacheControl = "no-cache";
				Response.CacheControl = "Private";
				Response.AddHeader("Pragma", "no-cache");
				Response.Expires = -1;
				
				//
				if ((Request["btn1"] ?? "") == "[下載彙總表(SheetHelper)]")
					ExcelHelperOut();
				else if ((Request["btn2"] ?? "") == "[下載彙總表(SheetExt)]")
					ExcelExtOut();
			}
		}

		private void ExcelExtOut() {
			string savePath = Server.MapPath("~/Template") + "/發票明細表Template.xls";
			string inv_dateS = Request["inv_dateS"] ?? "";
			string inv_dateE = Request["inv_dateE"] ?? "";
			string dept = Request["dept"] ?? "";
			string inv_class = Request["inv_class"] ?? "";

			FileStream file = new FileStream(savePath, FileMode.Open, FileAccess.Read);
			IWorkbook workbook = new HSSFWorkbook(file);
			ISheet outputSheet = workbook.GetSheetAt(0);
			ISheet templateSheet = workbook.GetSheetAt(1); //樣本格式來源sheet

			//處理報表抬頭部分
			string rptTitle = "";
			rptTitle += "◎發票日期：" + inv_dateS + "~" + inv_dateE;
			if (dept == "P") rptTitle += " ◎部門別：專利";
			if (dept == "T") rptTitle += " ◎部門別：商標";
			if (inv_class == "A") rptTitle += " ◎發票種類：電子計算機發票";
			if (inv_class == "E") rptTitle += " ◎發票種類：代收代付虛擬發票";

			//outputSheet.SetValue(0, 0, (Session["SeBranchName"] ?? "").ToString() + "發票明細表(SheetExt)");
			outputSheet.Pos(0, 0).SetValue((Session["SeBranchName"] ?? "").ToString() + "發票明細表(SheetExt)");
			outputSheet.Pos(1, 0).SetValue(rptTitle + Server.MapPath("~/OutReport"));
			outputSheet.Pos(2, 0).SetValue("◎列印日期：" + DateTime.Now.ToString("yyyy/MM/dd"));

			int irow = 4; //開始處理行數(從0開始)
			int dbCnt = 0; //資料筆數
			//Boolean _isFirst = true; //在處理第一資料時，要額外做其它事情

			//取得報表資料
			DataTable tbData = getData();

			foreach (DataRow dr in tbData.Rows) {
				dbCnt++;

				//顯示明細
				if (irow % 2 == 0)//從樣版複製明細行(連樣式一起複製),單雙行不同色
					outputSheet.Row(irow).CopyRow(templateSheet, 6);
				else
					outputSheet.Row(irow).CopyRow(templateSheet, 5);

				outputSheet.SetValue(irow, 0, dr["inv_no"].ToString());//發票號碼
				outputSheet.SetValue(irow, 1, dr["inv_date"].ToString());//發票日期
				outputSheet.SetValue(irow, 2, dr["inv_id"] + "_" + dr["ap_cname1"].ToString().PadRight(6, ' ').Substring(0, 6));//發票抬頭
				outputSheet.SetValue(irow, 3, dr["db_no"].ToString());//請款單號
				outputSheet.SetValue(irow, 4, double.Parse(dr["tot_money"].ToString()));//發票金額
				outputSheet.SetValue(irow, 5, double.Parse(dr["tax_money"].ToString()));//發票稅額
				outputSheet.SetValue(irow, 6, double.Parse(dr["inv_money"].ToString()));//發票總額
				outputSheet.SetValue(irow, 7, dr["inv_class_name"].ToString().Substring(0, 2));//種類
				outputSheet.SetValue(irow, 8, dr["tran_name"].ToString());//異動

				irow++;
			}

			//合計處理
			if (tbData.Rows.Count > 0) {
				//outputSheet.CopyRow(irow, templateSheet, 7);//複製合計行
				outputSheet.Row(irow).CopyRow(templateSheet, 7);//複製合計行

				outputSheet.SetValue(irow, 0, "◎合計筆數 : " + dbCnt + " 筆");
				double t_tot_money = (tbData.Compute("sum(tot_money)", string.Empty) is DBNull ? 0 : Convert.ToInt32(tbData.Compute("sum(tot_money)", string.Empty)));
				outputSheet.SetValue(irow, 4, t_tot_money);
				double t_tax_money = (tbData.Compute("sum(tax_money)", string.Empty) is DBNull ? 0 : Convert.ToInt32(tbData.Compute("sum(tax_money)", string.Empty)));
				outputSheet.SetValue(irow, 5, t_tax_money);
				double t_inv_money = (tbData.Compute("sum(inv_money)", string.Empty) is DBNull ? 0 : Convert.ToInt32(tbData.Compute("sum(inv_money)", string.Empty)));
				outputSheet.SetValue(irow, 6, t_inv_money);

				irow++;

				//計算作廢
				double x_count = (tbData.Compute("count(xmark)", "xmark='*'") is DBNull ? 0 : Convert.ToInt32(tbData.Compute("count(xmark)", "xmark='*'")));
				if (x_count > 0) {
					//outputSheet.CopyRow(irow, templateSheet, 8);//複製作廢行
					outputSheet.Row(irow).CopyRow(templateSheet, 8);//複製作廢行
					outputSheet.SetValue(irow, 0, "※作廢筆數 : " + dbCnt + " 筆");
					double x_tot_money = (tbData.Compute("sum(tot_money)", "xmark='*'") is DBNull ? 0 : Convert.ToInt32(tbData.Compute("sum(tot_money)", "xmark='*'")));
					outputSheet.SetValue(irow, 4, x_tot_money);
					double x_tax_money = (tbData.Compute("sum(tax_money)", "xmark='*'") is DBNull ? 0 : Convert.ToInt32(tbData.Compute("sum(tax_money)", "xmark='*'")));
					outputSheet.SetValue(irow, 5, x_tax_money);
					double x_inv_money = (tbData.Compute("sum(inv_money)", "xmark='*'") is DBNull ? 0 : Convert.ToInt32(tbData.Compute("sum(inv_money)", "xmark='*'")));
					outputSheet.SetValue(irow, 6, x_inv_money);

					irow++;
				}
			}

			workbook.RemoveSheetAt(1); //移除樣本頁籤

			//== 輸出Excel 2003檔案。==============================
			MemoryStream MS = new MemoryStream();
			workbook.Write(MS);
			//存一份在主機
			using (FileStream fs = File.Create(Server.MapPath("~/OutReport") + "\\發票明細表.xls")) {
				fs.Write(MS.GetBuffer(), 0, MS.GetBuffer().Length);
			}
			//== Excel檔名，請寫在最後面 filename的地方
			Response.AddHeader("Content-Disposition", "attachment; filename=\"" + HttpUtility.UrlEncode("發票明細表", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls\"");
			Response.BinaryWrite(MS.ToArray());
			//== 釋放資源
			workbook = null;
			MS.Close();
			MS.Dispose();
		}


		private void ExcelHelperOut() {
			string savePath = Server.MapPath("~/Template") + "/發票明細表Template.xls";
			string inv_dateS = Request["inv_dateS"] ?? "";
			string inv_dateE = Request["inv_dateE"] ?? "";
			string dept = Request["dept"] ?? "";
			string inv_class = Request["inv_class"] ?? "";

			FileStream file = new FileStream(savePath, FileMode.Open, FileAccess.Read);
			IWorkbook workbook = new HSSFWorkbook(file);
			ISheet outputSheet = workbook.GetSheetAt(0);
			ISheet templateSheet = workbook.GetSheetAt(1); //樣本格式來源sheet
			SheetHelper oXls = new SheetHelper(outputSheet);

			//處理報表抬頭部分
			string rptTitle = "";
			rptTitle += "◎發票日期：" + inv_dateS + "~" + inv_dateE;
			if (dept == "P") rptTitle += " ◎部門別：專利";
			if (dept == "T") rptTitle += " ◎部門別：商標";
			if (inv_class == "A") rptTitle += " ◎發票種類：電子計算機發票";
			if (inv_class == "E") rptTitle += " ◎發票種類：代收代付虛擬發票";

			oXls.Pos(0, 0).SetValue((Session["SeBranchName"] ?? "").ToString() + "發票明細表(SheetHelper)");
			oXls.Pos(1, 0).SetValue(rptTitle);
			oXls.Pos(2, 0).SetValue("◎列印日期：" + DateTime.Now.ToString("yyyy/MM/dd"));

			int irow = 4; //開始處理行數(從0開始)
			int dbCnt = 0; //資料筆數
			//Boolean _isFirst = true; //在處理第一資料時，要額外做其它事情

			//取得報表資料
			DataTable tbData = getData();

			foreach (DataRow dr in tbData.Rows) {
				dbCnt++;

				//顯示明細
				if (irow % 2 == 0)//從樣版複製明細行(連樣式一起複製),單雙行不同色
					oXls.Row(irow).CopyRow(templateSheet, 6);
				else
					oXls.Row(irow).CopyRow(templateSheet, 5);

				oXls.Pos(irow, 0).SetValue(dr["inv_no"].ToString());//發票號碼
				oXls.Pos(irow, 1).SetValue(dr["inv_date"].ToString());//發票日期
				oXls.Pos(irow, 2).SetValue(dr["inv_id"] + "_" + dr["ap_cname1"].ToString().PadRight(6, ' ').Substring(0, 6));//發票抬頭
				oXls.Pos(irow, 3).SetValue(dr["db_no"].ToString());//請款單號
				oXls.Pos(irow, 4).SetValue(double.Parse(dr["tot_money"].ToString()));//發票金額
				oXls.Pos(irow, 5).SetValue(double.Parse(dr["tax_money"].ToString()));//發票稅額
				oXls.Pos(irow, 6).SetValue(double.Parse(dr["inv_money"].ToString()));//發票總額
				oXls.Pos(irow, 7).SetValue(dr["inv_class_name"].ToString().Substring(0, 2));//種類
				oXls.Pos(irow, 8).SetValue(dr["tran_name"].ToString());//異動

				irow++;
			}

			//合計處理
			if (tbData.Rows.Count > 0) {
				oXls.Row(irow).CopyRow(templateSheet, 7);//複製合計行

				oXls.Pos(irow, 0).SetValue("◎合計筆數 : " + dbCnt + " 筆");
				double t_tot_money = (tbData.Compute("sum(tot_money)", string.Empty) is DBNull ? 0 : Convert.ToInt32(tbData.Compute("sum(tot_money)", string.Empty)));
				oXls.Pos(irow, 4).SetValue(t_tot_money);
				double t_tax_money = (tbData.Compute("sum(tax_money)", string.Empty) is DBNull ? 0 : Convert.ToInt32(tbData.Compute("sum(tax_money)", string.Empty)));
				oXls.Pos(irow, 5).SetValue(t_tax_money);
				double t_inv_money = (tbData.Compute("sum(inv_money)", string.Empty) is DBNull ? 0 : Convert.ToInt32(tbData.Compute("sum(inv_money)", string.Empty)));
				oXls.Pos(irow, 6).SetValue(t_inv_money);

				irow++;

				//計算作廢
				double x_count = (tbData.Compute("count(xmark)", "xmark='*'") is DBNull ? 0 : Convert.ToInt32(tbData.Compute("count(xmark)", "xmark='*'")));
				if (x_count > 0) {
					oXls.Row(irow).CopyRow(templateSheet, 8);//複製作廢行
					oXls.Pos(irow, 0).SetValue("※作廢筆數 : " + dbCnt + " 筆");
					double x_tot_money = (tbData.Compute("sum(tot_money)", "xmark='*'") is DBNull ? 0 : Convert.ToInt32(tbData.Compute("sum(tot_money)", "xmark='*'")));
					oXls.Pos(irow, 4).SetValue(x_tot_money);
					double x_tax_money = (tbData.Compute("sum(tax_money)", "xmark='*'") is DBNull ? 0 : Convert.ToInt32(tbData.Compute("sum(tax_money)", "xmark='*'")));
					oXls.Pos(irow, 5).SetValue(x_tax_money);
					double x_inv_money = (tbData.Compute("sum(inv_money)", "xmark='*'") is DBNull ? 0 : Convert.ToInt32(tbData.Compute("sum(inv_money)", "xmark='*'")));
					oXls.Pos(irow, 6).SetValue(x_inv_money);

					irow++;
				}
			}


			workbook.RemoveSheetAt(1); //移除樣本頁籤

			//== 輸出Excel 2003檔案。==============================
			MemoryStream MS = new MemoryStream();
			workbook.Write(MS);
			//== Excel檔名，請寫在最後面 filename的地方
			Response.AddHeader("Content-Disposition", "attachment; filename=\"" + HttpUtility.UrlEncode("發票明細表", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls\"");
			Response.BinaryWrite(MS.ToArray());
			//== 釋放資源
			workbook = null;
			MS.Close();
			MS.Dispose();
		}

		private DataTable getData() {
			string SQL = "";

			SQL = "Select top 100 im.* ,b.ap_cname1 ,b.ap_cname2,''inv_class_name,''tran_name,''xmark ";
			SQL += "from invmain im ";
			SQL += "inner join sindbs.dbo.apcust as b on im.inv_id = b.apcust_no where 1=1";
			SQL += " and im.branch ='N' ";
			SQL += "order by inv_no";

			DataTable dt = new DataTable();
			using (SqlConnection cnQry = new SqlConnection("Server=web08;Database=account;User ID=web_usr;Password=web1823")) {
				cnQry.Open();
				SqlCommand cmdQry = new SqlCommand(SQL, cnQry);
				new SqlDataAdapter(SQL, cnQry).Fill(dt);

				for (int x = 0; x < dt.Rows.Count; x++) {
					//發票種類
					switch (dt.Rows[x]["inv_class"].ToString().Trim().ToUpper()) {
						case "31":
							dt.Rows[x]["inv_class_name"] = "3聯式電子計算機發票";
							break;
						case "32":
							dt.Rows[x]["inv_class_name"] = "2聯式電子計算機發票";
							break;
						case "2":
							dt.Rows[x]["inv_class_name"] = "2聯式";
							break;
						case "3":
							dt.Rows[x]["inv_class_name"] = "3聯式";
							break;
						case "E":
							dt.Rows[x]["inv_class_name"] = "代收代付虛擬發票";
							break;
						default:
							break;
					}
					//異動狀態
					switch (dt.Rows[x]["tran_code"].ToString().Trim().ToUpper()) {
						case "U":
							dt.Rows[x]["tran_name"] = "修改";
							break;
						case "C":
							dt.Rows[x]["tran_name"] = "銷折";
							break;
						case "D":
							dt.Rows[x]["tran_name"] = "銷退";
							break;
						case "E":
							dt.Rows[x]["tran_name"] = "部份銷退";
							break;
						case "F":
							dt.Rows[x]["xmark"] = "*";
							dt.Rows[x]["tran_name"] = "跨月";
							break;
						case "G":
							dt.Rows[x]["xmark"] = "*";
							dt.Rows[x]["tran_name"] = "專案";
							break;
						case "H":
							dt.Rows[x]["tran_name"] = "呆帳";
							break;
						case "X":
							dt.Rows[x]["xmark"] = "*";
							dt.Rows[x]["tran_name"] = "作廢";
							break;
						case "N":
							dt.Rows[x]["tran_name"] = "";
							break;
						default:
							break;
					}
				}
			}

			return dt;
		}

	}
}