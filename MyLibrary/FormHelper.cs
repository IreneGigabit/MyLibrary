using System;
using System.Collections;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;

namespace MyLibrary {
	class SelectItem {
		public string Disp { get; set; }
		public string Value { get; set; }
		public string Attr { get; set; }

		public SelectItem(string sVal, string sDisp) : this(sVal, sDisp, false) { }

		public SelectItem(string sVal, string sDisp, bool bSelected) {
			this.Disp = sDisp;
			this.Value = sVal;
			this.Attr = "";

			if ((this.Value ?? "") == "") this.Attr = " style=\"color:blue\"";
			if (bSelected) this.Attr += " selected=\"selected\"";
		}
	}

	public static class FormHelper {
		/// <summary>
		/// 設定Repeater內容
		/// </summary>
		/// <param name="connString">連線字串</param>
		/// <param name="sql">抓取資料SQL</param>
		/// <returns></returns>
		public static Repeater SetItem(this Repeater rptr, string connString, string sql) {
			return rptr.SetItem(connString, sql, false, null);
		}

		/// <summary>
		/// 設定Repeater內容
		/// </summary>
		/// <param name="connString">連線字串</param>
		/// <param name="sql">抓取資料SQL</param>
		/// <param name="defaultVal">預設值</param>
		/// <returns></returns>
		public static Repeater SetItem(this Repeater rptr, string connString, string sql, string defaultVal) {
			return rptr.SetItem(connString, sql, false, defaultVal);
		}

		/// <summary>
		/// 設定Repeater內容
		/// </summary>
		/// <param name="connString">連線字串</param>
		/// <param name="sql">抓取資料SQL</param>
		/// <param name="haveChoice">是否有「請選擇...」</param>
		/// <returns></returns>
		public static Repeater SetItem(this Repeater rptr, string connString, string sql, bool haveChoice) {
			return rptr.SetItem(connString, sql, haveChoice, null);
		}

		/// <summary>
		/// 設定Repeater內容
		/// </summary>
		/// <param name="connString">連線字串</param>
		/// <param name="sql">抓取資料SQL</param>
		/// <param name="haveChoice">是否有「請選擇...」</param>
		/// <param name="defaultVal">預設值</param>
		/// <returns></returns>
		public static Repeater SetItem(this Repeater rptr, string connString, string sql, bool haveChoice, string defaultVal) {
			ArrayList SelAry = new ArrayList();

			using (SqlConnection cn = new SqlConnection(connString)) {
				cn.Open();
				SqlCommand cmd = new SqlCommand(sql, cn);
				SqlDataReader dr = cmd.ExecuteReader();

				if (haveChoice) {
					SelAry.Add(new SelectItem("", "請選擇...", true));
				}

				while (dr.Read()) {
					if (defaultVal == dr[0].ToString()) {
						SelAry.Add(new SelectItem(dr[0].ToString(), dr[1].ToString(), true));
					} else {
						SelAry.Add(new SelectItem(dr[0].ToString(), dr[1].ToString()));
					}
				}
				dr.Close();
			}

			rptr.DataSource = SelAry;
			rptr.DataBind();

			return rptr;
		}
	}


}
