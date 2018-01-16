using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Web;
using MyLibrary;

/// <summary>
/// 產生智慧局電子申請書用
/// </summary>
public class IpoReport : OpenXmlHelper {
	protected string _connStr = null;
	protected string _in_no = "";
	protected string _in_scode = "";
	protected string _branch = "";
	protected string _receipt_title = "";
	protected DBHelper _conn = null;
	protected DataTable _dtDmp = null;
	protected DataTable _dtApcust = null;
	protected DataTable _dtAgt = null;
	protected DataTable _dtAnt = null;

	/// <summary>
	/// 組合後的本所編號
	/// </summary>
	public string Seq = "";

	/// <summary>
	/// 收據抬頭
	/// </summary>
	public string RectitleName = "";

	/// <summary>
	/// 案件資料
	/// </summary>
	public DataTable Dmp {
		get { return _dtDmp; }
	}

	/// <summary>
	/// 申請人資料
	/// </summary>
	public DataTable Apcust {
		get { return _dtApcust; }
	}

	/// <summary>
	/// 代理人資料
	/// </summary>
	public DataTable Agent {
		get { return _dtAgt; }
	}

	/// <summary>
	/// 發明人/新型創作/設計人資料
	/// </summary>
	public DataTable Ant {
		get { return _dtAnt; }
	}

	public IpoReport(string connStr, string in_scode, string in_no, string branch, string receipt_title) {
		this._connStr = connStr;
		this._in_no = in_no;
		this._in_scode = in_scode;
		this._branch = branch;
		this._receipt_title = receipt_title;
		this._conn = new DBHelper(connStr, false).Debug(false);

		this._dtDmp = new DataTable();
		_conn.DataTable("select * from vdmpall where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'", _dtDmp);//抓案件資料
		this.Seq = getSeq();//組案件編號
		this.RectitleName = getRectitleName();//收據抬頭
		this._dtApcust = GetApcust();//抓申請人
		this._dtAgt = GetAgent();//抓代理人
		this._dtAnt = GetAnt();//抓發明人/新型創作/設計人
	}

	#region 關閉
	/// <summary>
	/// 關閉
	/// </summary>
	public void CloseRpt() {
		_conn.Dispose();
		this.Dispose();
	}
	#endregion

	#region 取得組合後的本所編號
	/// <summary>
	/// 取得組合後的本所編號
	/// </summary>
	private string getSeq() {
		string lseq = _branch + "P" + _dtDmp.Rows[0]["seq"];
		if (_dtDmp.Rows[0]["seq1"].ToString() != "_") {
			lseq += "-" + _dtDmp.Rows[0]["seq1"];
		}
		return lseq;
	}
	#endregion

	#region 取得申請人資料
	/// <summary>
	/// 取得申請人資料
	/// </summary>
	private DataTable GetApcust() {
		string SQL = "select b.ap_zip,b.ap_crep,b.ap_erep,b.ap_eaddr1,b.ap_eaddr2,b.ap_eaddr3,b.ap_eaddr4,b.ap_addr1,b.ap_addr2 " +
		",b.apcust_no,b.apclass,b.ap_country,a.ap_cname1,a.ap_cname2,b.ap_ename1,b.ap_ename2,b.ap_fcname,b.ap_lcname,b.ap_fename,b.ap_lename " +
		",''apclass_name,''Country_name,''Title_cname,''Cname_string,''Title_ename,''Ename_string " +
		" from dmp_apcust a " +
		" inner join apcust b on a.apsqlno=b.apsqlno " +
		" where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'  and kind='A'";
		DataTable dt = new DataTable();
		_conn.DataTable(SQL, dt);

		for (int i = 0; i < dt.Rows.Count; i++) {
			SQL = " select  isnull(b.coun_code,'')+isnull(b.coun_cname,'') Country_name " +
				"From sysctrl.dbo.country a inner join sysctrl.dbo.IPO_country b on a.coun_code=b.ref_coun_code " +
			 " where a.coun_code = '" + dt.Rows[i]["ap_country"] + "'";
			dt.Rows[i]["Country_name"] = (_conn.ExecuteScalar(SQL) ?? "").ToString();

			//本國公司
			if (dt.Rows[i]["apclass"].ToString().Left(1) == "A") {
				if (dt.Rows[i]["apclass"].ToString() == "AD") {
					dt.Rows[i]["apclass_name"] = "商號行號工廠";
				} else {
					dt.Rows[i]["apclass_name"] = "法人公司機關學校";
				}
				dt.Rows[i]["Title_cname"] = "中文名稱";
				dt.Rows[i]["Title_ename"] = "英文名稱";
				dt.Rows[i]["Cname_string"] = dt.Rows[i]["ap_cname1"].ToString().Trim() + dt.Rows[i]["ap_cname2"].ToString().Trim();
				dt.Rows[i]["Ename_string"] = dt.Rows[i]["ap_ename1"].ToString().Trim() + dt.Rows[i]["ap_ename2"].ToString().Trim();
			}
			//本國自然人
			if (dt.Rows[i]["apclass"].ToString().Left(1) == "B") {
				dt.Rows[i]["apclass_name"] = "自然人";
				dt.Rows[i]["Title_cname"] = "中文姓名";
				dt.Rows[i]["Title_ename"] = "英文姓名";
				dt.Rows[i]["Cname_string"] = dt.Rows[i]["ap_fcname"].ToString().Trim() + "," + dt.Rows[i]["ap_lcname"].ToString().Trim();
				if (dt.Rows[i]["Cname_string"].ToString() == ",") {
					dt.Rows[i]["Cname_string"] = dt.Rows[i]["ap_cname1"].ToString().Trim() + dt.Rows[i]["ap_cname2"].ToString().Trim();
				}

				dt.Rows[i]["Ename_string"] = dt.Rows[i]["ap_fename"].ToString().Trim() + "," + dt.Rows[i]["ap_lename"].ToString().Trim();
				if (dt.Rows[i]["Ename_string"].ToString() == ",") {
					dt.Rows[i]["Ename_string"] = dt.Rows[i]["ap_ename1"].ToString().Trim() + dt.Rows[i]["ap_ename2"].ToString().Trim();
				}
			}

			//20161206外國人/公司增加判斷-若有填寫申請人姓&名,則顯示姓名
			if (dt.Rows[i]["apclass"].ToString().Left(1) == "C") {
				dt.Rows[i]["Cname_string"] = dt.Rows[i]["ap_fcname"].ToString().Trim() + "," + dt.Rows[i]["ap_lcname"].ToString().Trim();
				if (dt.Rows[i]["Cname_string"].ToString() == ",") {
					dt.Rows[i]["Title_cname"] = "中文名稱";
					dt.Rows[i]["Title_ename"] = "英文名稱";
					dt.Rows[i]["Cname_string"] = dt.Rows[i]["ap_cname1"].ToString().Trim() + dt.Rows[i]["ap_cname2"].ToString().Trim();
				} else {
					dt.Rows[i]["Title_cname"] = "中文姓名";
					dt.Rows[i]["Title_ename"] = "英文姓名";
				}

				dt.Rows[i]["Ename_string"] = dt.Rows[i]["ap_fename"].ToString().Trim() + "," + dt.Rows[i]["ap_lename"].ToString().Trim();
				if (dt.Rows[i]["Ename_string"].ToString() == ",") {
					dt.Rows[i]["Ename_string"] = dt.Rows[i]["ap_ename1"].ToString().Trim() + dt.Rows[i]["ap_ename2"].ToString().Trim();
				}
			}
		}
		return dt;
	}
	#endregion

	#region 取得代理人資料
	/// <summary>
	/// 取得代理人資料
	/// </summary>
	private DataTable GetAgent() {
		string SQL = " Select b.agt_fax,b.agt_tel,b.agt_addr,b.agt_zip,b.agt_id1,b.agt_id2,b.agt_idno1,b.agt_idno2,b.agt_name1,b.agt_name2 from dmp a " +
		" inner join vdmpall c on a.dmp_sqlno = c.dmp_sqlno " +
		" inner join agt b on c.nagt_no = b.agt_no " +
		" where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'";
		DataTable dt = new DataTable();
		_conn.DataTable(SQL, dt);

		for (int i = 0; i < dt.Rows.Count; i++) {
			dt.Rows[i]["agt_idno1"] = dt.Rows[i]["agt_idno1"].ToString().Trim().PadLeft(5, '0');
			dt.Rows[i]["agt_idno2"] = dt.Rows[i]["agt_idno2"].ToString().Trim().PadLeft(5, '0');
			dt.Rows[i]["agt_name1"] = dt.Rows[i]["agt_name1"].ToString().Trim().Left(1) + "," + dt.Rows[i]["agt_name1"].ToString().Trim().Substring(1);
			dt.Rows[i]["agt_name2"] = dt.Rows[i]["agt_name2"].ToString().Trim().Left(1) + "," + dt.Rows[i]["agt_name2"].ToString().Trim().Substring(1);
		}
		return dt;
	}
	#endregion

	#region 取得發明人/新型創作/設計人
	/// <summary>
	/// 取得發明人/新型創作/設計人
	/// </summary>
	private DataTable GetAnt() {
		string SQL = " Select ant_id,ant_country,ant_cname1,ant_cname2,ant_ename1,ant_ename2,ant_fcname,ant_lcname,ant_fename,ant_lename " +
			",''Country_name,''Cname_string,''Ename_string " +
			"from dmp_ant " +
			" where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'";
		DataTable dt = new DataTable();
		_conn.DataTable(SQL, dt);

		for (int i = 0; i < dt.Rows.Count; i++) {
			SQL = " select  isnull(b.coun_code,'')+isnull(b.coun_cname,'') Country_name " +
				"From sysctrl.dbo.country a inner join sysctrl.dbo.IPO_country b on a.coun_code=b.ref_coun_code " +
			 " where a.coun_code = '" + dt.Rows[i]["ant_country"] + "'";
			dt.Rows[i]["Country_name"] = (_conn.ExecuteScalar(SQL) ?? "").ToString();

			dt.Rows[i]["Cname_string"] = dt.Rows[i]["ant_fcname"].ToString().Trim() + "," + dt.Rows[i]["ant_lcname"].ToString().Trim();
			if (dt.Rows[i]["Cname_string"].ToString() == ",") {
				dt.Rows[i]["Cname_string"] = dt.Rows[i]["ant_cname1"].ToString().Trim() + dt.Rows[i]["ant_cname2"].ToString().Trim();
			}

			dt.Rows[i]["Ename_string"] = dt.Rows[i]["ant_fename"].ToString().Trim() + "," + dt.Rows[i]["ant_lename"].ToString().Trim();
			if (dt.Rows[i]["Ename_string"].ToString() == ",") {
				dt.Rows[i]["Ename_string"] = dt.Rows[i]["ant_ename1"].ToString().Trim() + dt.Rows[i]["ant_ename2"].ToString().Trim();
			}
		}
		return dt;
	}
	#endregion

	#region 取得收據抬頭
	/// <summary>
	/// 取得收據抬頭
	/// </summary>
	private string getRectitleName() {
		string RectitleName = "";

		//申請人(只抓一個)
		string SQL = "select b.apclass,b.ap_country,a.ap_cname1,a.ap_cname2,b.ap_ename1,b.ap_ename2,b.ap_fcname,b.ap_lcname,b.ap_fename,b.ap_lename" +
					" from dmp_apcust a,apcust b " +
					" where in_scode='" + _in_scode + "' and in_no='" + _in_no + "'  and kind='A' and a.apsqlno=b.apsqlno";

		string Cname_string = "";
		using (SqlDataReader dr = _conn.ExecuteReader(SQL)) {
			if (dr.Read()) {
				//本國公司
				if (dr.GetString("apclass").Left(1) == "A") {
					Cname_string = dr.GetString("ap_cname1") + dr.GetString("ap_cname2");
				}

				//本國自然人
				if (dr.GetString("apclass").Left(1) == "B") {
					Cname_string = dr.GetString("ap_fcname") + "" + dr.GetString("ap_lcname");
					if (Cname_string == "") {
						Cname_string = dr.GetString("ap_cname1") + dr.GetString("ap_cname2");
					}
				}

				//20161206外國人/公司增加判斷-若有填寫申請人姓&名,則顯示姓名
				if (dr.GetString("apclass").Left(1) == "C") {
					Cname_string = dr.GetString("ap_fcname") + "" + dr.GetString("ap_lcname");
					if (Cname_string == "") {
						Cname_string = dr.GetString("ap_cname1") + dr.GetString("ap_cname2");
					}
				}
			}

			string receipt_title = _dtDmp.Rows[0]["receipt_title"].ToString();
			if (receipt_title == "A") {//專利權人
				RectitleName = Cname_string;
			} else if (receipt_title == "C") {//專利權人(代繳人)
				RectitleName = Cname_string + "(代繳人：聖島國際專利商標聯合事務所)";
			} else if (receipt_title == "B") {//空白
				RectitleName = "";
			}
		}

		return RectitleName;
	}
	#endregion

	#region 產生基本資料表
	/// <summary>
	/// 產生基本資料表
	/// </summary>
	public void AppendBaseData() {
		CopyBlock(baseDoc, "base_title");
		//申請人
		using (DataTable dtAp = GetApcust()) {
			for (int i = 0; i < dtAp.Rows.Count; i++) {
				CopyBlock(baseDoc, "base_apcust1");
				ReplaceBookmark("base_ap_num", (i + 1).ToString());
				ReplaceBookmark("base_ap_country", dtAp.Rows[i]["Country_name"].ToString());
				ReplaceBookmark("ap_class", dtAp.Rows[i]["apclass_name"].ToString());
				if (dtAp.Rows[i]["ap_country"].ToString() == "T") {
					CopyBlock(baseDoc, "base_apcust2");
					ReplaceBookmark("apcust_no", dtAp.Rows[i]["apcust_no"].ToString());
				}
				CopyBlock(baseDoc, "base_apcust3");
				ReplaceBookmark("base_ap_cname_title", dtAp.Rows[i]["Title_cname"].ToString());
				ReplaceBookmark("base_ap_ename_title", dtAp.Rows[i]["Title_ename"].ToString());
				ReplaceBookmark("base_ap_cname", dtAp.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
				ReplaceBookmark("base_ap_ename", dtAp.Rows[i]["Ename_string"].ToString().ToXmlUnicode());
				ReplaceBookmark("ap_live_country", dtAp.Rows[i]["Country_name"].ToString());
				ReplaceBookmark("ap_zip", dtAp.Rows[i]["ap_zip"].ToString());
				string ap_addr = dtAp.Rows[i]["ap_addr1"].ToString().ToXmlUnicode()
					+ dtAp.Rows[i]["ap_addr2"].ToString().ToXmlUnicode();
				ReplaceBookmark("ap_addr", ap_addr);
				string ap_eddr = dtAp.Rows[i]["ap_eaddr1"].ToString().ToXmlUnicode()
					+ dtAp.Rows[i]["ap_eaddr2"].ToString().ToXmlUnicode()
					+ dtAp.Rows[i]["ap_eaddr3"].ToString().ToXmlUnicode()
					+ dtAp.Rows[i]["ap_eaddr4"].ToString().ToXmlUnicode();
				ReplaceBookmark("ap_eddr", ap_eddr);
				ReplaceBookmark("ap_crep", dtAp.Rows[i]["ap_crep"].ToString().ToXmlUnicode());
				ReplaceBookmark("ap_erep", dtAp.Rows[i]["ap_erep"].ToString().ToXmlUnicode());
			}
		}
		//代理人
		CopyBlock(baseDoc,"base_agent");
		using (DataTable dtAgt = GetAgent()) {
			for (int i = 0; i < dtAgt.Rows.Count; i++) {
				CopyBlock(baseDoc,"base_apcust");
				ReplaceBookmark("agt_idno1", dtAgt.Rows[i]["agt_idno1"].ToString());
				ReplaceBookmark("agt_id1", dtAgt.Rows[i]["agt_id1"].ToString());
				ReplaceBookmark("base_agt_name1", dtAgt.Rows[i]["agt_name1"].ToString());
				ReplaceBookmark("agt_zip1", dtAgt.Rows[i]["agt_zip"].ToString());
				ReplaceBookmark("agt_addr1", dtAgt.Rows[i]["agt_addr"].ToString());
				ReplaceBookmark("agatt_tel1", dtAgt.Rows[i]["agt_tel"].ToString());
				ReplaceBookmark("agatt_fax1", dtAgt.Rows[i]["agt_fax"].ToString());
				ReplaceBookmark("agt_idno2", dtAgt.Rows[i]["agt_idno2"].ToString());
				ReplaceBookmark("agt_id2", dtAgt.Rows[i]["agt_id2"].ToString());
				ReplaceBookmark("base_agt_name2", dtAgt.Rows[i]["agt_name2"].ToString());
				ReplaceBookmark("agt_zip2", dtAgt.Rows[i]["agt_zip"].ToString());
				ReplaceBookmark("agt_addr2", dtAgt.Rows[i]["agt_addr"].ToString());
				ReplaceBookmark("agatt_tel2", dtAgt.Rows[i]["agt_tel"].ToString());
				ReplaceBookmark("agatt_fax2", dtAgt.Rows[i]["agt_fax"].ToString());
			}
		}
		//發明人/新型創作/設計人
		using (DataTable dtAnt = GetAnt()) {
			for (int i = 0; i < dtAnt.Rows.Count; i++) {
				CopyBlock(baseDoc,"base_ant1");
				ReplaceBookmark("base_ant_num", "發明人" + (i + 1).ToString());
				ReplaceBookmark("base_ant_country", dtAnt.Rows[i]["Country_name"].ToString());
				if (dtAnt.Rows[i]["ant_country"].ToString() == "T") {
					CopyBlock(baseDoc,"base_ant2");
					ReplaceBookmark("ant_id", dtAnt.Rows[i]["ant_id"].ToString());
				}
				CopyBlock(baseDoc,"base_ant3");
				ReplaceBookmark("base_ant_cname", dtAnt.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
				ReplaceBookmark("base_ant_ename", dtAnt.Rows[i]["Ename_string"].ToString().ToXmlUnicode());
			}
		}

		//AppendFoot(0);//頁尾
	}
/*
	public void AppendBaseData() {
		CopyBlock("base_title");
		//申請人
		using (DataTable dtAp = GetApcust()) {
			for (int i = 0; i < dtAp.Rows.Count; i++) {
				CopyBlock("base_apcust1");
				ReplaceBookmark("base_ap_num", (i + 1).ToString());
				ReplaceBookmark("base_ap_country", dtAp.Rows[i]["Country_name"].ToString());
				ReplaceBookmark("ap_class", dtAp.Rows[i]["apclass_name"].ToString());
				if (dtAp.Rows[i]["ap_country"].ToString() == "T") {
					CopyBlock("base_apcust2");
					ReplaceBookmark("apcust_no", dtAp.Rows[i]["apcust_no"].ToString());
				}
				CopyBlock("base_apcust3");
				ReplaceBookmark("base_ap_cname_title", dtAp.Rows[i]["Title_cname"].ToString());
				ReplaceBookmark("base_ap_ename_title", dtAp.Rows[i]["Title_ename"].ToString());
				ReplaceBookmark("base_ap_cname", dtAp.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
				ReplaceBookmark("base_ap_ename", dtAp.Rows[i]["Ename_string"].ToString().ToXmlUnicode());
				ReplaceBookmark("ap_live_country", dtAp.Rows[i]["Country_name"].ToString());
				ReplaceBookmark("ap_zip", dtAp.Rows[i]["ap_zip"].ToString());
				string ap_addr = dtAp.Rows[i]["ap_addr1"].ToString().ToXmlUnicode()
					+ dtAp.Rows[i]["ap_addr2"].ToString().ToXmlUnicode();
				ReplaceBookmark("ap_addr", ap_addr);
				string ap_eddr = dtAp.Rows[i]["ap_eaddr1"].ToString().ToXmlUnicode()
					+ dtAp.Rows[i]["ap_eaddr2"].ToString().ToXmlUnicode()
					+ dtAp.Rows[i]["ap_eaddr3"].ToString().ToXmlUnicode()
					+ dtAp.Rows[i]["ap_eaddr4"].ToString().ToXmlUnicode();
				ReplaceBookmark("ap_eddr", ap_eddr);
				ReplaceBookmark("ap_crep", dtAp.Rows[i]["ap_crep"].ToString().ToXmlUnicode());
				ReplaceBookmark("ap_erep", dtAp.Rows[i]["ap_erep"].ToString().ToXmlUnicode());
			}
		}
		//代理人
		CopyBlock("base_agent");
		using (DataTable dtAgt = GetAgent()) {
			for (int i = 0; i < dtAgt.Rows.Count; i++) {
				CopyBlock("base_apcust");
				ReplaceBookmark("agt_idno1", dtAgt.Rows[i]["agt_idno1"].ToString());
				ReplaceBookmark("agt_id1", dtAgt.Rows[i]["agt_id1"].ToString());
				ReplaceBookmark("base_agt_name1", dtAgt.Rows[i]["agt_name1"].ToString());
				ReplaceBookmark("agt_zip1", dtAgt.Rows[i]["agt_zip"].ToString());
				ReplaceBookmark("agt_addr1", dtAgt.Rows[i]["agt_addr"].ToString());
				ReplaceBookmark("agatt_tel1", dtAgt.Rows[i]["agt_tel"].ToString());
				ReplaceBookmark("agatt_fax1", dtAgt.Rows[i]["agt_fax"].ToString());
				ReplaceBookmark("agt_idno2", dtAgt.Rows[i]["agt_idno2"].ToString());
				ReplaceBookmark("agt_id2", dtAgt.Rows[i]["agt_id2"].ToString());
				ReplaceBookmark("base_agt_name2", dtAgt.Rows[i]["agt_name2"].ToString());
				ReplaceBookmark("agt_zip2", dtAgt.Rows[i]["agt_zip"].ToString());
				ReplaceBookmark("agt_addr2", dtAgt.Rows[i]["agt_addr"].ToString());
				ReplaceBookmark("agatt_tel2", dtAgt.Rows[i]["agt_tel"].ToString());
				ReplaceBookmark("agatt_fax2", dtAgt.Rows[i]["agt_fax"].ToString());
			}
		}
		//發明人/新型創作/設計人
		using (DataTable dtAnt = GetAnt()) {
			for (int i = 0; i < dtAnt.Rows.Count; i++) {
				CopyBlock("base_ant1");
				ReplaceBookmark("base_ant_num", "發明人" + (i + 1).ToString());
				ReplaceBookmark("base_ant_country", dtAnt.Rows[i]["Country_name"].ToString());
				if (dtAnt.Rows[i]["ant_country"].ToString() == "T") {
					CopyBlock("base_ant2");
					ReplaceBookmark("ant_id", dtAnt.Rows[i]["ant_id"].ToString());
				} 
				CopyBlock("base_ant3");
				ReplaceBookmark("base_ant_cname", dtAnt.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
				ReplaceBookmark("base_ant_ename", dtAnt.Rows[i]["Ename_string"].ToString().ToXmlUnicode());
			}
		}

		AppendFoot(1);//頁尾
	}
 * */
	#endregion
}