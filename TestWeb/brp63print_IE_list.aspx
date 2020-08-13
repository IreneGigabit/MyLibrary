<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace = "System.IO"%>
<%@ Import Namespace = "System.Linq"%>
<%@ Import Namespace = "System.Collections.Generic"%>

<script runat="server">
	protected string in_scode = "";
	protected string in_no = "";
	protected string branch = "";

	protected IpoReport ipoRpt = null;

	private void Page_Load(System.Object sender, System.EventArgs e) {
		Response.CacheControl = "Private";
		Response.AddHeader("Pragma", "no-cache");
		Response.Expires = -1;
		Response.Clear();

		in_scode = (Request["in_scode"] ?? "n100").ToString();//n100
		in_no = (Request["in_no"] ?? "20170103001").ToString();//20170103001
		branch = (Request["branch"] ?? "N").ToString();//N

		try {
			string dbSession = (Session["btbrtdb"] ?? "Server=web08;Database=sindbs;User ID=web_usr;Password=web1823").ToString();
			ipoRpt = new IpoReport(dbSession, in_scode, in_no, branch);
			//ipoRpt.Create();
			//ipoRpt.Flush(ipoRpt.Seq + "-發明.docx");
			WordOut();
		}
		catch (Exception ex) {
			Response.Write(ex.ToString().Replace("\n", "<BR>"));
		}
		finally {
			if (ipoRpt != null) ipoRpt.Close();
		}
	}

	protected void WordOut() {
		Dictionary<string, string> _TemplateFileList = new Dictionary<string, string>();
		_TemplateFileList.Add("apply", Server.MapPath("~/ReportTemplate") + @"\01發明專利申請書.docx");
		_TemplateFileList.Add("base", Server.MapPath("~/ReportTemplate") + @"\00基本資料表.docx");
		ipoRpt.CloneFromFile(_TemplateFileList, true);
		
		DataTable dmp = ipoRpt.Dmp;
		if (dmp.Rows.Count > 0) {
			//標題區塊
			ipoRpt.CopyBlock("b_title");
			//一併申請實體審查
			if (dmp.Rows[0]["reality"].ToString() == "Y") {
				ipoRpt.ReplaceBookmark("reality", "是");
			} else {
				ipoRpt.ReplaceBookmark("reality", "否");
			}

			//事務所或申請人案件編號
			ipoRpt.ReplaceBookmark("seq", ipoRpt.Seq + "-" + dmp.Rows[0]["scode1"].ToString());
			//中文發明名稱 / 英文發明名稱
			ipoRpt.ReplaceBookmark("cappl_name", dmp.Rows[0]["cappl_name"].ToString().ToXmlUnicode());
			ipoRpt.ReplaceBookmark("eappl_name", dmp.Rows[0]["eappl_name"].ToString().ToXmlUnicode());
			//申請人
			using (DataTable dtAp = ipoRpt.Apcust) {
				for (int i = 0; i < dtAp.Rows.Count; i++) {
					ipoRpt.CopyBlock("b_apply");
					ipoRpt.ReplaceBookmark("apply_num", (i + 1).ToString());
					ipoRpt.ReplaceBookmark("ap_country", dtAp.Rows[i]["Country_name"].ToString());
					ipoRpt.ReplaceBookmark("ap_cname_title", dtAp.Rows[i]["Title_cname"].ToString());
					ipoRpt.ReplaceBookmark("ap_ename_title", dtAp.Rows[i]["Title_ename"].ToString());
					ipoRpt.ReplaceBookmark("ap_cname", dtAp.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
					ipoRpt.ReplaceBookmark("ap_ename", dtAp.Rows[i]["Ename_string"].ToString().ToXmlUnicode());
				}
			}
			//代理人
			ipoRpt.CopyBlock("b_agent");
			using (DataTable dtAgt = ipoRpt.Agent) {
				ipoRpt.ReplaceBookmark("agt_name1", dtAgt.Rows[0]["agt_name1"].ToString().Trim());
				ipoRpt.ReplaceBookmark("agt_name2", dtAgt.Rows[0]["agt_name2"].ToString().Trim());
			}
			//發明人
			using (DataTable dtAnt = ipoRpt.Ant) {
				for (int i = 0; i < dtAnt.Rows.Count; i++) {
					ipoRpt.CopyBlock("b_ant");
					ipoRpt.ReplaceBookmark("ant_num", "發明人" + (i + 1).ToString());
					ipoRpt.ReplaceBookmark("ant_country", dtAnt.Rows[i]["Country_name"].ToString());
					ipoRpt.ReplaceBookmark("ant_cname", dtAnt.Rows[i]["Cname_string"].ToString().ToXmlUnicode());
					ipoRpt.ReplaceBookmark("ant_ename", dtAnt.Rows[i]["Ename_string"].ToString().ToXmlUnicode());
				}
			}
			//主張優惠期
			ipoRpt.CopyBlock("b_exh");
			string exh_date = "";
			if (dmp.Rows[0]["exhibitor"].ToString() == "Y") {//參展或發表日期填入表中的發生日期
				if (dmp.Rows[0]["exh_date"] != System.DBNull.Value && dmp.Rows[0]["exh_date"] != null) {
					exh_date = Convert.ToDateTime(dmp.Rows[0]["exh_date"]).ToString("yyyy/MM/dd");
				}
			}
			ipoRpt.ReplaceBookmark("exh_date", exh_date);

			//主張優先權
			using (DataTable dtPrior = ipoRpt.Prior) {
				for (int i = 0; i < dtPrior.Rows.Count; i++) {
					ipoRpt.CopyBlock("b_prior1");
					ipoRpt.ReplaceBookmark("prior_num", (i + 1).ToString());
					ipoRpt.ReplaceBookmark("prior_country", dtPrior.Rows[i]["Country_name"].ToString());
					string prior_date = "";
					if (dtPrior.Rows[0]["prior_date"] != System.DBNull.Value && dtPrior.Rows[0]["prior_date"] != null) {
						prior_date = Convert.ToDateTime(dtPrior.Rows[0]["prior_date"]).ToString("yyyy/MM/dd");
					}
					ipoRpt.ReplaceBookmark("prior_date", prior_date);
					ipoRpt.ReplaceBookmark("prior_no", dtPrior.Rows[i]["prior_no"].ToString());

					switch (dtPrior.Rows[i]["prior_country"].ToString().Trim()) {
						case "JA":
							ipoRpt.CopyBlock("b_prior2");
							ipoRpt.ReplaceBookmark("case1nm", dtPrior.Rows[i]["case1nm"].ToString());
							ipoRpt.CopyBlock("b_prior3");
							ipoRpt.ReplaceBookmark("mprior_access", dtPrior.Rows[i]["mprior_access"].ToString());
							break;
						case "KO":
							ipoRpt.CopyBlock("b_prior3");
							ipoRpt.ReplaceBookmark("mprior_access", "交換");
							break;
					}
					
					ipoRpt.AddParagraph();
				}
			}
			
			//主張利用生物材料/生物材料不須寄存/聲明本人就相同創作在申請本發明專利之同日-另申請新型專利/收據抬頭
			ipoRpt.CopyBlock("b_content");
			//聲明本人就相同創作在申請本發明專利之同日-另申請新型專利
			if (dmp.Rows[0]["same_apply"].ToString() == "Y") {
				ipoRpt.ReplaceBookmark("same_apply", "是");
			} else {
				ipoRpt.ReplaceBookmark("same_apply", "");
			}
			ipoRpt.ReplaceBookmark("receipt_name", ipoRpt.RectitleName);

			//附送書件
			ipoRpt.CopyReplaceBlock("b_attach", "#seq#", ipoRpt.Seq);
			//具結
			ipoRpt.CopyBlock("b_sign");


			bool baseflag = true;//是否產生基本資料表
			if (baseflag) {
				ipoRpt.CopyPageFoot("apply",true);
				ipoRpt.AppendBaseData("base");
			} else {
				ipoRpt.CopyPageFoot("apply", false);
			}
		}

		ipoRpt.Flush(ipoRpt.Seq + "-發明.docx");
	}
</script>
