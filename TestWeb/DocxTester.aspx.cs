using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using MyLibrary;

namespace TestWeb {
	public partial class DocxTester : System.Web.UI.Page {
		protected string in_scode = "";
		protected string in_no = "";
		protected string branch = "";
		protected string receipt_title = "";

		public OpenXmlHelper ipoRpt = null;
		OpenXmlHelperList ipoList = new OpenXmlHelperList();
		protected string templateFile = "";
		protected string baseFile = "";
		protected string outputFile = "";

		protected void Page_Load(object sender, EventArgs e) {
			try {
				//WordOut1();
				ListTest();
			}
			catch (Exception ex) {
				Response.Write(ex.ToString());
				throw ex;
			}
			finally {
				if(ipoRpt!=null)ipoRpt.Dispose();
			}
		}

		protected void ListTest() {
			Dictionary<string, string> _TemplateFileList = new Dictionary<string, string>();
			_TemplateFileList.Add("apply", Server.MapPath("~/ReportTemplate") + @"\01發明專利申請書.docx");
			_TemplateFileList.Add("base", Server.MapPath("~/ReportTemplate") + @"\00基本資料表.docx");
			ipoList.CloneFromFile(_TemplateFileList, true);

			//標題區塊
			ipoList.CopyBlock("b_title");
			//一併申請實體審查
			ipoList.ReplaceBookmark("reality", "是");

			//事務所或申請人案件編號
			ipoList.ReplaceBookmark("seq", "NP28758-n100");
			//中文發明名稱 / 英文發明名稱
			ipoList.ReplaceBookmark("cappl_name", "1112&峯3435");
			ipoList.ReplaceBookmark("eappl_name", "1112&FENG3435");
			//申請人
			ipoList.CopyBlock("apply","b_apply");
			ipoList.ReplaceBookmark("apply_num", "1");
			ipoList.ReplaceBookmark("ap_country", "TW中華民國");
			ipoList.ReplaceBookmark("ap_cname_title", "中文名稱");
			ipoList.ReplaceBookmark("ap_ename_title", "英文名稱");
			ipoList.ReplaceBookmark("ap_cname", "英業達股份有限公司");
			ipoList.ReplaceBookmark("ap_ename", "INVENTEC&LIFE CORPORATION");
			//代理人
			ipoList.ReplaceBookmark("agt_name1", "高,玉駿");
			ipoList.ReplaceBookmark("agt_name2", "楊,祺雄");
			//發明人
			ipoList.CopyBlock("b_ant");
			ipoList.ReplaceBookmark("ant_num", "發明人1");
			ipoList.ReplaceBookmark("ant_country", "AT奧地利");
			ipoList.ReplaceBookmark("ant_cname", "許,塪瑄塪");
			ipoList.ReplaceBookmark("ant_ename", "xu,yix&uang");
			//主張優惠期
			ipoList.CopyBlock("b_exh");
			ipoList.ReplaceBookmark("exh_date", "");

			//主張利用生物材料/生物材料不須寄存/聲明本人就相同創作在申請本發明專利之同日-另申請新型專利/收據抬頭
			ipoList.CopyBlock("b_content");
			//聲明本人就相同創作在申請本發明專利之同日-另申請新型專利
			ipoList.ReplaceBookmark("same_apply", "是");
			ipoList.ReplaceBookmark("receipt_name", "英業達股份有限公司(代繳人：聖島國際專利商標聯合事務所)");

			//附送書件
			ipoList.CloneReplaceBlock("b_attach", "#seq#", "NP28758");
			//具結
			ipoList.CopyBlock("b_sign");

			bool baseflag = true;//是否產生基本資料表
			if (baseflag) {
				ipoList.CopyPageFoot("apply", true);
				AppendBaseDataList("base");
				ipoList.CopyPageFoot("base", false);
			} else {
				ipoList.CopyPageFoot("apply", false);
			}

			ipoList.Flush("NP28758-發明.docx");
		}

		//抓取基本資料表
		protected void AppendBaseDataList(string baseDocName) {
			ipoList.CopyBlock(baseDocName, "base_title");
			//申請人
			ipoList.CopyBlock(baseDocName, "base_apcust1");
			ipoList.ReplaceBookmark("base_ap_num", "1");
			ipoList.ReplaceBookmark("base_ap_country", "TW中華民國");
			ipoList.ReplaceBookmark("ap_class", "法人公司機關學校");
			ipoList.CopyBlock(baseDocName, "base_apcust2");
			ipoList.ReplaceBookmark("apcust_no", "04322046");
			ipoList.CopyBlock(baseDocName, "base_apcust3");
			ipoList.ReplaceBookmark("base_ap_cname_title", "中文名稱");
			ipoList.ReplaceBookmark("base_ap_ename_title", "英文名稱");
			ipoList.ReplaceBookmark("base_ap_cname", "英業達股份有限公司");
			ipoList.ReplaceBookmark("base_ap_ename", "INVENTEC&LIFE CORPORATION");
			ipoList.ReplaceBookmark("ap_live_country", "TW中華民國");
			ipoList.ReplaceBookmark("ap_zip", "840");
			ipoList.ReplaceBookmark("ap_addr", "高雄市大樹區學城路1段9、13、15、17、19、21、23號");
			ipoList.ReplaceBookmark("ap_eddr", "abc");
			ipoList.ReplaceBookmark("ap_crep", "堃峯");
			ipoList.ReplaceBookmark("ap_erep", "Lee &Richard");

			//代理人
			ipoList.CopyBlock(baseDocName, "base_agent");
			ipoList.CopyBlock(baseDocName, "base_apcust");
			ipoList.ReplaceBookmark("agt_idno1", "A0060");
			ipoList.ReplaceBookmark("agt_id1", "B100379440");
			ipoList.ReplaceBookmark("base_agt_name1", "高,玉駿");
			ipoList.ReplaceBookmark("agt_zip1", "105");
			ipoList.ReplaceBookmark("agt_addr1", "臺北市松山區南京東路3段248號7樓");
			ipoList.ReplaceBookmark("agatt_tel1", "02-27751823");
			ipoList.ReplaceBookmark("agatt_fax1", "02-27316377");
			ipoList.ReplaceBookmark("agt_idno2", "02725");
			ipoList.ReplaceBookmark("agt_id2", "M120741174");
			ipoList.ReplaceBookmark("base_agt_name2", "楊,祺雄");
			ipoList.ReplaceBookmark("agt_zip2", "105");
			ipoList.ReplaceBookmark("agt_addr2", "臺北市松山區南京東路3段248號7樓");
			ipoList.ReplaceBookmark("agatt_tel2", "02-27751823");
			ipoList.ReplaceBookmark("agatt_fax2", "02-27316377");

			//發明人/新型創作/設計人
			ipoList.CopyBlock(baseDocName, "base_ant1");
			ipoList.ReplaceBookmark("base_ant_num", "發明人1");
			ipoList.ReplaceBookmark("base_ant_country", "AT奧地利");
			ipoList.CopyBlock(baseDocName, "base_ant3");
			ipoList.ReplaceBookmark("base_ant_cname", "許,塪瑄塪");
			ipoList.ReplaceBookmark("base_ant_ename", "xu,yix&uang");
		}


		protected void WordOut1() {
			templateFile = Server.MapPath("~/ReportTemplate") + @"\01發明專利申請書.docx";
			baseFile = Server.MapPath("~/ReportTemplate") + @"\00基本資料表.docx";

			ipoRpt = new OpenXmlHelper();
			ipoRpt.CloneFromFile(templateFile, baseFile, true);

			//標題區塊
			ipoRpt.CopyBlock("b_title");
			//一併申請實體審查
			ipoRpt.ReplaceBookmark("reality", "是");

			//事務所或申請人案件編號
			ipoRpt.ReplaceBookmark("seq", "NP28758-n100");
			//中文發明名稱 / 英文發明名稱
			ipoRpt.ReplaceBookmark("cappl_name", "1112&峯3435");
			ipoRpt.ReplaceBookmark("eappl_name", "1112&FENG3435");
			//申請人
			ipoRpt.CopyBlock("b_apply");
			ipoRpt.ReplaceBookmark("apply_num", "1");
			ipoRpt.ReplaceBookmark("ap_country", "TW中華民國");
			ipoRpt.ReplaceBookmark("ap_cname_title", "中文名稱");
			ipoRpt.ReplaceBookmark("ap_ename_title", "英文名稱");
			ipoRpt.ReplaceBookmark("ap_cname", "英業達股份有限公司");
			ipoRpt.ReplaceBookmark("ap_ename", "INVENTEC&LIFE CORPORATION");
			//代理人
			ipoRpt.ReplaceBookmark("agt_name1", "高,玉駿");
			ipoRpt.ReplaceBookmark("agt_name2", "楊,祺雄");
			//發明人
			ipoRpt.CopyBlock("b_ant");
			ipoRpt.ReplaceBookmark("ant_num", "發明人1");
			ipoRpt.ReplaceBookmark("ant_country", "AT奧地利");
			ipoRpt.ReplaceBookmark("ant_cname", "許,塪瑄塪");
			ipoRpt.ReplaceBookmark("ant_ename", "xu,yix&uang");
			//主張優惠期
			ipoRpt.CopyBlock("b_exh");
			ipoRpt.ReplaceBookmark("exh_date", "");

			//主張利用生物材料/生物材料不須寄存/聲明本人就相同創作在申請本發明專利之同日-另申請新型專利/收據抬頭
			ipoRpt.CopyBlock("b_content");
			//聲明本人就相同創作在申請本發明專利之同日-另申請新型專利
			ipoRpt.ReplaceBookmark("same_apply", "是");
			ipoRpt.ReplaceBookmark("receipt_name", "英業達股份有限公司(代繳人：聖島國際專利商標聯合事務所)");

			//附送書件
			ipoRpt.CloneReplaceBlock("b_attach", "#seq#", "NP28758");
			//具結
			ipoRpt.CopyBlock("b_sign");

			bool baseflag = true;//是否產生基本資料表
			if (baseflag) {
				ipoRpt.CopyPageFoot(ipoRpt.tempDoc,true);
				AppendBaseData();
				ipoRpt.CopyPageFoot(ipoRpt.baseDoc, false);
			} else {
				ipoRpt.CopyPageFoot(ipoRpt.tempDoc, false);
			}
			ipoRpt.Flush("NP28758-發明.docx");
		}

		//抓取基本資料表
		protected void AppendBaseData() {
			ipoRpt.CopyBlock(ipoRpt.baseDoc, "base_title");
			//申請人
			ipoRpt.CopyBlock(ipoRpt.baseDoc, "base_apcust1");
			ipoRpt.ReplaceBookmark("base_ap_num", "1");
			ipoRpt.ReplaceBookmark("base_ap_country", "TW中華民國");
			ipoRpt.ReplaceBookmark("ap_class", "法人公司機關學校");
			ipoRpt.CopyBlock(ipoRpt.baseDoc, "base_apcust2");
			ipoRpt.ReplaceBookmark("apcust_no", "04322046");
			ipoRpt.CopyBlock(ipoRpt.baseDoc, "base_apcust3");
			ipoRpt.ReplaceBookmark("base_ap_cname_title", "中文名稱");
			ipoRpt.ReplaceBookmark("base_ap_ename_title", "英文名稱");
			ipoRpt.ReplaceBookmark("base_ap_cname", "英業達股份有限公司");
			ipoRpt.ReplaceBookmark("base_ap_ename", "INVENTEC&LIFE CORPORATION");
			ipoRpt.ReplaceBookmark("ap_live_country", "TW中華民國");
			ipoRpt.ReplaceBookmark("ap_zip", "840");
			ipoRpt.ReplaceBookmark("ap_addr", "高雄市大樹區學城路1段9、13、15、17、19、21、23號");
			ipoRpt.ReplaceBookmark("ap_eddr", "abc");
			ipoRpt.ReplaceBookmark("ap_crep", "堃峯");
			ipoRpt.ReplaceBookmark("ap_erep", "Lee &Richard");

			//代理人
			ipoRpt.CopyBlock(ipoRpt.baseDoc, "base_agent");
			ipoRpt.CopyBlock(ipoRpt.baseDoc, "base_apcust");
			ipoRpt.ReplaceBookmark("agt_idno1", "A0060");
			ipoRpt.ReplaceBookmark("agt_id1", "B100379440");
			ipoRpt.ReplaceBookmark("base_agt_name1", "高,玉駿");
			ipoRpt.ReplaceBookmark("agt_zip1", "105");
			ipoRpt.ReplaceBookmark("agt_addr1", "臺北市松山區南京東路3段248號7樓");
			ipoRpt.ReplaceBookmark("agatt_tel1", "02-27751823");
			ipoRpt.ReplaceBookmark("agatt_fax1", "02-27316377");
			ipoRpt.ReplaceBookmark("agt_idno2", "02725");
			ipoRpt.ReplaceBookmark("agt_id2", "M120741174");
			ipoRpt.ReplaceBookmark("base_agt_name2", "楊,祺雄");
			ipoRpt.ReplaceBookmark("agt_zip2", "105");
			ipoRpt.ReplaceBookmark("agt_addr2", "臺北市松山區南京東路3段248號7樓");
			ipoRpt.ReplaceBookmark("agatt_tel2", "02-27751823");
			ipoRpt.ReplaceBookmark("agatt_fax2", "02-27316377");

			//發明人/新型創作/設計人
			ipoRpt.CopyBlock(ipoRpt.baseDoc, "base_ant1");
			ipoRpt.ReplaceBookmark("base_ant_num", "發明人1");
			ipoRpt.ReplaceBookmark("base_ant_country", "AT奧地利");
			ipoRpt.CopyBlock(ipoRpt.baseDoc, "base_ant3");
			ipoRpt.ReplaceBookmark("base_ant_cname", "許,塪瑄塪");
			ipoRpt.ReplaceBookmark("base_ant_ename", "xu,yix&uang");
		}
	}
}