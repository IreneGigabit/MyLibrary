<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace = "System.IO"%>
<%@ Import Namespace = "System.Linq"%>
<%@ Import Namespace = "System.Collections.Generic"%>
<%@ Import Namespace = "DocumentFormat.OpenXml"%>
<%@ Import Namespace = "DocumentFormat.OpenXml.Packaging"%>
<%@ Import Namespace = "DocumentFormat.OpenXml.Wordprocessing"%>
<%@ Import Namespace = "A=DocumentFormat.OpenXml.Drawing" %>
<%@ Import Namespace = "DW=DocumentFormat.OpenXml.Drawing.Wordprocessing"%>
<%@ Import Namespace = "PIC=DocumentFormat.OpenXml.Drawing.Pictures"%>

<script runat="server">
	protected string in_scode = "";
	protected string in_no = "";
	protected string branch = "";
	protected string receipt_title = "";

	public IpoReport ipoRpt = null;
	protected string templateFile = "";
	protected string outputFile = "";

	private void Page_Load(System.Object sender, System.EventArgs e) {
		Response.CacheControl = "Private";
		Response.AddHeader("Pragma", "no-cache");
		Response.Expires = -1;
		Response.Clear();

		in_scode = (Request["in_scode"] ?? "").ToString();//n100
		in_no = (Request["in_no"] ?? "").ToString();//20170103001
		branch = (Request["branch"] ?? "").ToString();//N
		receipt_title = (Request["receipt_title"] ?? "").ToString();//B

		try {
			//WordOut();
			WordOutNew();
		}
		catch (Exception ex) {
			//Response.Write(ex.ToString());
			throw ex;
		}
		finally {
			//ipoRpt.CloseRpt();
		}
	}

	protected void WordOutNew() {
		templateFile = Server.MapPath("~/ReportTemplate") + @"\01_發明專利申請書.docx";
		OpenXmlHelper ipoRpt = new OpenXmlHelper();
		Dictionary<string, string> tplDict = new Dictionary<string, string>();
		tplDict.Add("apply", templateFile);

		ipoRpt.CloneFromFile(tplDict, false);
		ipoRpt.Flush("-發明-" + DateTime.Now.ToString("yyyyMMdd") + ".docx");
	}
</script>
