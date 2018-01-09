using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using MyLibrary;

namespace TestWeb {
	public partial class EncodeTester : System.Web.UI.Page {
		protected string gov_no = "";
		protected void Page_Load(object sender, EventArgs e) {
			string SQL = "";

			if (IsPostBack) {
				SQL = "UPDATE tarmain ";
				SQL += "  SET gov_date='" + Request["gov_date"] + "' ";
				SQL += "     ,gov_code='" + Request["gov_code"] + "' ";
				SQL += "     ,gov_no='" + Request["gov_no"].toNCR() + "' ";
				SQL += "     ,tran_reason='" + Request["tran_reason"] + "' ";
				SQL += "     ,gtran_scode='" + Session["UserID"] + "' ";
				SQL += "     ,gtran_date=getdate() ";
				SQL += "WHERE tmain_sqlno=" + Request["tmain_sqlno"];
				Response.Write(SQL);
			} else {
				DataTable dt = new DataTable();
				using (SqlConnection cnQry = new SqlConnection("Server=web08;Database=account;User ID=web_usr;Password=web1823")) {
					cnQry.Open();
					SQL = "Select * from sindbs.dbo.dmt where seq='20623' ";
					SqlCommand cmdQry = new SqlCommand(SQL, cnQry);
					new SqlDataAdapter(SQL, cnQry).Fill(dt);

					Dictionary<string, object> RtnVal = dt.ToDictionary(false);
					gov_no = (string)RtnVal["appl_name"];
				}
				Response.Write("source:三嗪堃系前驅物、製備三嗪系前驅物的方法<BR>");
				Response.Write("toNCR:" + "三嗪堃系前驅物、製備三嗪系前驅物的方法<BR>".toNCR());
				Response.Write("fromNCR:" + "三&#21994;&#22531;系前驅物、製備三&#22531;系前驅物的方法<BR>".fromNCR());
			}
		}
	}
}