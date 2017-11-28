using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using MyLibrary;

namespace TestWeb {
	public partial class DataExtTester : System.Web.UI.Page {
		protected void Page_Load(object sender, EventArgs e) {
			if (!IsPostBack) {
				string SQL = "";
				SQL = "Select top 100 * from sindbs.dbo.dmt ";
				//SQL = "Select top 100 * from recitem_log ";

				DataTable dt = new DataTable();
				using (SqlConnection cnQry = new SqlConnection("Server=web08;Database=account;User ID=web_usr;Password=web1823")) {
					cnQry.Open();
					SqlCommand cmdQry = new SqlCommand(SQL, cnQry);
					new SqlDataAdapter(SQL, cnQry).Fill(dt);

					Dictionary<string, object> RtnVal = dt.ToDictionary(true);

					Response.Write(RtnVal["s_mark"].GetType() + "<BR>");
					if ((string)RtnVal["s_mark"] == "S")
						Response.Write(true);
					else
						Response.Write(false);

					ArrayList lists = dt.ToList(true);

				}

			}
		}

	}
}