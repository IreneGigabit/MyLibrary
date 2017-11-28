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
	public partial class LINQPadExtTester : System.Web.UI.Page {
		protected void Page_Load(object sender, EventArgs e) {
			using (SqlConnection cnQry = new SqlConnection("Server=web08;Database=account;User ID=web_usr;Password=web1823")) {
				cnQry.Open();
				// 這邊修改為您要執行的 SQL Command
				var sqlCommand = @"Select top 10 * from sindbs.dbo.dmt";
				// 在 DumpClass 方法裡放 SQL Command 和 Class 名稱
				Response.Write(cnQry.DumpClass(sqlCommand.ToString(), "ClassName"));
			}
		}
	}
}