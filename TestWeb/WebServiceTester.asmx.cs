using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;

namespace TestWeb {
	/// <summary>
	///WebServiceTester 的摘要描述
	/// </summary>
	[WebService(Namespace = "http://tempuri.org/")]
	[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
	[System.ComponentModel.ToolboxItem(false)]
	// 若要允許使用 ASP.NET AJAX 從指令碼呼叫此 Web 服務，請取消註解下列一行。
	// [System.Web.Script.Services.ScriptService]
	public class WebServiceTester : System.Web.Services.WebService {

		[WebMethod]
		public int Compute_it(int a,int b) {
			int End_Answer = a * b;
			return End_Answer;
		}
	}
}
