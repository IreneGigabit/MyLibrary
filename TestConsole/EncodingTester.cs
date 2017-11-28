using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Diagnostics;

namespace TestConsole {
	class EncodingTester {
		private static string CurrDir = System.Environment.CurrentDirectory;
		static string outputFile = CurrDir + @"\testfile.txt";

		static void Main(string[] args) {
			DataTable dt = new DataTable();
			using (SqlConnection cnQry = new SqlConnection("Server=web08;Database=account;User ID=web_usr;Password=web1823")) {
				cnQry.Open();
				string SQL = "select top 10 * from dmp where cappl_name like '%&#%' ";
				SqlCommand cmdQry = new SqlCommand(SQL, cnQry);
				new SqlDataAdapter(SQL, cnQry).Fill(dt);

				for (int x = 0; x < dt.Rows.Count; x++) {
					Console.WriteLine("source:" + dt.Rows[x]["cappl_name"]);
					Console.WriteLine("fromNCR:" + fromNCR(dt.Rows[x]["cappl_name"].ToString()));
					using (System.IO.StreamWriter file = new System.IO.StreamWriter(outputFile, x == 0 ? false : true, Encoding.UTF8)) {
						file.WriteLine("fromNCR:" + fromNCR(dt.Rows[x]["cappl_name"].ToString()));
					}
				}
			}
			
			Console.WriteLine("==================================================");
			Console.WriteLine("source:三嗪堃系前驅物、製備三嗪系前驅物的方法");
			Console.WriteLine("toNCR:" + toNCR("三嗪堃系前驅物、製備三嗪系前驅物的方法"));
			Console.WriteLine("fromNCR:" + fromNCR("三&#21994;&#22531;系前驅物、製備三&#22531;系前驅物的方法"));
			Console.WriteLine("fromNCR:" + fromNCR("三&#21994;&#22531;系前驅物、製備三&#22531;系前驅物的方法"));

			string dst = "fromNCR:" + fromNCR("三&#21994;&#22531;系前驅物、製備三&#22531;系前驅物的方法");
			//File.WriteAllText(outputFile, dst, Encoding.UTF8);
			using (System.IO.StreamWriter file = new System.IO.StreamWriter(outputFile, true, Encoding.UTF8)) {
				file.WriteLine("fromNCR:" + fromNCR("三&#21994;&#22531;系前驅物、製備三&#22531;系前驅物的方法"));
			}


			Process.Start(outputFile);
			//Console.WriteLine("請按任一鍵離開..");
			//Console.ReadKey();
		}

		//判斷字元是否為中文BIG5編碼所不支援的難字，若是，則轉成&#29319;這種NCR格式
		private static string toNCR(string rawString) {
			StringBuilder sb = new StringBuilder();
			Encoding big5 = Encoding.GetEncoding("big5");
			foreach (char c in rawString) {
				//強迫轉碼成Big5，看會不會變成問號
				string cInBig5 = big5.GetString(big5.GetBytes(new char[] { c }));
				//原來不是問號，轉碼後變問號，判定為難字
				if (c != '?' && cInBig5 == "?")
					sb.AppendFormat("&#{0};", Convert.ToInt32(c));
				else
					sb.Append(c);
			}
			return sb.ToString();
		}

		//&#nnnn;轉成char
		private static string fromNCR(string s) {
			foreach (System.Text.RegularExpressions.Match m
				in System.Text.RegularExpressions.Regex.Matches(s, "&#(?<ncr>\\d+?);"))
				s = s.Replace(m.Value,Convert.ToChar(int.Parse(m.Groups["ncr"].Value)).ToString());
			return s;
		}
	}
}
