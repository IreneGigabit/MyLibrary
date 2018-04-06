using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TestConsole
{
	class SplitTester
	{
		static void Main(string[] args) {
			method_4();
			//method_5("U", "AC05", "ext", new Dictionary<string, string>() { { "seq", "AAAA" }, { "seq1", "BBBB" }, { "Key3", "CCCC" }, { "Key4", "DDDD" } });
			method_5("U", "AC05", "ext", new Dictionary<string, string>() { { "rec_comp", "AAAA" }, { "rec_no", "BBBB" }, { "rec_no1", "0" } });
			Console.WriteLine("請按任一鍵離開..");
			Console.ReadKey();
		}

		//基本切割
		public static void method_1() {
			string s = "abcdeabcdeabcde";
			string[] sArray = s.Split('c');
			foreach (string i in sArray)
				Console.WriteLine(i.ToString());
			/*
			輸出結果:
			ab
			deab
			deab
			de
			*/
		}

		//字串切割
		public static void method_2() {
			string s = "abcdeabcdeabcde";
			string[] sArray1 = s.Split(new char[3] { 'c', 'd', 'e' });
			foreach (string i in sArray1)
				Console.WriteLine(i.ToString());
			/*
			輸出結果：
			ab
			ab
			ab
			*/
		}

		////字串切割(正規表示)
		public static void method_3() {
			string content = "agcyongfa365macyongfa365gggyongfa365ytx";
			string[] resultString = Regex.Split(content, "yongfa365", RegexOptions.IgnoreCase);
			foreach (string i in resultString)
				Console.WriteLine(i.ToString());
			/*
			輸出結果:
			agc
			mac
			ggg
			ytx
			*/
		}
		public static void method_4() {
			string str1 = "我**是*****一*****個*****教*****師";
			string[] str2 = System.Text.RegularExpressions.Regex.Split(str1, @"\*+");
			foreach (string i in str2)
				Console.WriteLine(i.ToString());
			/*
			輸出結果:
			我
			是
			一
			個
			教
			師
			*/
		}


		public static void method_5(string pUd_flag, string pPrgid, string pTable, Dictionary<string, string> whereKey) {
			string where = "";
			foreach (KeyValuePair<string, string> item in whereKey) {
				where += string.Format("and {0} ='{1}' ", item.Key, item.Value);
			}

			string SQL = "";
			SQL = "insert into " + pTable + "_log(ud_flag,ud_date,ud_scode,prgid)";
			SQL += " select '" + pUd_flag + "',getdate(),'m1583',";
			SQL += "'" + pPrgid + "'";
			SQL += " from " + pTable;
			SQL += " where 1 = 1 ";
			SQL += where;

			Console.WriteLine(SQL);
		}
	}
}
