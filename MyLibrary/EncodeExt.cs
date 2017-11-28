using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MyLibrary {
	public static class EncodeExt {
		/// <summary>
		/// 判斷字元是否為中文BIG5編碼所不支援的難字，若是，則轉成&amp;#nnnn;這種NCR格式
		/// </summary>
        /// <example> 
        /// <code>
		/// "堃峯".toNCR()
        /// </code>
        /// </example>
		public static string toNCR(this string s) {
			StringBuilder sb = new StringBuilder();
			Encoding big5 = Encoding.GetEncoding("big5");
			foreach (char c in s) {
				//強迫轉碼成Big5，看會不會變成問號
				string cInBig5 = big5.GetString(big5.GetBytes(new char[] { c }));
				//原來不是問號，但轉碼後變問號，判定為難字
				if (c != '?' && cInBig5 == "?")
					sb.AppendFormat("&#{0};", Convert.ToInt32(c));
				else
					sb.Append(c);
			}
			return sb.ToString();
		}

		/// <summary>
		/// 將字串內有NCR格式字元(&amp;#nnnn;)轉成char字元
		/// </summary>
		/// <example> 
		/// <code>
		/// "三&#21994;&#22531;系".fromNCR()
		/// </code>
		/// </example>
		public static string fromNCR(this string s) {
			foreach (System.Text.RegularExpressions.Match m
				in System.Text.RegularExpressions.Regex.Matches(s, "&#(?<ncr>\\d+?);"))
				s = s.Replace(m.Value, Convert.ToChar(int.Parse(m.Groups["ncr"].Value)).ToString());
			return s;
		}
	}
}
