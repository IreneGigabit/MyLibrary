using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// server_code 的摘要描述
/// </summary>
public static class server_code {
	public static string Left(this string str, int ln) {
		string sret = str.Substring(0, Math.Min(ln, str.Length));
		return sret;
	}

	public static string Right(this string str, int ln) {
		//string sret = s.Substring(str.Length - ln, ln);
		//return sret;
		ln = Math.Max(ln, 0);
		if (str.Length > ln) {
			return str.Substring(str.Length - ln, ln);
		} else {
			return str;
		}
	}

	public static string ToXmlUnicode(this string str) {
		str = HttpUtility.HtmlDecode(str);
		foreach (System.Text.RegularExpressions.Match m
			in System.Text.RegularExpressions.Regex.Matches(str, "&#(?<ncr>\\d+?);"))
			str = str.Replace(m.Value, Convert.ToChar(int.Parse(m.Groups["ncr"].Value)).ToString());
		//str = str.Replace("&", "&amp;");
		//str = str.Replace("<", "&lt;");
		str = str.Replace("＆", "&");//防止英文欄位只能半型
		//ret=str.Replace(">","&gt;");
		//ret=str.Replace("'","&apos;");
		//ret=str.Replace("""","&quot;");

		return str.Trim();
	}

}