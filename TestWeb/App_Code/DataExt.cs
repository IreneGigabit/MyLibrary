using System;
using System.Collections.Generic;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using System.Text;

/// <summary>
/// DataExt 的摘要描述
/// ref:https://msdn.microsoft.com/zh-tw/library/cc716729(v=vs.110).aspx
/// </summary>

namespace System.Runtime.CompilerServices
{
	public class ExtensionAttribute : Attribute { }
}


public static class DataExt {
	static string debugStr = "";

	#region DataTable 擴展
	public static void ToDictionary(this DataTable table, Dictionary<string, string> RtnVal) {
		table.ToDictionary(RtnVal, false);
	}

	public static void ToDictionary(this DataTable table, Dictionary<string, string> RtnVal, bool debug) {
		if (table.Rows.Count > 0) {
			foreach (DataColumn column in table.Columns) {
				string colValue;
				column.MappingType(table.Rows[0][column.ColumnName], true, out colValue);
				try {
					RtnVal.Add(column.ColumnName.ToLower(), colValue);
				}
				catch (Exception ex) {
					throw new Exception("[" + column.ColumnName + "]", ex);
				}
			}
		} else {
			foreach (DataColumn column in table.Columns) {
				string colValue;
				column.MappingType((object)DBNull.Value, true, out colValue);
				RtnVal.Add(column.ColumnName.ToLower(), colValue);
			}
		}

		if (debug) {
			debugStr = "";
			debugStr += String.Format("筆數:{0}<BR>", table.Rows.Count);
			debugStr += "<table border=1>";
			debugStr += "<tr>";
			foreach (var entry in RtnVal) {
				debugStr += "<td>" + entry.Key + "(" + (entry.Value ?? "").GetType() + ")</td>";
			}
			debugStr += "</tr>";
			debugStr += "<tr>";
			foreach (var entry in RtnVal) {
				debugStr += "<td>" + entry.Value + "</td>";
			}
			debugStr += "</tr>";
			debugStr += "</table>";
			HttpContext.Current.Response.Write(debugStr);

		}
	}

	public static string ToHexString(this byte[] hex) {
		if (hex == null) return null;
		if (hex.Length == 0) return string.Empty;

		var s = new StringBuilder();
		foreach (byte b in hex) {
			s.Append(b.ToString("x2"));
		}
		return s.ToString();
	}

	private static void MappingType(this DataColumn col, object inVal, bool debugFlag, out object RtnVal) {
		string mType = "";
		string mValue = "";
		if (col.DataType == System.Type.GetType("System.Int16")) {
			mType = "int16";
			if (inVal is DBNull) {
				RtnVal = "";
			} else {
				RtnVal = Int16.Parse(inVal.ToString());
			}
			mValue = RtnVal.ToString();
		} else if (col.DataType == System.Type.GetType("System.Int32")) {
			mType += "int32";
			if (inVal is DBNull) {
				RtnVal = "";
			} else {
				RtnVal = Int32.Parse(inVal.ToString());
			}
			mValue = RtnVal.ToString();
		} else if (col.DataType == System.Type.GetType("System.Int64")) {
			mType += "int64";
			if (inVal is DBNull) {
				RtnVal = "";
			} else {
				RtnVal = Int64.Parse(inVal.ToString());
			}
			mValue = RtnVal.ToString();
		} else if (col.DataType == System.Type.GetType("System.Double")) {
			mType += "float";
			RtnVal = inVal is DBNull ? 0 : float.Parse(inVal.ToString());
			mValue = RtnVal.ToString();
		} else if (col.DataType == System.Type.GetType("System.Decimal")) {
			mType += "decimal";
			RtnVal = inVal is DBNull ? 0 : decimal.Parse(inVal.ToString());
			mValue = RtnVal.ToString();
		} else if (col.DataType == System.Type.GetType("System.DateTime")) {
			mType += "datetime";
			DateTime dt = new DateTime();
			if (DateTime.TryParse(inVal.ToString(), out dt)) {
				RtnVal = dt.ToString("yyyy/MM/dd HH:mm:ss").Replace(" 00:00:00", "");
			} else {
				RtnVal = "";
			}
			mValue = (string)RtnVal;
		} else if (col.DataType == System.Type.GetType("System.Byte")) {
			mType += "byte";
			RtnVal = inVal is DBNull ? new Byte() : Byte.Parse(inVal.ToString());
			mValue = ((byte)RtnVal).ToString("x2");
		} else if (col.DataType == System.Type.GetType("System.Byte[]")) {
			mType += "byte[]";
			RtnVal = inVal is DBNull ? (Byte[])null : (Byte[])inVal;
			mValue = ((Byte[])RtnVal).ToHexString();
		} else if (col.DataType == System.Type.GetType("System.Boolean")) {
			mType += "bool";
			RtnVal = inVal is DBNull ? (bool)false : inVal.ToString().ToLower().StartsWith("true");
			mValue = RtnVal.ToString();
		} else {
			mType += "string";
			RtnVal = inVal is DBNull ? "" : inVal.ToString().Trim();
			mValue = (string)RtnVal;
		}
		//if (debugFlag) {
		//	debugStr += String.Format("{0}({1})→{2}={3}<BR>", col.ColumnName, col.DataType.ToString(), mType, mValue);
		//	HttpContext.Current.Response.Write(debugStr);
		//}
	}

	private static void MappingType(this DataColumn col, object inVal, bool debugFlag, out string RtnVal) {
		string mType = "";
		if (col.DataType == System.Type.GetType("System.Int16")) {
			mType = "int16";
			RtnVal = inVal is DBNull ? "0" : inVal.ToString();
		} else if (col.DataType == System.Type.GetType("System.Int32")) {
			mType += "int32";
			RtnVal = inVal is DBNull ? "0" : inVal.ToString();
		} else if (col.DataType == System.Type.GetType("System.Int64")) {
			mType += "int64";
			RtnVal = inVal is DBNull ? "0" : inVal.ToString();
		} else if (col.DataType == System.Type.GetType("System.Double")) {
			mType += "float";
			RtnVal = inVal is DBNull ? "0" : inVal.ToString();
		} else if (col.DataType == System.Type.GetType("System.Decimal")) {
			mType += "decimal";
			RtnVal = inVal is DBNull ? "0" : inVal.ToString();
		} else if (col.DataType == System.Type.GetType("System.DateTime")) {
			mType += "datetime";
			DateTime dt = new DateTime();
			if (DateTime.TryParse(inVal.ToString(), out dt)) {
				RtnVal = dt.ToString("yyyy/MM/dd HH:mm:ss").Replace(" 00:00:00", "");
			} else {
				RtnVal = "";
			}
		} else if (col.DataType == System.Type.GetType("System.Byte")) {
			mType += "byte";
			RtnVal = inVal is DBNull ? "" : inVal.ToString();
		} else if (col.DataType == System.Type.GetType("System.Byte[]")) {
			mType += "byte[]";
			RtnVal = inVal is DBNull ? "" : ((Byte[])inVal).ToHexString();
		} else {
			mType += "string";
			RtnVal = inVal is DBNull ? "" : inVal.ToString().Trim();
		}
		//if (debugFlag) {
		//	debugStr += String.Format("{0}({1})→{2}<BR>", col.ColumnName, col.DataType.ToString(), mType);
		//	HttpContext.Current.Response.Write(debugStr);
		//}
	}
	#endregion

	#region DataReader 擴展
	/// <summary>
	/// 獲取指定型別
	/// </summary> 
	/// <param name="fieldName">欄位名稱</param>
	/// <param name="defaultValue">若是null時回傳預設值</param>
	/// <returns></returns>
	public static T SafeRead<T>(this IDataReader reader, string fieldName, T defaultValue) {
		try {
			object obj = reader[fieldName];
			if (obj == null || obj == System.DBNull.Value)
				return defaultValue;

			return (T)Convert.ChangeType(obj, defaultValue.GetType());
		}
		catch {
			return defaultValue;
		}
	}

	/// <summary>
	/// 獲取字串
	/// </summary> 
	/// <param name="colName">欄位名稱</param>  
	/// <returns></returns>  
	public static string GetString(this IDataReader dr, string colName) {
		if (dr[colName] != DBNull.Value && dr[colName] != null)
			return dr[colName].ToString().Trim();
		return String.Empty;
	}
	/// <summary>
	/// 獲取DateTime(null時回傳現在時間)
	/// </summary>
	/// <param name="colName">欄位名稱</param>  
	/// <returns></returns>
	public static DateTime GetDateTime(this IDataReader dr, string colName) {
		DateTime result = DateTime.Now;
		if (dr[colName] != DBNull.Value && dr[colName] != null) {
			if (!DateTime.TryParse(dr[colName].ToString(), out result))
				throw new Exception("日期格式數據轉換失敗(" + colName + ")");
		}
		return result;
	}
	/// <summary>
	/// 獲取DateTime(可回傳null)
	/// </summary>
	/// <param name="colName">欄位名稱</param>  
	/// <returns></returns>
	public static DateTime? GetNullDateTime(this IDataReader dr, string colName) {

		DateTime? result = null;
		DateTime time = DateTime.Now;
		if (dr[colName] != DBNull.Value && dr[colName] != null) {
			if (!DateTime.TryParse(dr[colName].ToString(), out time))
				throw new Exception("日期格式數據轉換失敗(" + colName + ")");
			result = time;
		}
		return result;
	}

	/// <summary>
	/// 獲取Int16
	/// </summary>
	/// <param name="colName">欄位名稱</param>  
	/// <returns></returns>
	public static Int16 GetInt16(this IDataReader dr, string colName) {
		short result = 0;
		if (dr[colName] != DBNull.Value && dr[colName] != null) {
			if (!short.TryParse(dr[colName].ToString(), out result))
				throw new Exception("短整形轉換失敗(" + colName + ")");
		}
		return result;
	}

	/// <summary>
	/// 獲取Int32
	/// </summary>
	/// <param name="colName">欄位名稱</param>  
	/// <returns></returns>
	public static int GetInt32(this IDataReader dr, string colName) {
		int result = 0;

		if (dr[colName] != DBNull.Value && dr[colName] != null) {
			if (!int.TryParse(dr[colName].ToString(), out result))
				throw new Exception("整形轉換失敗(" + colName + ")");
		}
		return result;
	}

	/// <summary>
	/// 獲取Double
	/// </summary>
	/// <param name="colName">欄位名稱</param>  
	/// <returns></returns>
	public static double GetDouble(this IDataReader dr, string colName) {
		double result = 0.00;
		if (dr[colName] != DBNull.Value && dr[colName] != null) {
			if (!double.TryParse(dr[colName].ToString(), out result))
				throw new Exception("雙精度類型轉換失敗(" + colName + ")");
		}
		return result;
	}
	/// <summary>
	/// 獲取Single
	/// </summary>
	/// <param name="colName">欄位名稱</param>  
	/// <returns></returns>
	public static float GetSingle(this IDataReader dr, string colName) {
		float result = 0.00f;
		if (dr[colName] != DBNull.Value && dr[colName] != null) {
			if (!float.TryParse(dr[colName].ToString(), out result))
				throw new Exception("單精度類型轉換失敗(" + colName + ")");
		}

		return result;
	}

	/// <summary>
	/// 獲取Decimal
	/// </summary>
	/// <param name="colName">欄位名稱</param>  
	/// <returns></returns>
	public static decimal GetDecimal(this IDataReader dr, string colName) {
		decimal result = 0.00m;
		if (dr[colName] != DBNull.Value && dr[colName] != null) {
			if (!decimal.TryParse(dr[colName].ToString(), out result))
				throw new Exception("Decimal類型轉換失敗(" + colName + ")");
		}
		return result;
	}

	/// <summary>
	/// 獲取Byte
	/// </summary> 
	/// <param name="colName">欄位名稱</param>  
	/// <returns></returns>
	public static byte GetByte(this IDataReader dr, string colName) {
		byte result = 0;
		if (dr[colName] != DBNull.Value && dr[colName] != null) {
			if (!byte.TryParse(dr[colName].ToString(), out result))
				throw new Exception("Byte類型轉換失敗(" + colName + ")");
		}
		return result;
	}

	/// <summary>
	/// 獲取bool(如果是1或Y時回傳true);
	/// </summary>
	/// <param name="colName">欄位名稱</param>  
	/// <returns></returns>
	public static bool GetBool(this IDataReader dr, string colName) {
		if (dr[colName] != DBNull.Value && dr[colName] != null) {
			return dr[colName].ToString() == "1" || dr[colName].ToString() == "Y" || dr[colName].ToString().ToLower() == "true";
		}
		return false;
	}
	#endregion
}