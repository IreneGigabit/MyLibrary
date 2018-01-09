using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Collections;
using System.Reflection;
using System.Text;

/// <summary>
/// DataExt 的摘要描述
/// ref:https://msdn.microsoft.com/zh-tw/library/cc716729(v=vs.110).aspx
/// </summary>

namespace MyLibrary {
	public static class DataExt {
		static string debugStr = "";

		#region datatable mapping Dictionary
		public static Dictionary<string, object> ToDictionary(this DataTable table) {
			return table.ToDictionary(false);
		}
		
		public static Dictionary<string, object> ToDictionary(this DataTable table, bool debug) {
			Dictionary<string, object> RtnVal = new Dictionary<string, object>();
			//採用Linq轉換
			if (table.Rows.Count > 0) {
				DataRow row0 = table.Select().First(); //Single/SingleOrDefault
				RtnVal = row0.Table.Columns.Cast<DataColumn>().ToDictionary(col => col.ColumnName, col => col.MappingType(row0[col.ColumnName], true));
			} else {
				RtnVal = table.Columns.Cast<DataColumn>().ToDictionary(col => col.ColumnName, col => col.MappingType((object)DBNull.Value, true));
			}
			/*不採用Linq
			if (table.Rows.Count > 0) {
				foreach (DataColumn column in table.Columns) {
					object colValue=column.MappingType(table.Rows[0][column.ColumnName], true);
					RtnVal.Add(column.ColumnName, colValue);
				}
			} else {
				foreach (DataColumn column in table.Columns) {
					object colValue=column.MappingType((object)DBNull.Value, true);
					RtnVal.Add(column.ColumnName, colValue);
				}
			}
			*/

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

			return RtnVal;
		}

		public static ArrayList ToList(this DataTable table) {
			return table.ToList(false);
		}

		public static ArrayList ToList(this DataTable table, bool debug) {
			ArrayList rtnArry = new ArrayList();

			if (table.Rows.Count > 0) {
				foreach (DataRow row in table.Rows) {
					Dictionary<string, object> RtnVal = new Dictionary<string, object>();
					//採用Linq轉換
					RtnVal = row.Table.Columns.Cast<DataColumn>().ToDictionary(col => col.ColumnName, col => col.MappingType(row[col.ColumnName], true));
					/*不採用Linq
					foreach (DataColumn column in table.Columns) {
						object colValue=column.MappingType(row[column.ColumnName], false);
						RtnVal.Add(column.ColumnName, colValue);
					}
					*/
					rtnArry.Add(RtnVal);
				}
			}

			if (debug) {
				debugStr = "";
				debugStr+=String.Format("筆數:{0}<BR>", table.Rows.Count);
				if (table.Rows.Count > 0) {
					debugStr += "<table border=1>";
					debugStr += "<tr>";
					foreach (var entry in (Dictionary<string, object>)rtnArry[0]) {
						debugStr += "<td>" + entry.Key + "(" + (entry.Value ?? "").GetType() + ")</td>";
					}
					debugStr += "</tr>";
					foreach (Dictionary<string, object> item in rtnArry) {
						debugStr += "<tr>";
						foreach (var entry in item) {
							debugStr += "<td>" + entry.Value + "</td>";
						}
						debugStr += "</tr>";
					}
					debugStr += "</table>";
				}
				HttpContext.Current.Response.Write(debugStr);
			}

			return rtnArry;
		}

		private static object MappingType(this DataColumn col, object inVal, bool debugFlag) {
			object RtnVal;
			string mType = "";
			string mValue = "";
			if (col.DataType == System.Type.GetType("System.Int16")) {
				mType = "int16";
				RtnVal = inVal is DBNull ? 0 : Int16.Parse(inVal.ToString());
				mValue = RtnVal.ToString();
			} else if (col.DataType == System.Type.GetType("System.Int32")) {
				mType += "int32";
				RtnVal = inVal is DBNull ? 0 : Int32.Parse(inVal.ToString());
				mValue = RtnVal.ToString();
			} else if (col.DataType == System.Type.GetType("System.Int64")) {
				mType += "int64";
				RtnVal = inVal is DBNull ? 0 : Int64.Parse(inVal.ToString());
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
				RtnVal = inVal is DBNull ? "" : inVal.ToString();
				mValue = (string)RtnVal;
			}
			return RtnVal;
			//if (debugFlag) {
			//	debugStr += String.Format("{0}({1})→{2}={3}<BR>", col.ColumnName, col.DataType.ToString(), mType, mValue);
			//}
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
		#endregion

		#region datatable mapping class list
		/// <summary>
		/// 將DataTable轉成List物件(物件屬性名稱=欄位名稱,直接對應)
		/// </summary>
		/// <example> 
		/// <code>
		/// var Wary2 = dt2.ToList<Document>();
		/// </code>
		/// </example>
		public static IList<T> ToList<T>(this DataTable table) where T : new() {
			//IList<PropertyInfo> properties = typeof(T).GetProperties().ToList();
			IList<PropertyInfo> properties = typeof(T).GetProperties();
			IList<T> result = new List<T>();

			foreach (var row in table.Rows) {
				var item = CreateItemFromRow<T>((DataRow)row, properties);
				result.Add(item);
			}

			return result;
		}

		/// <summary>
		/// 將DataTable轉成List物件(物件屬性名稱<>欄位名稱，用Dictionary型態的Mapping物件)
		/// </summary>
		/// <example> 
		/// <code>
		/// var mappings = new Dictionary<string, string>();
		/// mappings.Add("CompId", "CompId");
		/// mappings.Add("HandleUnit", "HandleUnit");
		/// mappings.Add("No", "No");
		/// var Way3 = dt3.ToList<Document>(mappings);
		/// </code>
		/// </example>
		public static IList<T> ToList<T>(this DataTable table, Dictionary<string, string> mappings) where T : new() {
			//IList<PropertyInfo> properties = typeof(T).GetProperties().ToList();
			IList<PropertyInfo> properties = typeof(T).GetProperties();
			IList<T> result = new List<T>();

			foreach (var row in table.Rows) {
				var item = CreateItemFromRow<T>((DataRow)row, properties, mappings);
				result.Add(item);
			}

			return result;
		}

		private static T CreateItemFromRow<T>(DataRow row, IList<PropertyInfo> properties) where T : new() {
			T item = new T();
			foreach (var property in properties) {
				property.SetValue(item, row[property.Name], null);
			}
			return item;
		}

		private static T CreateItemFromRow<T>(DataRow row, IList<PropertyInfo> properties, Dictionary<string, string> mappings) where T : new() {
			T item = new T();
			foreach (var property in properties) {
				if (mappings.ContainsKey(property.Name))
					property.SetValue(item, row[mappings[property.Name]], null);
			}
			return item;
		}
		#endregion
	}
}