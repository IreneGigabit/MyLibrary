using System;
using System.Collections.Generic;
using System.Web;
using System.Reflection;
using System.Collections.Specialized;
using System.Data;
using System.Data.Common;

public static class DataMapping {
	#region ToDictionary - Datatable轉Dictionary
	/// <summary>
	/// Datatable轉Dictionary
	/// </summary>
	/// <param name="table">DataTable</param>
	/// <returns></returns>
	public static Dictionary<string, object> ToDictionary(DataTable table) {
		Dictionary<string, object> RtnVal = new Dictionary<string, object>();

		if (table.Rows.Count > 0) {
			foreach (DataRow row in table.Rows) {
				foreach (DataColumn column in table.Columns) {
					if (column.DataType == System.Type.GetType("System.Int32") || column.DataType == System.Type.GetType("System.Decimal") || column.DataType == System.Type.GetType("System.Int16")) {
						if (row[column.ColumnName] is DBNull) {
							RtnVal.Add(column.ColumnName, 0);
						} else {
							RtnVal.Add(column.ColumnName, row[column.ColumnName]);
						}
					} else if (column.DataType == System.Type.GetType("System.DateTime")) {
						if (row[column.ColumnName] is DBNull) {
							RtnVal.Add(column.ColumnName, "");
						} else {
							RtnVal.Add(column.ColumnName, Convert.ToDateTime(row[column.ColumnName]).ToString("yyyy/MM/dd"));
						}
					} else {
						RtnVal.Add(column.ColumnName, row[column.ColumnName].ToString().Trim());
						//Response.Write(column.ColumnName + "=" + SrvrVal[column.ColumnName] + "<BR>");
					}

				}
			}
		} else {
			foreach (DataColumn column in table.Columns) {
				if (column.DataType == System.Type.GetType("System.Int32") || column.DataType == System.Type.GetType("System.Decimal") || column.DataType == System.Type.GetType("System.Int16")) {
					RtnVal.Add(column.ColumnName, 0);
				} else {
					RtnVal.Add(column.ColumnName, "");
				}
			}
		}

		return RtnVal;
	}
	#endregion

	#region ToList-Datatable Mapping class
	/// <summary>
	/// Datatable轉Class
	/// </summary>
	/// <param name="table">DataTable</param>
	/// <returns></returns>
	public static IList<T> ToList<T>(this DataTable table) where T : new() {
		IList<T> result = new List<T>();

		//取得DataTable所有的row data
		foreach (var row in table.Rows) {
			T item = new T();
			MappingItem((DataRow)row, item);
			result.Add(item);
		}

		return result;
	}
	#endregion

	#region MappingItem - DataRow轉Class
	/// <summary>
	/// DataRow轉Class
	/// </summary>
	/// <param name="row">DataRow</param>
	/// <param name="obj">Instance Class</param>
	/// <returns></returns>
	public static void MappingItem(DataRow row, Object obj) {

		IList<PropertyInfo> properties = obj.GetType().GetProperties();

		foreach (var property in properties) {
			if (row.Table.Columns.Contains(property.Name)) {
				SetPropertyValue(obj, property.Name, row[property.Name] is DBNull ? null : row[property.Name].ToString());
			}
		}
	}
	#endregion

	#region MappingItem - DbDataReader轉Class
	/// <summary>
	/// DbDataReader轉Class
	/// </summary>
	/// <param name="row">DataRow</param>
	/// <param name="obj">Instance Class</param>
	/// <returns></returns>
	public static void MappingItem(DbDataReader row, Object obj) {

		IList<PropertyInfo> properties = obj.GetType().GetProperties();

		foreach (var property in properties) {
			if (HasColumn(row, property.Name)) {
				SetPropertyValue(obj, property.Name, row[property.Name] is DBNull ? null : row[property.Name].ToString());
			}
		}
	}
	#endregion

	#region HasColumn
	private static bool HasColumn(this IDataRecord dr, string columnName) {
		for (int i = 0; i < dr.FieldCount; i++) {
			if (dr.GetName(i).Equals(columnName, StringComparison.InvariantCultureIgnoreCase))
				return true;
		}
		return false;
	}
	#endregion

	#region GetProperty
	private static PropertyInfo GetProperty(object instance, string propName) {
		try {
			IList<PropertyInfo> infos = instance.GetType().GetProperties();
			foreach (PropertyInfo info in infos) {
				if (propName.ToLower().Equals(info.Name.ToLower())) {
					return info;
				}
			}
		}
		catch (Exception ex) {
			return null;
			throw ex;
		}
		return null;
	}
	#endregion

	#region SetPropertyValue
	private static bool SetPropertyValue(object instance, string propertyName, string val) {
		if (null == instance) return false;

		PropertyInfo property = GetProperty(instance, propertyName);
		if (null == property) return false;

		if (property.PropertyType == typeof(Nullable<DateTime>)) {
			DateTime dt = new DateTime();
			if (DateTime.TryParse(val, out dt)) {
				property.SetValue(instance, dt, null);
			} else {
				property.SetValue(instance, null, null);
			}
		} else if (property.PropertyType == typeof(decimal)) {
			decimal value = new decimal();
			decimal.TryParse(val, out value);
			property.SetValue(instance, value, null);
		} else if (property.PropertyType == typeof(double)) {
			double value = new double();
			double.TryParse(val, out value);
			property.SetValue(instance, value, null);
		} else if (property.PropertyType == typeof(int)) {
			int value = new int();
			int.TryParse(val, out value);
			property.SetValue(instance, value, null);
		} else if (property.PropertyType == typeof(Boolean)) {
			bool value = false;
			value = val.ToLower().StartsWith("true");
			property.SetValue(instance, value, null);
		} else {
			property.SetValue(instance, val, null);
		}
		return true;
	}
	#endregion
}