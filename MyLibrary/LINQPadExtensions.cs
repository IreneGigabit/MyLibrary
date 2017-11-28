using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

/// <summary>
/// 使用 LINQ 快速產生相對映 SQL Command 查詢結果的類別程式碼
/// </summary>

namespace MyLibrary {
	public static class LINQPadExtensions {
		private static readonly Dictionary<Type, string> TypeAliases = new Dictionary<Type, string> {
		{ typeof(int), "int" },
		{ typeof(short), "short" },
		{ typeof(byte), "byte" },
		{ typeof(byte[]), "byte[]" },
		{ typeof(long), "long" },
		{ typeof(double), "double" },
		{ typeof(decimal), "decimal" },
		{ typeof(float), "float" },
		{ typeof(bool), "bool" },
		{ typeof(string), "string" }
	};

		private static readonly HashSet<Type> NullableTypes = new HashSet<Type> {
		typeof(int),
		typeof(short),
		typeof(long),
		typeof(double),
		typeof(decimal),
		typeof(float),
		typeof(bool),
		typeof(DateTime)
	};

		public static string DumpClass(this SqlConnection connection, string sql, string className = "Info") {
			if (connection.State != ConnectionState.Open) {
				connection.Open();
			}

			var cmd = connection.CreateCommand();
			cmd.CommandText = sql;
			var reader = cmd.ExecuteReader();

			var builder = new StringBuilder();
			do {
				if (reader.FieldCount <= 1) continue;

				builder.AppendFormat("public class {0}{1}", className, "<BR>");
				builder.AppendLine("{<BR>");
				var schema = reader.GetSchemaTable();

				foreach (DataRow row in schema.Rows) {
					var type = (Type)row["DataType"];
					var name = TypeAliases.ContainsKey(type) ? TypeAliases[type] : type.Name;
					var isNullable = (bool)row["AllowDBNull"] && NullableTypes.Contains(type);
					var collumnName = (string)row["ColumnName"];

					builder.AppendLine(string.Format("\tpublic {0}{1} {2} {{ get; set; }}<BR>", name, isNullable ? "?" : string.Empty, collumnName));
					//builder.AppendLine();
				}

				builder.AppendLine("}");
				builder.AppendLine();
			} while (reader.NextResult());

			return builder.ToString();
		}
	}
}
