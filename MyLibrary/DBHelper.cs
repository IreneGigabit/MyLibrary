using System;
using System.Web;
using System.Data;
using System.Data.SqlClient;

namespace MyLibrary
{
	/// <summary>
	/// 資料庫操作類別
	/// </summary>
	public class DBHelper : IDisposable
	{
		private SqlConnection _conn = null;
		private SqlTransaction _tran = null;
		private SqlCommand _cmd = null;
		public string ConnString { get; set; }
		private bool _debug = false;
		private bool _isTran = true;

		public DBHelper(string connectionString) : this(connectionString, true) { }

		public DBHelper(string connectionString, bool isTransaction) {
			//this._debug = showDebugStr;
			this.ConnString = connectionString;
			this._isTran = isTransaction;

			this._conn = new SqlConnection(this.ConnString);
			_conn.Open();

			if (this._isTran) {
				this._tran = _conn.BeginTransaction();
				this._cmd = new SqlCommand("", _conn, _tran);
			} else {
				this._cmd = new SqlCommand("", _conn);
			}
		}

		public DBHelper Debug(bool showDebugStr) {
			this._debug = showDebugStr;
			return this;
		}

		public void Dispose() {
			this._conn.Close(); this._conn.Dispose();
			this._cmd.Dispose();
			if (this._tran != null) this._tran.Dispose();

			GC.SuppressFinalize(this);
		}

		public void Commit() {
			if (this._tran != null) _tran.Commit();
		}

		public void RollBack() {
			if (this._tran != null) _tran.Rollback();
		}

		/// <summary>
		/// 執行查詢，取得SqlDataReader；SqlDataReader使用後須Close，否則會Lock(強烈建議使用using)。
		/// </summary>
		public SqlDataReader ExecuteReader(string commandText) {
			if (this._debug) {
				HttpContext.Current.Response.Write(commandText + "<HR>");
			}
			this._cmd.CommandText = commandText;
			SqlDataReader dr = this._cmd.ExecuteReader();

			return dr;
		}

		/// <summary>
		/// 執行T-SQL，並傳回受影響的資料筆數。
		/// </summary>
		public int ExecuteNonQuery(string commandText) {
			if (this._debug) {
				HttpContext.Current.Response.Write(commandText + "<HR>");
			}
			this._cmd.CommandText = commandText;
			return this._cmd.ExecuteNonQuery();
		}

		/// <summary>
		/// 執行查詢，取得第一行第一欄資料，會忽略其他的資料行或資料列。
		/// </summary>
		public object ExecuteScalar(string commandText) {
			if (this._debug) {
				HttpContext.Current.Response.Write(commandText + "<HR>");
			}
			this._cmd.CommandText = commandText;
			return this._cmd.ExecuteScalar();
		}

		/// <summary>
		/// 執行查詢，並傳回DataTable。
		/// </summary>
		public void DataTable(string commandText, DataTable dt) {
			if (this._debug) {
				HttpContext.Current.Response.Write(commandText + "<HR>");
			}
			using (SqlDataAdapter adapter = new SqlDataAdapter(commandText, this._conn)) {
				if (this._isTran) {
					adapter.SelectCommand.Transaction = this._tran;
				}
				adapter.Fill(dt);
			}
		}

		/// <summary>
		/// 執行查詢，並傳回DataSet。
		/// </summary>
		public void DataSet(string commandText, DataSet ds) {
			if (this._debug) {
				HttpContext.Current.Response.Write(commandText + "<HR>");
			}
			using (SqlDataAdapter adapter = new SqlDataAdapter(commandText, this._conn)) {
				if (this._isTran) {
					adapter.SelectCommand.Transaction = this._tran;
				}
				//DataSet ds = new DataSet();
				adapter.Fill(ds);
			}
		}
	}
}
