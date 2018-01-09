using System;
using System.Collections.Generic;
using System.Web;
using System.IO;
using System.Data.SqlClient;

/// <summary>
/// SFile 的摘要描述
/// 處理上傳檔案
/// 如同asp的server_file
/// </summary>
public class SFile
{
	public SFile()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}

    //建立檔案目錄
    public static void create_folder(string folder1, string folder2) {
        string check_folder = "";

        if (!Directory.Exists(folder1)) Directory.CreateDirectory(folder1);

        check_folder = folder1;

        string[] aryfolder2 = folder2.Split('\\');

        //逐層判斷目錄是否存在，若不存在則建立
        for (int i = 0; i < aryfolder2.Length; i++) {
            check_folder += "\\" + aryfolder2[i];
            if (HttpContext.Current.Request.Form["chkTest"] != null && HttpContext.Current.Request.Form["chkTest"] == "TEST") HttpContext.Current.Response.Write("建立檔案目錄：" + check_folder + "<br/>");
            if (!Directory.Exists(check_folder)) Directory.CreateDirectory(check_folder);
        }
    }

    //取得備份檔案名稱
    public static string backup_file(string tfile) {
        string tfile_back = "";
        int n = tfile.LastIndexOf(".");   //副檔名的起始位置

        tfile_back = tfile.Substring(0, n) + "_" + DateTime.Now.Year + Funcs.Right("0" + DateTime.Now.Month, 2) + Funcs.Right("0" + DateTime.Now.Day, 2) + Funcs.Right("0" + DateTime.Now.Hour, 2) + Funcs.Right("0" + DateTime.Now.Minute, 2) + Funcs.Right("0" + DateTime.Now.Second, 2) + tfile.Substring(n);
        
        return tfile_back;
    }

    //rename上傳文件
    public static void rename_attach(SqlCommand cmd, string prgid, int HTProgRight, string HTProgCap, string source, string source_sqlno, string attach_sqlno)
    {
        //系統文件(舊)：
        //web01：將文件從web01/job/data/temp 移到 web01/job/data/sysdoc
        //sin20：將文件從sin20/job/data/temp 移到 \\sin07\systemDoc        
        //web08：將文件從\\web08\data\document\Njob\temp 移到 \\web08\data\document\SystemDoc
        //sin17：將文件從\\sin17\data$\document\Njob\temp 移到 \\sin07\systemDoc  

        //工作管理文件
        //web08：將文件從\\web08\data\document\Njob\temp\上傳人員薪號 → \\web08\data\document\文件主分類\上傳年度(西元)
        //sin25：將文件從\\sin17\data$\document\Njob\temp\上傳人員薪號 → \\sin07\data\document\文件主分類\上傳年度(西元)

        int nrec = 1;

        string isql = "";
        string usql = "";

        string attach_folder = "";
        string attach_path = "";
        string attach_name = "";

        //求取需執行rename的檔案
        isql = "SELECT a.*";
        isql += ", (SELECT code_name FROM cust_code WHERE code_type = 'doc_class1' AND cust_code = a.doc_class1) AS doc_class1nm";
        isql += " FROM job_attach AS a";
        isql += " WHERE 1=1";
        
        if (attach_sqlno.Length > 0)
        {
            isql += " AND a.attach_sqlno = " + Funcs.pkN(attach_sqlno);
        }
        else
        {
            isql += " AND a.source = " + Funcs.pkStr(source);
            isql += " AND a.source_sqlno = " + Funcs.pkN(source_sqlno);
        }
        
        isql += " AND a.attach_flag<>'D'";
        isql += " AND ISNULL(a.rename_flag, '')<>'Y'";
        isql += " ORDER BY a.attach_sqlno";

        if ((HTProgRight & 512) > 0 && HttpContext.Current.Request.Form["chkTest"] != null && HttpContext.Current.Request.Form["chkTest"] == "TEST") HttpContext.Current.Response.Write(isql + "<br/><br/>");
        cmd.CommandText = isql;
        SqlDataReader dr = cmd.ExecuteReader();

        while (dr.Read())
        {
            switch (dr["source"].ToString().Trim().ToLower())
            { 
                case "job":
                    //工作：文件主分類/上傳年度(西元)
                    attach_folder = "upload/job/" + dr["doc_class1"].ToString().Trim() + "_" + dr["doc_class1nm"].ToString().Trim() + "/" + Convert.ToDateTime(dr["in_date"].ToString()).Year;
                    break;
                case "comp_att":
                    //廠商聯絡人
                    attach_folder = "upload/comp_att/" + dr["doc_class1"].ToString().Trim() + "_" + dr["doc_class1nm"].ToString().Trim() + "/" + Convert.ToDateTime(dr["in_date"].ToString()).Year;
                    break;
                case "comp_job":
                    //廠商往來紀錄
                    attach_folder = "upload/comp_job/" + dr["doc_class1"].ToString().Trim() + "_" + dr["doc_class1nm"].ToString().Trim() + "/" + Convert.ToDateTime(dr["in_date"].ToString()).Year;
                    break;
            }

            //if (source == "job")
            //    attach_name = dr["jobno"].ToString() + "_" + dr["source_name"].ToString();       
            //else
            //    attach_name = dr["source_name"].ToString();

            //求取原始檔案之副檔名
            int n = dr["source_name"].ToString().LastIndexOf(".");
            string sExt = (n < 0) ? "" : dr["source_name"].ToString().Substring(n);       

            //檔名rename規則：原始檔名+_+attach_sqlno
            attach_name = dr["source_name"].ToString().Replace(sExt, "") + "_" + dr["attach_sqlno"].ToString() + sExt;

            attach_path = attach_folder + "/" + attach_name;

            //檢查檔案是否存在，若存在才執行rename    
            if (File.Exists(HttpContext.Current.Server.MapPath("~/" + dr["attach_path"].ToString())))
            {
                if ((HTProgRight & 512) > 0 && HttpContext.Current.Request.Form["chkTest"] != null && HttpContext.Current.Request.Form["chkTest"] == "TEST")
                {
                    HttpContext.Current.Response.Write("原檔名：" + dr["attach_name"].ToString() + "<br/>");
                    HttpContext.Current.Response.Write("新檔名：" + attach_name + "<br/>");
                    HttpContext.Current.Response.Write("rename後路徑：" + attach_folder + "/" + attach_name + "<br/><br/>");
                }

                //判斷目的地檔案目錄是否存在，若不存在則建立
                //if (source == "sysdoc")
                //    SFile.create_folder(HttpContext.Current.Application["systemdoc"].ToString(), HttpContext.Current.Server.MapPath("~/" + attach_folder).Replace(HttpContext.Current.Application["systemdoc"].ToString(), ""));
                //else
                //    SFile.create_folder(HttpContext.Current.Application["upload"].ToString(), HttpContext.Current.Server.MapPath("~/" + attach_folder).Replace(HttpContext.Current.Application["upload"].ToString(), ""));

                SFile.create_folder(HttpContext.Current.Application["upload"].ToString(), HttpContext.Current.Server.MapPath("~/" + attach_folder).Replace(HttpContext.Current.Application["upload"].ToString(), ""));

                //判斷目的地檔案是否存在，若存在則rename
                if (File.Exists(HttpContext.Current.Server.MapPath("~/" + attach_folder) + "\\" + attach_name))
                    File.Move(HttpContext.Current.Server.MapPath("~/" + attach_folder) + "\\" + attach_name, HttpContext.Current.Server.MapPath("~/" + attach_folder) + "\\" + SFile.backup_file(attach_name));

                //移動檔案
                File.Move(HttpContext.Current.Server.MapPath("~/" + dr["attach_path"].ToString()), HttpContext.Current.Server.MapPath("~/" + attach_folder) + "\\" + attach_name);

                //組usql，最後執行一次update
                usql += " UPDATE job_attach SET";
                usql += " attach_path = " + Funcs.pkNStr(attach_path);
                usql += ", attach_name = " + Funcs.pkNStr(attach_name);
                usql += ", rename_flag = 'Y'";
                usql += ", tran_date = GETDATE()";
                usql += ", tran_scode = " + Funcs.pkStr(HttpContext.Current.Session["scode"].ToString());
                usql += " OUTPUT 'U', GETDATE(), " + Funcs.pkStr(HttpContext.Current.Session["scode"].ToString()) + "," + Funcs.pkStr(prgid) + "," + "DELETED.* INTO job_attach_log";
                usql += " WHERE attach_sqlno = " + Funcs.pkN(dr["attach_sqlno"].ToString());
                usql += ";";
            }
        }
        dr.Close();

        if (usql.Length > 0) {
            //執行sql
            if ((HTProgRight & 512) > 0 && HttpContext.Current.Request.Form["chkTest"] != null && HttpContext.Current.Request.Form["chkTest"] == "TEST") HttpContext.Current.Response.Write(usql + "<br/><br/>");
            cmd.CommandText = usql;
            nrec = cmd.ExecuteNonQuery();
            if (nrec <= 0) throw new System.Exception(HTProgCap + "-rename失敗！");
        }
    }  



}
