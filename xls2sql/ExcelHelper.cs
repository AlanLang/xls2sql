using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xls2sql
{
    public class ExcelHelper
    {
        private string _filepath, connString;
        private OleDbConnection objConn = new OleDbConnection();
        private DataTable dt;
        public string returnMsg;
        public ExcelHelper()
        {

        }
        public ExcelHelper(string FilePath)
        {
            _filepath = FilePath;
            if(Path.GetExtension(_filepath) == ".xlsx")
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _filepath + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1';";
            else
                connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _filepath + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1';";
            objConn.ConnectionString = connString;
            objConn.Open();
        }
        public string[] GetSheetNames
        {
            get
            {
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dt == null)
                {
                    return null;
                }

                String[] excelSheets = new String[dt.Rows.Count];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    excelSheets[i] = dt.Rows[i]["TABLE_NAME"].ToString();
                }
                return excelSheets;
            }
        }
        public DataTable GetSheetTable(int index)
        {
            return GetSheetTable(GetSheetNames[index]);
        }
        public DataTable GetSheetTable(string SheetName)
        {
            dt = new DataTable();
            OleDbDataAdapter myDa = new OleDbDataAdapter("select * from [" + SheetName + "]", objConn);
            myDa.Fill(dt);
            myDa.Dispose();
            return dt;
        }
        public void Close()
        {
            objConn.Close();
            objConn.Dispose();
        }
    }
}
