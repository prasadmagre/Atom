using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;

namespace Atom.Utilities
{
    public class ExcelFileReader
    {
        /// <summary>
        /// Method to Read excel file.
        /// </summary>
        /// <param name="fullFilePath"></param>
        /// <param name="filterCriteria"></param>
        /// <returns></returns>
        public static DataTable ReadExcel(string fullFilePath)
        {
            OleDbConnection oledbConn=null;
            OleDbCommand cmd = new OleDbCommand();
            OleDbDataAdapter oleda = new OleDbDataAdapter();
            System.Data.DataTable dt = null;
            //OleDbConnection oledbConn = null;
            try
            {
                if (Path.GetExtension(fullFilePath) == ".xls")
                {
                    oledbConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fullFilePath + "; Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"");
                }
                else if (Path.GetExtension(fullFilePath) == ".xlsx")
                {
                    oledbConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + fullFilePath + "; Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1;OLE DB Services = -1;\"");
                }

                // Trace.WriteLine("Connection State  before opening is : " + oledbConn.State);
                if (oledbConn.State != ConnectionState.Open)
                {
                    oledbConn.Open();
                }
                //DataSet ds = new DataSet();
                dt = oledbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                
                //cmd.Connection = oledbConn;

                //Logging to Context
                Trace.WriteLine("Successfully Read the Data File: " + fullFilePath);
                return dt;
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
                //TestContext.WriteLine(e.ToString());
                return null;
            }
            finally
            {
                if (oledbConn.State != ConnectionState.Closed)
                {
                    oledbConn.Close();
                }
                //Trace.WriteLine("Connection State is after closing :" + oledbConn.State);
            }
        }

        /// <summary>
        /// Function for returning the Excel data
        /// </summary>
        /// <param name="fullFilePath"></param>
        /// <param name="SheetName"></param>
        /// <param name="filterCriteria"></param>
        /// <param name="TestContext"></param>
        /// <returns></returns>
        public static DataSet ReadExcel(string fullFilePath, string filterCriteria)
        {
            OleDbConnection oledbConn = null;
            OleDbCommand cmd = new OleDbCommand();
            OleDbDataAdapter oleda = new OleDbDataAdapter();
            try
            {
                if (Path.GetExtension(fullFilePath) == ".xls")
                {
                    oledbConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fullFilePath + "; Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"");
                }
                else if (Path.GetExtension(fullFilePath) == ".xlsx")
                {
                    oledbConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + fullFilePath + "; Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1;OLE DB Services = -1;\"");
                }

                // Trace.WriteLine("Connection State  before opening is : " + oledbConn.State);
                //if (oledbConn.State != ConnectionState.Open)
                //{
                //    oledbConn.Open();
                //}
                DataSet ds = new DataSet();
                using (cmd.Connection = oledbConn)
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = filterCriteria;
                    oleda = new OleDbDataAdapter(cmd);
                    oleda.Fill(ds, "MasterTestPlan");
                }
                //cmd.Connection = oledbConn;

                //Logging to Context
                Trace.WriteLine("Successfully Read the Data File: " + fullFilePath);
                return ds;
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
                //TestContext.WriteLine(e.ToString());
                return null;
            }
            finally
            {
                if (oledbConn.State != ConnectionState.Closed)
                {
                    oledbConn.Close();
                }
                //Trace.WriteLine("Connection State is after closing :" + oledbConn.State);
            }
        }
    }
}
