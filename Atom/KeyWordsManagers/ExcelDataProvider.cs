using Atom.Utilities;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Atom.KeyWordsManagers
{
    class ExcelDataProvider
    {
        /// <summary>
        /// Function for returning the Excel data
        /// </summary>
        /// <param name="fullFilePath"></param>
        /// <param name="SheetName"></param>
        /// <param name="filterCriteria"></param>
        /// <param name="TestContext"></param>
        /// <returns></returns>
        public DataSet ReadExcel(string fullFilePath, string filterCriteria)
        {
            OleDbConnection oledbConn = null;
            OleDbCommand cmd = new OleDbCommand();
            OleDbDataAdapter oleda = new OleDbDataAdapter();
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
        /// <summary>
        /// FUnction to Read Data from Excel Source
        /// </summary>
        /// <param name="fullFilePath">FUll Path of Excel File</param>
        /// <param name="sheetName">SHeet Name in Excel File</param>
        /// <returns>DataSet Object</returns>
        public DataSet ReadExcel(string fullFilePath, string sheetName, TestContext testContext)
        {
            try
            {

                DataSet ds = new DataSet();
                string FilterValue = "SELECT * FROM [" + sheetName + "$] Where IncludeInTestRun='yes'";
                ds = this.ReadExcel(fullFilePath, FilterValue);
                return ds;
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
                //TestContext.WriteLine(e.ToString());
                return null;
            }
        }

        /// <summary>
        /// Reads data fe
        /// </summary>
        /// <param name="fullFilePath"></param>
        /// <param name="sheetName"></param>
        /// <param name="testContext"></param>
        /// <returns></returns>
        public DataTable ReadDataCSV(string fullFilePath, string csvFileName)
        {
            System.Data.DataTable dt = null;
            try
            {

                string[] csvRows = System.IO.File.ReadAllLines(ConfigHelper.ProjectFolderPath + fullFilePath + csvFileName);
                string[] fields = null;
                dt = new System.Data.DataTable();
                fields = csvRows[0].Split(',');
                for (int i = 0; i < fields.Length; i++)
                {
                    dt.Columns.Add(fields[i]);
                }
                for (int i = 1; i < csvRows.Length; i++)
                {
                    fields = csvRows[i].Split(',');
                    System.Data.DataRow row = dt.NewRow();
                    row.ItemArray = fields;
                    dt.Rows.Add(row);
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
                //TestContext.WriteLine(e.ToString());
                return null;
            }

            return dt;
        }

        /// <summary>
        /// Function for returning specific column from an Excel based of search criteria passed
        /// </summary>
        /// <param name="FullFilePath"></param>
        /// <param name="SheetName"></param>
        /// <param name="ColumnName"></param>
        /// <param name="FilterValue"></param>
        /// <param name="objTestContext"></param>
        /// <returns></returns>
        public string GetSpecificColumnfromExcel(string FullFilePath, string SheetName, string ColumnName, string FilterValue, TestContext objTestContext)
        {
            try
            {

                string filter = "SELECT " + ColumnName + " FROM [" + SheetName + "$] Where " + FilterValue;
                DataSet ds = new DataSet();
                ds = this.ReadExcel(FullFilePath, filter);
                return ds.Tables[0].Rows[0][0].ToString();
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
                //  objTestContext.WriteLine(e.ToString());
                return null;
            }
        }
    }
}
