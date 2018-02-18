using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Atom.Utilities.SqlDataManager
{
    public class SQLDataProvider
    {/// <summary>
     /// Method for Getting  data from SQL Server
     /// </summary>
     /// <param name="ConnectionString"></param>
     /// <param name="SQLQuery"></param>
     /// <returns></returns>
        public DataTable ExecuteNonScalar(string ConnectionString, string SQLQuery, TestContext TestContext)
        {
            string Query = SQLQuery;

            try
            {
                using (SqlConnection Connection = new SqlConnection())
                {
                    using (SqlCommand Command = new SqlCommand())
                    {
                        Connection.ConnectionString = ConnectionString;
                        Command.CommandTimeout = 0;
                        Command.CommandType = CommandType.Text;
                        Command.CommandText = SQLQuery;
                        Command.Connection = Connection;
                        Command.Prepare();

                        using (var da = new SqlDataAdapter(Command))
                        {
                            var table = new DataTable("SQLData");
                            da.Fill(table);
                            return table;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("Exception Details: " + e.Message);
                return null;
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="SQLQuery"></param>
        /// <param name="TestContext"></param>
        /// <returns></returns>
        public string ExecuteScalar(string ConnectionString, string SQLQuery, TestContext TestContext)
        {
            string Query = SQLQuery;
            string Result = "";
            SqlConnection conn = null;
            try
            {
                using (conn = new SqlConnection(ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(Query, conn))
                    {
                        cmd.CommandTimeout = 0;
                        conn.Open();
                        Result = cmd.ExecuteScalar().ToString();
                    }
                }
                return Result;
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("ExecuteNonScalar Method Failed. Exception Details: " + e.Message);
                return null;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="SQLQuery"></param>
        /// <param name="TestContext"></param>
        /// <returns></returns>
        public int ExecuteNonQuery(string ConnectionString, string SQLQuery, TestContext TestContext)
        {
            string Query = SQLQuery;
            int Result = -1;
            SqlConnection conn = null;
            try
            {
                using (conn = new SqlConnection(ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(Query, conn))
                    {
                        cmd.CommandTimeout = 0;
                        conn.Open();
                        Result = cmd.ExecuteNonQuery();
                    }
                }
                return Result;
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("ExecuteNonQuery Method Failed. Exception Details: " + e.Message);
                return -1;
            }
        }
    }
}
