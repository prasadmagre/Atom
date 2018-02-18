using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Atom.Utilities
{
   public class PageLabelTableManager
    {
        public static ThreadLocal<Dictionary<string, string>> pageLabelsCollection = new ThreadLocal<Dictionary<string, string>>();
        public static ThreadLocal<Dictionary<string, DataTable>> pageTablesCollection = new ThreadLocal<Dictionary<string, DataTable>>();
        public static ThreadLocal<Dictionary<string, List<string>>> pageListsCollection = new ThreadLocal<Dictionary<string, List<string>>>();
        public static ThreadLocal<TestContext> testcontextCollection = new ThreadLocal<TestContext>();
        public static ThreadLocal<String> URLTypeCollection = new ThreadLocal<String>();
        //public static ThreadLocal<SQLDataProvider> sqlDataProviderCollection = new ThreadLocal<SQLDataProvider>();
        //wpublic static ThreadLocal<TestContext> testcontextCollection = new ThreadLocal<TestContext>();

        public static void SetPageLabels(Dictionary<string, string> pageLabels)
        {
            pageLabelsCollection.Value = pageLabels;
        }

        public static Dictionary<string, string> GetPageLabels()
        {
            return pageLabelsCollection.Value;
        }
        public static void SetPageTables(Dictionary<string, DataTable> pageTables)
        {
            pageTablesCollection.Value = pageTables;
        }

        public static Dictionary<string, DataTable> GetPageTables()
        {
            return pageTablesCollection.Value;
        }
        public static void SetPageLists(Dictionary<string, List<string>> pageLists)
        {
            pageListsCollection.Value = pageLists;
        }

        public static Dictionary<string, List<string>> GetPageList()
        {
            return pageListsCollection.Value;
        }


        public static void SetTestContext(TestContext testContext)
        {
            testcontextCollection.Value = testContext;
        }

        public static TestContext GetTestContext()
        {
            return testcontextCollection.Value;
        }

        public static void SetURLType(string UrlType)
        {
            URLTypeCollection.Value = UrlType;
        }

        public static String GetURLType()
        {
            return URLTypeCollection.Value;
        }

    }
}
