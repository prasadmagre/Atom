using Atom.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Atom.KeyWordsManagers
{
    class TestCaseReader
    {

        /// <summary>
        /// Gets Excel sheet data from respective input modulename (or) SharedSteps (or) SharedSteps(Common folder)
        /// </summary>
        /// <param name="moduleName"></param>
        /// <param name="value"></param>
        /// <param name="excelSheetName"></param>
        /// <returns></returns>
        public DataSet GetExcelSheetData(string moduleName, string value, string excelSheetName)
        {
            string module = PageLabelTableManager.GetTestContext().Properties["Module"].ToString();
            string TestScenarioFileNameWithPath = string.Empty;
            DataSet TestCaseDS = null;
            ExcelDataProvider excelDataProviderObj = new ExcelDataProvider();
            if (value.ToLower().Contains("sharedsteps"))
            {
                moduleName = module;
                TestScenarioFileNameWithPath = Path.Combine(ConfigHelper.TestCaseFolderPath, value + "\\" + moduleName + "\\", excelSheetName);
                bool fileExist = File.Exists(TestScenarioFileNameWithPath);
                if (fileExist)
                {
                    TestCaseDS = excelDataProviderObj.ReadExcel(TestScenarioFileNameWithPath, ConfigHelper.TestCaseSheetName, PageLabelTableManager.GetTestContext());
                }
                else
                {
                    TestScenarioFileNameWithPath = Path.Combine(ConfigHelper.TestCaseFolderPath, value + "\\" + "Common\\", excelSheetName);
                    TestCaseDS = excelDataProviderObj.ReadExcel(TestScenarioFileNameWithPath, ConfigHelper.TestCaseSheetName, PageLabelTableManager.GetTestContext());
                }
            }
            else
            {
                TestScenarioFileNameWithPath = Path.Combine(ConfigHelper.TestCaseFolderPath, module, excelSheetName);
                TestCaseDS = excelDataProviderObj.ReadExcel(TestScenarioFileNameWithPath, ConfigHelper.TestCaseSheetName, PageLabelTableManager.GetTestContext());
            }
            // HF.LogMessage($"Test case folder path: {TestScenarioFileNameWithPath}");

            return TestCaseDS;
        }
    }
}
