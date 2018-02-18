using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Atom.Utilities
{
    [TestClass]
    public class GetExcelTestData
    {
        public static Dictionary<String, Dictionary<String, Dictionary<String, String>>> allVariable = new Dictionary<String, Dictionary<String, Dictionary<String, String>>>();
       [TestMethod]
        public  void GetAllTestData()
        {
            var projectPath = ConfigHelper.ProjectFolderPath;
            string[] fileEntries = Directory.GetFiles(projectPath + "\\Utilities\\TestDataExcel\\");
            foreach (string filePath in fileEntries)
            {
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                String[] excelSheets = GetAllSheetNames(filePath);
                Dictionary<String, Dictionary<String, String>> excelData = new Dictionary<String, Dictionary<String, String>>();
                for (int j = 0; j < excelSheets.Length; j++)
                {
                    Dictionary<String, String> sheetVariables = new Dictionary<String, String>();
                    DataSet ds = new DataSet();
                    string FilterValue = "SELECT * FROM [" + excelSheets[j] + "$] ";
                    
                    ds = ExcelFileReader.ReadExcel(filePath, FilterValue);


                    DataTable allVaiablesData = ds.Tables["MasterTestPlan"];


                    foreach (DataRow row in allVaiablesData.Rows)
                    {

                        string variableName = row[0].ToString().ToUpper().Trim();
                        string variableValue = row[1].ToString().Trim();
                        //moduleData.Add(Keyword, uiElement);
                        try
                        {
                            sheetVariables.Add(variableName, variableValue);
                        }
                        catch (Exception)
                        {

                        }
                        
                        
                    }
                    excelData.Add(excelSheets[j].ToLower(), sheetVariables);
                }
                allVariable.Add(fileName.ToLower(), excelData);
            }
            string value = allVariable["DE_Departemnt".ToLower()]["Prasad".ToLower()]["M_DE_GL_C1_DE1"];
            string value2 = allVariable["DE_Departemnt".ToLower()]["Prasad2".ToLower()]["Customer"];
        }

        public static String[] GetAllSheetNames(String filePath)
        {
            DataTable dt = ExcelFileReader.ReadExcel(filePath);
            String[] excelSheets = new String[dt.Rows.Count];
            int i = 0;

            // Add the sheet name to the string array.
            foreach (DataRow row in dt.Rows)
            {
                excelSheets[i] = row["TABLE_NAME"].ToString().Replace("$","");
                i++;
            }

           
            return excelSheets;
        }
       
    }
}
