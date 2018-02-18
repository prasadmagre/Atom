using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Atom.Utilities
{
    public class ConfigHelper
    {

        
        private static JObject ReadConfiguration(string filename)
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), filename);
            return JObject.Parse(File.ReadAllText(path));
        }

        private static string GetConfigValue(string key) => Configuration[key] != null ? Configuration[key].ToString().Trim() : null;

        //    #endregion

        //    /*Static methods for reading different configurations*/
        //    #region

        public static string BrowserType { get; set; } //= GetConfigValue("BrowserType");

        public static JObject Configuration { get; } = ReadConfiguration("appsettings.json");

        public static JObject ControlObjects { get; } = ReadConfiguration("objects.json");

        public static string LoginPage { get; } = GetConfigValue("LoginPage");

        public static string ProjectFolderPath { get; set; } = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).ToString()).ToString();

        public static string TFSService { get; } = GetConfigValue("TFSService");

        //public static string ADDomain { get; } = GetConfigValue("ADDomain");

        //public static string TFSProject { get; } = GetConfigValue("TFSProject");

        public static string DriverPath { get; } = Path.Combine(ProjectFolderPath, GetConfigValue("DriverPath"));

        public static string TestCaseSheetName { get; } = GetConfigValue("TestCaseSheet");

        public static string TestCaseFolderName { get; } = GetConfigValue("TestCaseFolderName");

        public static string ScreenshotsPath { get; } = Path.Combine(ProjectFolderPath, GetConfigValue("ScreenshotsPath"));

        public static string TestCaseFolderPath { get; } = Path.Combine(ProjectFolderPath, TestCaseFolderName);

        public static char ObjectPropertyValueSeparator { get; } = Convert.ToChar(GetConfigValue("ObjectPropertyValueSeparator"));

        public static string DBQueryFolderName { get; } = GetConfigValue("DBQueryFolderName");

        public static string DBQueryFolderPath { get; } = Path.Combine(ProjectFolderPath, DBQueryFolderName);

        public static bool IsScreenCaptureEnabledOnStepFailure { get; } = Convert.ToBoolean(GetConfigValue("CaptureScreenshotonstepFailure"));

        //public static EnvironmentType Environment { get; private set; } //= TFSHelper.GetTestRunTitle();

        public static string ConnectionString { get; private set; } //= SetEnvironmentConnectionString();

        public static string BOURL { get; private set; } //= SetSiteURL();

        public static string CFOURL { get; private set; } //= SetSiteURL();



        public static int MinimumWait { get; } = Convert.ToInt32(GetConfigValue("MinimumWait"));

        public static int MediumWait { get; } = Convert.ToInt32(GetConfigValue("MediumWait"));

        public static int MaximumWait { get; } = Convert.ToInt32(GetConfigValue("MaximumWait"));

        public static int ExplicitWait { get; } = Convert.ToInt32(GetConfigValue("ExplicitWait"));

        public static string SeleniumGirdHubURL { get; } = GetConfigValue("SeleniumHubURL");



    }
}
