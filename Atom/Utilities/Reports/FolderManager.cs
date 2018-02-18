using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Atom.Utilities.Reports
{
    class FolderManager
    {


        /// <summary>
        /// Method to create folder for test result
        /// </summary>
        /// <returns>dir path</returns>

        public static String createTestRunFolder()
        {

            //var projectPath = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).ToString());
            var projectPath = ConfigHelper.ProjectFolderPath;

            string dirName = string.Format("{0}{1}{2}{3:MM-dd-yyy}", projectPath, "\\", "TestResult\\", DateTime.Now);

            Directory.CreateDirectory(dirName);

            return dirName;


        }


        /// <summary>
        /// Method to create screenshot folder for test result
        /// </summary>
        /// <returns>screenshot dir path</returns>
        public static String createScreenshotFolderRunFolder()
        {
            ///String path = null;
            //String projectPath="";
            //if (ConfigHelper.TestCaseRunOn.Equals("1"))
            //{
            // projectPath = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).ToString()).ToString();
            //}
            //else
            //{
            //    projectPath = ConfigHelper.ScreenshotsPath;
            //}
            var projectPath = ConfigHelper.ProjectFolderPath;


            string dirName = string.Format("{0}{1}{2}{3:MM-dd-yyy}{4}", projectPath, "\\", "TestResult\\", DateTime.Now, "\\Screenshot");

            Directory.CreateDirectory(dirName);
            return dirName;

        }
    }
}
