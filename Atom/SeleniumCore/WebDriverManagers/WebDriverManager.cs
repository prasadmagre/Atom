using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Atom.SeleniumCore.WebDriverManagers
{
    class WebDriverManager
    {
        public static ThreadLocal<IWebDriver> driverCollection = new System.Threading.ThreadLocal<IWebDriver>();

        public static void SetDriver(IWebDriver driver1)
        {
            driverCollection.Value = driver1;
        }

        public static IWebDriver GetWebdriver()
        {
            return driverCollection.Value;
        }
    }
}
