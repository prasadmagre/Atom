using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.PhantomJS;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Atom.SeleniumCore.WebDriverManagers
{
    class WebDriverFactory
    {
        public IWebDriver CreateDriverInstance(String BrowserName)
        {
            IWebDriver driver = null;
            try
            {
                switch (BrowserName)
                {
                    case "ie":
                        InternetExplorerOptions options = new InternetExplorerOptions()
                        {
                            EnsureCleanSession = true,
                            InitialBrowserUrl = string.Empty,
                            IgnoreZoomLevel = true,
                            //PageLoadStrategy = InternetExplorerPageLoadStrategy.Eager

                        };
                        //options.ForceCreateProcessApi = true;
                        // options.BrowserCommandLineArguments = "-private";
                        driver = new InternetExplorerDriver("", options, new TimeSpan(0, 5, 0));
                        driver.Manage().Window.Maximize();
                        // driver = new InternetExplorerDriver(ConfigHelper.DriverPath);
                        break;
                    case "chrome":
                        // var chromeOptions = new ChromeOptions();
                        //chromeOptions.AddArgument("incognito");
                        driver = new ChromeDriver("C:\\Users\\prmagre\\Downloads\\chromedriver_win32");
                        //driver.Manage().Window.Maximize();
                        //driver = new ChromeDriver(ConfigHelper.DriverPath);
                        break;
                    case "firefox":
                        driver = new FirefoxDriver();
                        break;
                    case "edge":
                        driver = new EdgeDriver("");
                        break;
                    default:
                        driver = new PhantomJSDriver();
                        //driver.Manage().Timeouts().ImplicitWait(new TimeSpan(0, 0, 25));
                        break;
                }
                return driver;
            }
            catch (Exception e)
            {
                //Utilities.Logger.LogMessage($"Failed to launch the browser, \n Exception: {exception.Message} \n Details: {exception.StackTrace}");
                return null;
            }
        }
    }
}
