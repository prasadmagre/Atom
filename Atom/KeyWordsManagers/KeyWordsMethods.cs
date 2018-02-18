using Atom.SeleniumCore.WebDriverManagers;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Atom.Utilities;
using System.Diagnostics;
using System.Runtime.InteropServices;
using OpenQA.Selenium.Interactions;
using Atom.Utilities.TestDataExcel;

namespace Atom.KeyWordsManagers
{
   public static class KeyWordsMethods
    {
        #region Find object methods

        /// <summary>
        /// Function for Finding a GUI Control on page with given set of Property:PropertyValues
        /// </summary>
        /// <param name="ObjectName"></param>
        /// <param name="BrowserWindow"></param>
        /// <returns></returns>
        private static IWebElement FindObjectWithText(IWebDriver BrowserWindow, string ObjectName, string value)
        {

            IWebElement Control = null;
            int controlindex = 0;

            string[] values = null;
            string propertyName = string.Empty;
            string PropertyValue = string.Empty;
            string Property = string.Empty;
            if (ObjectName.Trim().ToLower() == "na")
            {
                PropertyValue = value;
            }
            else
            {
                values = GetObject(ObjectName).Split(ConfigHelper.ObjectPropertyValueSeparator);
                if (values.Length == 1)
                {
                    PropertyValue = value;
                    Property = values[0];
                }
                else
                {
                    PropertyValue = values[1];
                    Property = values[0];
                }
            }

            if (Property.ToLower().Contains("contains"))
            {
                Property = Property.Remove(Property.IndexOf("contains"), "contains".Length);
            }

            WebDriverWait wait = new WebDriverWait(BrowserWindow, TimeSpan.FromSeconds(60));
            if (controlindex == 0)
            {
                switch (Property.Trim().ToLower())
                {

                    case "innertext":
                        try
                        {
                            Control = BrowserWindow.FindElement(By.LinkText(PropertyValue));
                            break;
                        }
                        catch (Exception e)
                        {
                            MessageLogger.LogMessage("Failed to find the web element" + e.Message);
                            Control = wait.Until(ExpectedConditions.ElementIsVisible(By.LinkText(PropertyValue)));
                            break;
                        }
                    case "partialinnertext":
                        try
                        {
                            Control = BrowserWindow.FindElement(By.PartialLinkText(PropertyValue));
                            break;
                        }
                        catch (Exception e)
                        {
                            MessageLogger.LogMessage("Failed to find the web element" + e.Message);
                            Control = wait.Until(ExpectedConditions.ElementIsVisible(By.PartialLinkText(PropertyValue)));
                            break;
                        }
                    default:
                        IList<IWebElement> controls = GetControlsFromXPath(BrowserWindow, ObjectName, value, Property, PropertyValue);

                        foreach (IWebElement item in controls)
                        {
                            // Value can contain verification property too: Ex: enabled,false
                            if ((item.Displayed || value.Contains("enabled") || value.Contains("selected")))
                            {
                                Control = item;
                                break;
                            }
                            else if (ObjectName.Trim().ToLower() != "na")
                            {
                                controls = item.FindElements(By.XPath("./following-sibling::*"));

                                foreach (IWebElement item1 in controls)
                                {
                                    if (item1.Displayed)
                                    {
                                        Control = item1;
                                        goto Finish;
                                    }
                                }
                            }
                        }
                        Finish:
                        break;
                }
            }
            else
            {
                switch (Property.Trim().ToLower())
                {

                    case "innertext":
                        Control = BrowserWindow.FindElements(By.LinkText(PropertyValue)).ElementAt(controlindex);
                        break;
                    case "partialinnertext":
                        Control = BrowserWindow.FindElements(By.PartialLinkText(PropertyValue)).ElementAt(controlindex);
                        break;

                }

            }

            return Control;
        }

        /// <summary>
        /// Function  for Finding a GUI Control on page with given set of Property:PropertyValues
        /// </summary>
        /// <param name="ObjectName"></param>
        /// <param name="BrowserWindow"></param>
        /// <returns></returns>
        private static IWebElement GetFirstControlWithText(IWebDriver BrowserWindow, string ObjectName, string value)
        {

            IWebElement Control = null;
            int controlindex = 0;

            string[] values = null;
            string propertyName = string.Empty;
            string PropertyValue = string.Empty;
            string Property = string.Empty;
            if (ObjectName.Trim().ToLower() == "na")
            {
                PropertyValue = value;
            }
            else
            {
                values = GetObject(ObjectName).Split(ConfigHelper.ObjectPropertyValueSeparator);
                if (values.Length == 1)
                {
                    PropertyValue = value;
                    Property = values[0];
                }
                else
                {
                    PropertyValue = values[1];
                    Property = values[0];
                }
            }

            if (Property.ToLower().Contains("contains"))
            {
                Property = Property.Remove(Property.IndexOf("contains"), "contains".Length);
            }

            WebDriverWait wait = new WebDriverWait(BrowserWindow, TimeSpan.FromSeconds(60));
            if (controlindex == 0)
            {
                switch (Property.Trim().ToLower())
                {

                    case "innertext":
                        try
                        {
                            Control = BrowserWindow.FindElement(By.LinkText(PropertyValue));
                            break;
                        }
                        catch
                        {
                            Control = wait.Until(ExpectedConditions.ElementIsVisible(By.LinkText(PropertyValue)));
                            break;
                        }
                    case "partialinnertext":
                        try
                        {
                            Control = BrowserWindow.FindElement(By.PartialLinkText(PropertyValue));
                            break;
                        }
                        catch
                        {
                            Control = wait.Until(ExpectedConditions.ElementIsVisible(By.PartialLinkText(PropertyValue)));
                            break;
                        }
                    default:

                        int i = 0;
                        while (i < 15)
                        {

                            IList<IWebElement> controls = GetControlsFromXPath(BrowserWindow, ObjectName, value, Property, PropertyValue);
                            var wElement = controls.Count > 0 ? controls.Select(c => c.Enabled == true ? c : null) : null;


                            if (wElement == null)
                            {
                                Thread.Sleep(1000);
                                i++;
                            }
                            else
                            {
                                string[] PropertyValueWithIndex = PropertyValue.Split(':');

                                if (PropertyValueWithIndex.Count() > 1)
                                {
                                    Control = (IWebElement)wElement.ToList()[Convert.ToInt16(PropertyValueWithIndex[1])];
                                }
                                else
                                {
                                    Control = (IWebElement)wElement.ToList()[0];
                                }
                                break;
                            }
                        }

                        break;
                }
            }
            else
            {
                switch (Property.Trim().ToLower())
                {

                    case "innertext":
                        Control = BrowserWindow.FindElements(By.LinkText(PropertyValue)).ElementAt(controlindex);
                        break;
                    case "partialinnertext":
                        Control = BrowserWindow.FindElements(By.PartialLinkText(PropertyValue)).ElementAt(controlindex);
                        break;

                }

            }

            return Control;
        }

        private static IList<IWebElement> GetControlsFromXPath(IWebDriver BrowserWindow, string ObjectName, string value, string Property, string PropertyValue)
        {
            IList<IWebElement> controls = null;
            if (ObjectName.Trim().ToLower() == "na")
            {
                controls = BrowserWindow.FindElements(By.XPath("//*[text()[normalize-space()='" + PropertyValue + "']]"));
            }
            else
            {
                string[] propValues = value.Split(',');
                if (propValues.Length == 1)
                {
                    controls = BrowserWindow.FindElements(By.XPath("//*[text()[normalize-space()='" + propValues[0] + "']]/following::span[@" + Property + "='" + PropertyValue + "']"));
                }
                else
                {
                    string[] PropertyValueWithIndex = PropertyValue.Split(':');

                    if (PropertyValueWithIndex.Count() > 1)
                    {
                        controls = BrowserWindow.FindElements(By.XPath("//*[text()[normalize-space()='" + propValues[0] + "']]/following::*[text()[normalize-space()='" + propValues[1] + "']]/following::*[contains(@" + Property + ",'" + PropertyValueWithIndex[0] + "')][" + PropertyValueWithIndex[1] + "]"));

                        if (controls.Count == 0)
                        {
                            controls = BrowserWindow.FindElements(By.XPath("//*[text()[normalize-space()='" + propValues[0] + "']]/following::*[text()[normalize-space()='" + propValues[1] + "']]/following::*[text()[normalize-space()='" + PropertyValueWithIndex[0] + "']][" + PropertyValueWithIndex[1] + "]"));
                        }
                    }

                    else
                    {
                        if (Property.ToLower().Trim() != "text")
                        {
                            controls = BrowserWindow.FindElements(By.XPath("//*[text()[normalize-space()='" + propValues[0] + "']]/following::*[text()[normalize-space()='" + propValues[1] + "']]/following::*[contains(@" + Property + ",'" + PropertyValue + "')]"));
                        }
                        else
                        {
                            controls = BrowserWindow.FindElements(By.XPath("//*[text()[normalize-space()='" + propValues[0] + "']]/following::*[text()[normalize-space()='" + propValues[1] + "']]/following::*[text()[normalize-space()='" + PropertyValue + "']]"));
                        }
                    }

                }
            }

            return controls;
        }

        /// <summary>
        /// Function for Finding a GUI Control on page with given set of Property:PropertyValues
        /// </summary>
        /// <param name="ObjectName"></param>
        /// <param name="BrowserWindow"></param>
        /// <returns></returns>
        private static IWebElement FindObject(IWebDriver BrowserWindow, string ObjectName, string valueFromExcel)
        {
            IWebElement Control = null;
            List<IWebElement> ControlList = new List<IWebElement>();
            if (String.IsNullOrEmpty(ObjectName))
            {
                Control = null;
            }
            else
            {
                int controlindex = 0;
                string[] value = null;
                string PropertyValue = string.Empty;
                string Property = string.Empty;
                if (GetObject(ObjectName).Contains('|'))
                {
                    value = GetObject(ObjectName).Split(ConfigHelper.ObjectPropertyValueSeparator);

                    PropertyValue = value[1];
                    Property = value[0];
                    string[] PropertyValueWithIndex = new string[] { };

                    if (!PropertyValue.Contains("::"))
                    {
                        PropertyValueWithIndex = PropertyValue.Split(':');
                    }

                    if (PropertyValueWithIndex.Count() > 1)
                    {
                        PropertyValue = PropertyValueWithIndex[0];
                        controlindex = Convert.ToInt16(PropertyValueWithIndex[1]);
                    }
                    else
                    {
                        PropertyValue = value[1];
                        controlindex = 0;
                    }
                }


                MessageLogger.LogMessage($"Trying to find the control by: {Property}, and control identifier value: {PropertyValue}");
                WebDriverWait wait = new WebDriverWait(BrowserWindow, TimeSpan.FromSeconds(60));
                if (controlindex == 0)
                {
                    switch (Property.Trim().ToLower())
                    {
                        case "class":
                            try
                            {
                                Control = BrowserWindow.FindElement(By.ClassName(PropertyValue));
                            }
                            catch (Exception)
                            {
                                Control = wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName(PropertyValue)));
                            }

                            break;
                        case "css":
                            try
                            {
                                Control = BrowserWindow.FindElement(By.CssSelector(PropertyValue));
                            }
                            catch (Exception)
                            {
                                Control = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(PropertyValue)));
                            }

                            break;
                        case "id":
                            try
                            {
                                IReadOnlyCollection<IWebElement> controls = BrowserWindow.FindElements(By.Id(PropertyValue));
                                Control = controls.Count > 0 ? controls.ElementAt(controlindex) : null;
                                if (Control == null)
                                {
                                    throw new NoSuchElementException();
                                }
                            }
                            catch
                            {
                                Control = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(PropertyValue)));

                            }
                            //    Control = wait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@id='" + PropertyValue + "'][not(contains(@style, 'display: none')) and not(contains(@visibility, 'hidden'))]")));

                            //int i = 0;
                            //while (i < 20)
                            //{
                            //    try
                            //    {

                            //        if (!Control.Enabled)
                            //        {
                            //            Thread.Sleep(1000);
                            //            i++;
                            //        }
                            //        else
                            //        {
                            //            break;
                            //        }
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        goto Finish;
                            //    }60
                            //}
                            break;
                        case "innertext":
                            try
                            {
                                Control = BrowserWindow.FindElement(By.LinkText(PropertyValue));
                            }
                            catch (Exception)
                            {
                                Control = wait.Until(ExpectedConditions.ElementIsVisible(By.LinkText(PropertyValue)));
                            }

                            break;
                        case "name":
                            try
                            {
                                Control = BrowserWindow.FindElement(By.Name(PropertyValue));
                            }
                            catch (Exception)
                            {
                                Control = wait.Until(ExpectedConditions.ElementIsVisible(By.Name(PropertyValue)));
                            }

                            break;
                        case "linktext":
                            try
                            {
                                Control = BrowserWindow.FindElement(By.LinkText(PropertyValue));
                            }
                            catch (Exception)
                            {
                                Control = wait.Until(ExpectedConditions.ElementIsVisible(By.LinkText(PropertyValue)));
                            }

                            break;
                        case "partialinnertext":
                            try
                            {
                                Control = BrowserWindow.FindElement(By.PartialLinkText(PropertyValue));
                            }
                            catch (Exception)
                            {
                                Control = wait.Until(ExpectedConditions.ElementIsVisible(By.PartialLinkText(PropertyValue)));
                            }

                            break;
                        case "tagname":
                            try
                            {
                                Control = BrowserWindow.FindElement(By.TagName(PropertyValue)); ;
                            }
                            catch
                            {
                                Control = wait.Until(ExpectedConditions.ElementIsVisible(By.TagName(PropertyValue)));
                            }
                            break;
                        case "xpath":
                            try
                            {
                                Control = BrowserWindow.FindElement(By.XPath(PropertyValue));
                            }
                            catch
                            {
                                try
                                {
                                    Control = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PropertyValue)));
                                }
                                catch
                                {
                                    Control = BrowserWindow.FindElements(By.XPath(PropertyValue)).ElementAt(controlindex);
                                }

                            }
                            break;
                        case "js":

                            Control = Execute<IWebElement>(BrowserWindow, PropertyValue);
                            break;
                        default:
                            Control = FindObjectWithText(BrowserWindow, ObjectName, valueFromExcel);
                            break;
                    }
                }
                else
                {
                    switch (Property.Trim().ToLower())
                    {
                        case "class":
                            Control = BrowserWindow.FindElements(By.ClassName(PropertyValue)).ElementAt(controlindex);
                            break;
                        case "css":
                            Control = BrowserWindow.FindElements(By.CssSelector(PropertyValue)).ElementAt(controlindex);
                            break;
                        case "id":
                            Control = BrowserWindow.FindElements(By.Id(PropertyValue)).ElementAt(controlindex);
                            break;
                        case "innertext":
                            Control = BrowserWindow.FindElements(By.LinkText(PropertyValue)).ElementAt(controlindex);
                            break;
                        case "name":
                            Control = BrowserWindow.FindElements(By.Name(PropertyValue)).ElementAt(controlindex);
                            break;
                        case "linktext":
                            Control = BrowserWindow.FindElements(By.LinkText(PropertyValue)).ElementAt(controlindex);
                            break;
                        case "partialinnertext":
                            Control = BrowserWindow.FindElements(By.PartialLinkText(PropertyValue)).ElementAt(controlindex);
                            break;
                        case "tagname":
                            Control = BrowserWindow.FindElements(By.TagName(PropertyValue)).ElementAt(controlindex);
                            break;
                        case "xpath":
                            Control = BrowserWindow.FindElements(By.XPath(PropertyValue)).ElementAt(controlindex);
                            break;
                        case "js":
                            Control = Execute<IWebElement>(BrowserWindow, PropertyValue);
                            break;
                        default:
                            
                            Control = FindObjectWithText(BrowserWindow, ObjectName, valueFromExcel);
                            break;
                    }

                }
            }
            return Control;
        }

        public static T Execute<T>(IWebDriver browserWindow, string script)
        {
            T control = default(T);
            try
            {
                control = (T)((IJavaScriptExecutor)browserWindow).ExecuteScript(script);
                MessageLogger.LogMessage($"Successfully executed Javascript: {script}");
            }
            catch (Exception e)
            {
                MessageLogger.LogMessage($"Failed to execute Javascript: {script}");
                MessageLogger.LogMessage($"Exception message: {e.Message}");
            }
            return control;
        }

        /// <summary>
        /// Functon retrieving the Objects inside a parent Object
        /// </summary>
        /// <param name="ChildObjectName"></param>
        /// <param name="ParentControl"></param>
        /// <returns></returns>
        private static IWebElement FindChildObject(IWebElement ParentControl, string ChildObjectName)
        {
            IWebElement Control =null;
            if (String.IsNullOrEmpty(ChildObjectName))
            {
                Control = null;
            }
            else
            {
                //Separating each Property:PropertyValue string and creating search criteria
                string[] value = GetObject(ChildObjectName).Split(ConfigHelper.ObjectPropertyValueSeparator);
                string PropertyValue = value[1];
                string Property = value[0];
                switch (Property.Trim().ToLower())
                {
                    case "class":
                        Control = ParentControl.FindElement(By.ClassName(PropertyValue));
                        break;
                    case "css":
                        Control = ParentControl.FindElement(By.CssSelector(PropertyValue));
                        break;
                    case "id":
                        Control = ParentControl.FindElement(By.Id(PropertyValue));
                        break;
                    case "innertext":
                        Control = ParentControl.FindElement(By.LinkText(PropertyValue));
                        break;
                    case "name":
                        Control = ParentControl.FindElement(By.Name(PropertyValue));
                        break;
                    case "linktext":
                        Control = ParentControl.FindElement(By.LinkText(PropertyValue));
                        break;
                    case "partialinnertext":
                        Control = ParentControl.FindElement(By.PartialLinkText(PropertyValue));
                        break;
                    case "tagname":
                        Control = ParentControl.FindElement(By.TagName(PropertyValue));
                        break;
                    case "xpath":
                        Control = ParentControl.FindElement(By.XPath(PropertyValue));
                        break;
                }
            }

            return Control;
        }

        /// <summary>
        /// Functon retrieving the Objects inside a parent Object
        /// </summary>
        /// <param name="ChildObjectName"></param>
        /// <param name="ParentControl"></param>
        /// <returns></returns>
        private static IWebElement FindChildObject(IWebElement ParentControl, string ChildObjectName, string value)
        {
            IWebElement Control = null;
            //Separating each Property:PropertyValue string and creating search criteria
            string propertyName = GetObject(ChildObjectName);
            string PropertyValue = value;
            string Property = propertyName;
            switch (Property.Trim().ToLower())
            {

                case "innertext":
                    Control = ParentControl.FindElement(By.LinkText(PropertyValue));
                    break;


                case "partialinnertext":
                    Control = ParentControl.FindElement(By.PartialLinkText(PropertyValue));
                    break;

            }


            return Control;
        }

        /// <summary>
        /// Method to identify Object either from Parent or Directly under window then calling the correct FindObject Method
        /// </summary>
        /// <param name="ParentObject"></param>
        /// <param name="ObjectName"></param>
        /// <param name="BrowserWindow"></param>
        /// <returns></returns>
        public static IWebElement FindObject(string ObjectName, IWebDriver BrowserWindow, string valueFromExcel)
        {
            string ParentObject = "";
            string Object = "";
            //=================Check if the Object Definition is having parent control========================//
            if (ObjectName.Contains(":"))
            {
                //Separating each Property:PropertyValue string and creating search criteria
                string[] value = ObjectName.Split(':');
                Object = value[1].Trim();
                ParentObject = value[0].Trim();
            }
            else
            {
                ParentObject = "NA";
                Object = ObjectName.Trim();
            }
            //=================End of Object identification===================================================//

            IWebElement Control = null;
            if (ParentObject == "NA")
            {
                Control = FindObject(BrowserWindow, Object, valueFromExcel);
            }
            else
            {
                Control = FindChildObject(FindObject(ParentObject, BrowserWindow, valueFromExcel), Object);
            }

            return Control;
        }

        /// <summary>
        /// Method to identify Object either from Parent or Directly under window then calling the correct FindObject Method
        /// </summary>
        /// <param name="ParentObject"></param>
        /// <param name="ObjectName"></param>
        /// <param name="BrowserWindow"></param>
        /// <returns></returns>
        public static IWebElement FindObject(string ObjectName, string elemValue, IWebDriver BrowserWindow)
        {
            string ParentObject = "";
            string Object = "";
            //=================Check if the Object Definition is having parent control========================//
            if (ObjectName.Contains(":"))
            {
                //Separating each Property:PropertyValue string and creating search criteria
                string[] value = ObjectName.Split(':');
                Object = value[1].Trim();
                ParentObject = value[0].Trim();
            }
            else
            {
                ParentObject = "NA";
                Object = ObjectName.Trim();
            }
            //=================End of Object identification===================================================//

            IWebElement Control = null;
            if (ParentObject == "NA")
            {
                Control = FindObject(BrowserWindow, Object, elemValue);
            }
            else
            {
                Control = FindChildObject(FindObject(ParentObject, BrowserWindow, elemValue), Object, elemValue);
            }

            return Control;
        }

        #endregion

        /*Public  members of class*/
        #region Selenium Core helper methods

        public static IWebDriver NavigateToUrl(string url)
        {
            try
            {
                ////LogMessage($"Navigating to URL: {url}");
                WebDriverManager.GetWebdriver().Navigate().GoToUrl(url);
                WaitForPageLoad();
            }
            catch (Exception exception)
            {
                MessageLogger.LogMessage($"Failed to navigate to URL: {url} \n Exception: {exception.Message}");
            }
            return WebDriverManager.GetWebdriver();
        }
        public static List<string> GetTextFromListOfControls(string uiElementIdentifier, string uiElementValue)
        {

            List<String> allText = new List<string>();
            if (uiElementValue != null)
            {

                try
                {
                    string[] value = GetObject(uiElementIdentifier).Split(ConfigHelper.ObjectPropertyValueSeparator);
                    string tableXPath = value[1];
                    string Property = value[0];
                    IList<IWebElement> allOptions = WebDriverManager.GetWebdriver().FindElements(By.XPath(tableXPath));

                    foreach (var a in allOptions)
                    {
                        allText.Add(a.Text);
                    }


                    return allText;
                }
                catch (Exception e)
                {
                    MessageLogger.LogMessage("Failed. Invalid Value for Enter Text. Value is NULL.");
                    return null;
                }

            }
            return allText;
        }


        /// <summary>
        /// Maximizes browser window
        /// </summary>
        public static void MaximizeWindow()
        {
            if (Process.GetProcessesByName("chrome").Any())
            {
                Process[] processes = Process.GetProcessesByName("chrome");
                foreach (Process proc in processes)
                {
                    if (proc.MainWindowTitle.Contains("data"))
                    {
                        IntPtr pointer = proc.MainWindowHandle;
                        MaximizeWindow(pointer);
                    }
                }
            }
        }

        private const int SW_SHOWNORMAL = 1;
        private const int SW_SHOWMINIMIZED = 2;
        private const int SW_SHOWMAXIMIZED = 3;

        /// <summary>
        /// Maximizes window without waiting for browser to load
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="nCmdShow"></param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        /// <summary>
        /// Maximizes browser window after it is loaded
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="nCmdShow"></param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        /// <summary>
        /// Maximizes browser window
        /// </summary>
        /// <param name="win"></param>
        private static void MaximizeWindow(IntPtr win)
        {

            if (!win.Equals(IntPtr.Zero))
            {
                ShowWindowAsync(win, SW_SHOWMAXIMIZED);
            }
        }

        public static IWebDriver LaunchAndNavigate(string browserType, string url)
        {
            IWebDriver driver = null;
            WebDriverFactory webDriverFactoryObj = new WebDriverFactory();
            driver = webDriverFactoryObj.CreateDriverInstance(ConfigHelper.BrowserType);
            WebDriverManager.SetDriver(driver);
            //LaunchBrowser(browserType);
            return NavigateToUrl(url);
        }
        /// <summary>
        /// Function for finding and Entering value in a Editbox
        /// </summary>
        /// <param name="TextBoxIdentifier"></param>
        /// <param name="Value"></param>
        /// <param name="BrowserWindow"></param>
        /// <returns></returns>
        public static bool EnterText(string TextBoxIdentifier, string Value)
        {
            bool Result = false;
            try
            {
                //Value = VariableManager.GetVariableValue(Value);
                var control = SetFocus(TextBoxIdentifier, Value);
                try
                {
                    control.Clear();
                }
                catch
                {
                    //do nothing;
                }
                control.SendKeys(Value);
                Result = true;
                MessageLogger.LogMessage($"EnterText: { TextBoxIdentifier}: { Value.Replace("^{a}", string.Empty)}");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed EnterText: {TextBoxIdentifier} : {Value.Replace("^{a}", string.Empty)} \n Exception Details : {e.Message}");
            }
            return Result;
        }

        public static bool EnterTextOnPopUp(string TextBoxIdentifier, string Value)
        {
            bool Result = false;
            try
            {
                Value = VariableManager.GetVariableValue(Value);
                var control = WebDriverManager.GetWebdriver().SwitchTo().ActiveElement();
                control = FindChildObject(control, TextBoxIdentifier);
                control.Clear();
                control.SendKeys(Value);
                Result = true;
                MessageLogger.LogMessage($"EnterText: { TextBoxIdentifier}: { Value.Replace("^{a}", string.Empty)}");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed EnterText: {TextBoxIdentifier} : {Value.Replace("^{a}", string.Empty)} \n Exception Details : {e.Message}");
            }
            return Result;
        }

        public static bool EnterSpecialKeys(string ControlIdentifier, string Value)
        {
            bool Result = false;
            try
            {
                var control = FindObject(ControlIdentifier, WebDriverManager.GetWebdriver(), Value);
                //control.Click();
                EnterSpecialKeys(control, Value);
                WaitForPageLoad();
                Result = true;
                MessageLogger.LogMessage($"Entered Text: { ControlIdentifier }, value :  { Value}");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed EnterText: {ControlIdentifier} : {Value} \n Exception Details: {e.Message}");
            }
            return Result;
        }

        private static void WaitForPageLoad()
        {
            //  browserWindow.Manage().Timeouts().ImplicitWait= TimeSpan.FromSeconds(180);
        }

        public static bool EnterSpecialKeys(IWebElement uiElement, string key)
        {
            switch (key.ToLower())
            {
                case "tab":
                    uiElement.SendKeys(Keys.Tab);
                    break;
                case "enter":
                    uiElement.SendKeys(Keys.Enter);
                    break;
                case "down":
                    uiElement.SendKeys(Keys.Down);
                    break;
                case "arrowright":
                    uiElement.SendKeys(Keys.ArrowRight);
                    break;
                case "ctrl+a+backspace":
                    uiElement.SendKeys(Keys.Control + "a" + Keys.Backspace);
                    break;
                case "backspace":
                    uiElement.SendKeys(Keys.Backspace);
                    break;
                default:
                    break;
            }
            return true;
        }

        /// <summary>
        /// Method for setting focus on a control. This step can be used in excel sheet before pwerforming any action
        /// </summary>
        /// <param name="ControlIdentifier"></param>
        /// <param name="BrowserWindow"></param>
        /// <returns></returns>
        public static IWebElement SetFocus(string ControlIdentifier, string valueFromExcel)
        {
            IWebElement Control = null;
            try
            {
                KeyWordsMethods.WaitForPageLoad();
                MinWait();
                Control = FindObject(ControlIdentifier, WebDriverManager.GetWebdriver(), valueFromExcel);
                //new Actions(BrowserWindow).MoveToElement(Control).Perform();   
                // -- Replacing above statement with below statement as set focus working correctly in IE by Javascript             
                ((IJavaScriptExecutor)WebDriverManager.GetWebdriver()).ExecuteScript("arguments[0].focus()", Control);
                MessageLogger.LogMessage($"SetFocus: {ControlIdentifier}");
            }
            catch (Exception e)
            {
                MessageLogger.LogMessage($"Failed SetFocus: {ControlIdentifier} \n Exception Details: {e.Message}");
            }
            return Control;
        }

        /// <summary>
        /// Function to Switches Between Windows
        /// </summary>
        /// <param name="BrowserWindow"></param>
        /// <param name="ControlIdentifier"></param>
        /// <param name="valueFromExcel"></param>
        /// <returns></returns>

        public static bool SwitchToChildWindow(string ControlIdentifier, string valueFromExcel)
        {
            string currentWindowHandle = WebDriverManager.GetWebdriver().CurrentWindowHandle;
            IList<String> allWinHandles = WebDriverManager.GetWebdriver().WindowHandles;

            try
            {
                foreach (string current in allWinHandles)
                {

                    if (!current.Equals(currentWindowHandle))
                    {
                        WebDriverManager.GetWebdriver().SwitchTo().Window(current);
                    }

                }

            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.InnerException.ToString());
            }

            bool Result = true;
            return Result;


        }







        public static void MinWait()
        {
            Thread.Sleep(TimeSpan.FromMilliseconds(500));
        }

        /// <summary>
        /// Method for finding the HtmlInputButton and click it
        /// </summary>
        /// <param name="ControlIdentifier"></param>
        /// <param name="BrowserWindow"></param>
        /// <returns></returns>
        public static bool ClickControl(IWebDriver BrowserWindow, string ControlIdentifier, string valueFromExcel)
        {
            bool Result = false;
            try
            {
                var control = SetFocus(ControlIdentifier, valueFromExcel);
                control.Click();
                WaitForPageLoad();
                Result = true;
                MessageLogger.LogMessage($"ClickButton: {ControlIdentifier}");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed ClickButton: {ControlIdentifier} \n Exception Details: {e.Message}");
            }
            return Result;
        }

        public static bool ClickControlJs(string ControlIdentifier, string valueFromExcel)
        {
            bool Result = false;
            try
            {
                var control = SetFocus(ControlIdentifier, valueFromExcel);
                MessageLogger.LogMessage($"Clicking control using Java Script: {ControlIdentifier}");
                ((IJavaScriptExecutor)WebDriverManager.GetWebdriver()).ExecuteScript("arguments[0].click()", control);
                WaitForPageLoad();
                Result = true;
                MessageLogger.LogMessage($"ClickButton: {ControlIdentifier}");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed ClickButton: {ControlIdentifier} \n Exception Details: {e.Message}");
            }
            return Result;
        }
        public static bool ClickControlOnPopUp(string ControlIdentifier, string valueFromExcel)
        {
            bool Result = false;
            try
            {
                var control = SetFocus(ControlIdentifier, valueFromExcel);
                control.Click();
                //control.SendKeys(Keys.Enter);
                WaitForPageLoad();
                Result = true;
                MessageLogger.LogMessage($"ClickButton: {ControlIdentifier}");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed ClickButton: {ControlIdentifier} \n Exception Details: {e.Message}");
            }
            return Result;
        }

        public static bool ClickControlWithText(string ControlIdentifier, string value)
        {
            bool Result = false;
            try
            {
                if (value.Equals("Go back to the old one"))
                {
                    try
                    {
                        WebDriverManager.GetWebdriver().FindElement(By.Id("uxOptOutLink"));

                    }
                    catch (Exception e)
                    {
                        return true;
                    }
                }
                var control = FindObjectWithText(WebDriverManager.GetWebdriver(), ControlIdentifier, value);
                var firstControl = GetFirstControlWithText(WebDriverManager.GetWebdriver(), ControlIdentifier, value);
                if (!firstControl.Selected)
                {

                    try
                    {
                        WaitForPageLoad();
                        control.Click();
                        WaitForPageLoad();
                        Result = true;
                        MessageLogger.LogMessage($"ClickButton: {ControlIdentifier}");
                    }
                    catch
                    {
                        MessageLogger.LogMessage($"Clicking control using Java Script: {ControlIdentifier}");
                        ((IJavaScriptExecutor)WebDriverManager.GetWebdriver()).ExecuteScript("arguments[0].click()", control);
                        WaitForPageLoad();
                        Result = true;
                        MessageLogger.LogMessage($"ClickButton: {ControlIdentifier}");
                    }
                }
                WaitForPageLoad();
                Result = true;
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed ClickButton: {ControlIdentifier} \n Exception Details: {e.Message}");
            }
            return Result;
        }

        public static bool UnClickControlWithText(string ControlIdentifier, string value)
        {
            bool Result = false;
            try
            {
                var control = FindObjectWithText(WebDriverManager.GetWebdriver(), ControlIdentifier, value);
                var firstControl = GetFirstControlWithText(WebDriverManager.GetWebdriver(), ControlIdentifier, value);
                if (firstControl.Selected)
                {
                    try
                    {
                        control.Click();
                        WaitForPageLoad();
                        Result = true;
                        //LogMessage($"ClickButton: {ControlIdentifier}");
                    }
                    catch
                    {
                        MessageLogger.LogMessage($"Clicking control using Java Script: {ControlIdentifier}");
                        ((IJavaScriptExecutor)WebDriverManager.GetWebdriver()).ExecuteScript("arguments[0].click()", control);
                        WaitForPageLoad();
                        Result = true;
                        MessageLogger.LogMessage($"ClickButton: {ControlIdentifier}");
                    }
                }
                WaitForPageLoad();
                Result = true;
            }
            catch (Exception e)
            {
                Result = false;
            }
            return Result;
        }

        public static string ReadControlInnerText(string Object, string valueFromExcel)
        {
            string innerText = "";
            try
            {
                innerText = SetFocus(Object, valueFromExcel).Text.ToString();
                //LogMessage($"Control inner text: {Object} : {innerText}");
            }
            catch (Exception e)
            {
                MessageLogger.LogMessage($"Failed to find control: { Object} : \n Exception Details: {e.Message}");
                innerText = "NF";
            }
            return innerText;
        }

        public static bool HoverOnControl(string uiElement, string valueFromExcel)
        {
            bool Result = false;
            try
            {

                WebDriverWait wait = new WebDriverWait(WebDriverManager.GetWebdriver(), TimeSpan.FromSeconds(60));
                var element = SetFocus(uiElement, valueFromExcel);

                Actions action = new Actions(WebDriverManager.GetWebdriver());
                action.MoveToElement(element).Perform();

                //SetFocus(uiElement, valueFromExcel);
                Result = true;
                //LogMessage($"hovered on control: {uiElement}");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed to hover on control: {uiElement} \n Exception Details: {e.Message}");
            }
            return Result;
        }

        public static bool DismissPopupWindow()
        {
            bool Result = false;
            try
            {
                WebDriverManager.GetWebdriver().SwitchTo().Alert().Dismiss();
                WaitForPageLoad();
                Result = true;
                //LogMessage("Dismissed the popup window : ");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed to to dismiss popup window, \n Exception Details: {e.Message}");
            }
            return Result;
        }

        public static bool AcceptPopupWindow()
        {
            bool Result = false;
            try
            {
                WebDriverManager.GetWebdriver().SwitchTo().Alert().Accept();
                WaitForPageLoad();
                Result = true;
                //LogMessage("Accepted the popup window ");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed to Accept popup window, \n Exception Details: {e.Message}");
            }
            return Result;
        }

        public static bool VerifyMessageOnPopupWindow(string uiElementValue)
        {
            bool Result = false;
            try
            {
                uiElementValue = VariableManager.GetVariableValue(uiElementValue);
                Result = WebDriverManager.GetWebdriver().SwitchTo().Alert().Text.Trim().Equals(uiElementValue.Trim()) ? true : false;
                //LogMessage($"successfully verified message on popup window: {uiElementValue}");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed to verify message on popup window: {uiElementValue}, \n Exception Details:  {e.Message}");
            }
            return Result;
        }

        public static bool ExecuteJS(string script, string comments)
        {
            try
            {
                //LogMessage($"Executing Java Script: {script}; \n Comments: {comments}");
                WaitForPageLoad();
                ((IJavaScriptExecutor)WebDriverManager.GetWebdriver()).ExecuteScript(script);

                return true;
            }
            catch (Exception e)
            {
                MessageLogger.LogMessage($"Failed to excute script: {script} \n Exception Details: {e.Message}");
                return false;
            }
        }

        /// <summary>
        /// Method for opening the dropdown
        /// </summary>
        /// <param name="Object"></param>
        /// <param name="BrowserWindow"></param>
        /// <param name="TestContext"></param>
        /// <returns></returns>
        public static bool OpenDropDown(string Object, string valueFromExcel)
        {
            bool Result = false;
            try
            {
                SetFocus(Object, valueFromExcel).Click();
                Result = true;
                //LogMessage($"Open Drop Down: {Object}");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed Open Drop Down: {Object} \n Exception Details: {e.Message}");
            }

            return Result;
        }

        /// <summary>
        /// Method for selecting value in a dropdown
        /// </summary>
        /// <param name="uiElement"></param>
        /// <param name="Value"></param>
        /// <param name="BrowserWindow"></param>
        /// <param name="TestContext"></param>
        /// <returns></returns>
        public static bool SelectDropDownItem(string uiElement, string[] dropDownValues)
        {
            bool Result = false;

            try
            {
                IWebElement control = SetFocus(uiElement, string.Empty);
                ((IJavaScriptExecutor)WebDriverManager.GetWebdriver()).ExecuteScript("arguments[0].scrollIntoView()", control);
                ((IJavaScriptExecutor)WebDriverManager.GetWebdriver()).ExecuteScript("window.scrollBy(0,-100)", control);

                switch (control.TagName.ToLower())
                {
                    case "kendo-dropdownlist":
                        {
                            control.Click();
                            Result = SelectDropDownKendoDropDown(WebDriverManager.GetWebdriver(), dropDownValues);
                            break;
                        }

                    case "kendo-multiselect":
                        {
                            control.Click();
                            Result = SelectDropDownKendoDropDownNewMulti(WebDriverManager.GetWebdriver(), control, dropDownValues);
                            break;
                        }

                    case "kendo-combobox":
                        {
                            control.Click();
                            Result = SelectKendoCombobox(WebDriverManager.GetWebdriver(), control, dropDownValues);
                            break;
                        }

                    default:
                        {
                            Result = control.SelectDropDown(dropDownValues);
                            break;
                        }

                }


                //LogMessage($"Select Drop Down Item: {uiElement}  : {dropDownValues}");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed Select Drop Down Item: {uiElement} : {dropDownValues} \n Exception Details: {e.Message}");
            }

            return Result;
        }

        private static bool SelectDropDownKendoDropDownNewMulti(this IWebDriver BrowserWindow, IWebElement element, string[] dropDownValues)
        {
            IList<IWebElement> options = null;

            bool result = false;

            try
            {
                for (int i = 0; i < 10; i++)
                {
                    try
                    {
                        Thread.Sleep(1300);
                        options = BrowserWindow.FindElements(By.ClassName("k-item"));
                        //IList<IWebElement> e = BrowserWindow.FindElements(By.XPath("//ul[@role='listbox']/li"));
                        if (options.Count <= 0)
                        {
                            throw new Exception();
                        }
                        IEnumerator<IWebElement> iterator = options.GetEnumerator();
                        foreach (string ddlValue in dropDownValues)
                        {
                            result = false;
                            while (iterator.MoveNext())
                            {
                                if (iterator.Current.Text.Equals(ddlValue))
                                {

                                    iterator.Current.Click();
                                    result = true;
                                    break;
                                    //if (BrowserWindow.FindElement(By.XPath("(//kendo-multiselect[@id='" + element.GetAttribute("id") + "']//span)[1]")).Text.Equals(ddlValue))
                                    //{
                                    //    break;
                                    //}
                                    //else
                                    //{
                                    //    throw new Exception();
                                    //}

                                }
                            }
                            break;
                        }
                        break;
                    }


                    catch (Exception e)
                    {
                        element.Click();
                        result = false;
                    }
                }




            }
            catch (Exception e)
            {
                //element.Click();
            }




            return result;
        }

        public static List<string> ReadDropDownItem(string uiElement, string valueFromExcel)
        {
            List<string> result = new List<string>();
            IList<IWebElement> options = null;
            IList<IWebElement> finalOptions = null;
            try
            {
                var control = SetFocus(uiElement, valueFromExcel);
                if (control.TagName.ToLower().Contains("kendo-dropdownlist"))
                {
                    control.Click();
                    Thread.Sleep(1000);
                    IList<IWebElement> ddlValues = WebDriverManager.GetWebdriver().FindElements(By.XPath("//kendo-popup//li"));
                    IReadOnlyCollection<IWebElement> options1 = control.FindElements(By.ClassName("k-input"));
                    finalOptions = new List<IWebElement>();
                    foreach (IWebElement wElement in options1)
                    {
                        finalOptions.Add(wElement);
                    }

                    if (ddlValues.Count == 0)
                    {
                        options = WebDriverManager.GetWebdriver().FindElements(By.ClassName("k-item"));

                        foreach (IWebElement wElement in options)
                        {
                            finalOptions.Add(wElement);
                        }
                    }
                    else
                    {
                        foreach (IWebElement wElement in ddlValues)
                        {
                            finalOptions.Add(wElement);
                        }
                    }
                }
                else if (control.TagName.ToLower().Contains("my-dropdown-filter"))
                {
                    control.Click();
                    Thread.Sleep(1000);
                    options = WebDriverManager.GetWebdriver().FindElements(By.ClassName("k-item"));
                    finalOptions = options.ToList<IWebElement>();
                }
                else
                {
                    SelectElement dropdown = new SelectElement(control);
                    finalOptions = new List<IWebElement>();
                    finalOptions = dropdown.Options;
                }
                foreach (IWebElement option in finalOptions)
                {
                    result.Add(option.Text.Trim());
                }
                //LogMessage($"Captured list of dropdown items from element : {uiElement} \n total count: {result.Count}");
            }
            catch (Exception e)
            {
                result = null;
                MessageLogger.LogMessage($"Failed find the list of Drop Down Item: {uiElement} \n Exception Details: {e.Message}");
            }

            return result;
        }

        public static string ReadSelectedDropDownValue(string uiElement, string valueFromExcel)
        {
            string result = string.Empty;

            try
            {
                var control = SetFocus(uiElement, valueFromExcel);
                SelectElement dropdown = new SelectElement(control);
                result = dropdown.SelectedOption.Text;
                MessageLogger.LogMessage($"Captured list of dropdown items from element : {uiElement}");
            }
            catch (Exception e)
            {
                result = null;
                MessageLogger.LogMessage($"Failed find the list of Drop Down Item: {uiElement} \n Exception Details: {e.Message}");
            }

            return result;
        }

        private static bool SelectDropDown(this IWebElement control, string[] dropDownValues)
        {
            bool result = false;
            Thread.Sleep(1000);
            SelectElement dropdown = new SelectElement(control);
            if (!dropdown.IsMultiple)
            {

                dropdown.SelectByText(dropDownValues[0]);
                result = true;
            }
            else
            {
                foreach (IWebElement option in dropdown.Options)
                {
                    result = false;
                    if (dropDownValues.Contains(option.Text.Trim()))
                    {
                        option.Click();
                        result = true;
                        if (dropDownValues.Length == 1)
                        {
                            break;
                        }
                    }
                }
            }

            return result;
        }

        private static bool SelectDropDownKendoDropDown(this IWebDriver BrowserWindow, string[] dropDownValues)
        {
            bool result = false;

            Thread.Sleep(1000);

            foreach (string ddlValue in dropDownValues)
            {
                IList<IWebElement> options = BrowserWindow.FindElements(By.XPath("//kendo-popup//li[text()[normalize-space()='" + ddlValue + "']]"));

                if (options.Count == 0)
                {
                    options = BrowserWindow.FindElements(By.XPath("//*[(@class='k-item') or (@class='k-item k-state-focused') or (@class='k-item k-state-focused k-state-selected')]"));

                    IEnumerator<IWebElement> iterator = options.GetEnumerator();


                    while (iterator.MoveNext())
                    {
                        if (iterator.Current.Text.Equals(ddlValue))
                        {
                            if (iterator.Current.Displayed && iterator.Current.Enabled)
                            {
                                try
                                {
                                    iterator.Current.Click();
                                    result = true;
                                    break;
                                }
                                catch (Exception e)
                                {

                                }
                            }
                        }

                    }
                    break;

                }
                else
                {
                    options[0].Click();
                    result = true;
                }
            }

            //IList<IWebElement> options = BrowserWindow.FindElements(By.XPath("//*[(@class='k-item') or (@class='k-item k-state-focused') or (@class='k-item k-state-focused k-state-selected')]"));

            //IEnumerator<IWebElement> iterator = options.GetEnumerator();

            //foreach (string ddlValue in dropDownValues)
            //{
            //    while (iterator.MoveNext())
            //    {
            //        if (iterator.Current.Text.Equals(ddlValue))
            //        {
            //            if (iterator.Current.Displayed && iterator.Current.Enabled)
            //            {
            //                try
            //                {
            //                    iterator.Current.Click();
            //                    result = true;
            //                    break;
            //                }
            //                catch (Exception e)
            //                {

            //                }
            //            }
            //        }

            //    }
            //    break;
            //}
            if (!result)
            {
                IList<IWebElement> options = BrowserWindow.FindElements(By.XPath("//*[(@class='k-item') or (@class='k-item k-state-focused') or (@class='k-item k-state-focused k-state-selected')]"));

                IEnumerator<IWebElement> iterator = options.GetEnumerator();
                iterator = options.GetEnumerator();
                foreach (string ddlValue in dropDownValues)
                {
                    while (iterator.MoveNext())
                    {
                        if (iterator.Current.Text.Contains(ddlValue))
                        {
                            iterator.Current.Click();
                            result = true;
                            break;
                        }
                    }
                    break;
                }
            }

            return result;
        }

        private static bool SelectKendoCombobox(this IWebDriver BrowserWindow, IWebElement element, string[] dropDownValues)
        {
            bool result = false;

            Thread.Sleep(1000);

            IList<IWebElement> dropdownArrows = element.FindElements(By.XPath("//*[@class='k-i-arrow-s k-icon']")).Where(e => e.Enabled && e.Displayed).ToList();
            dropdownArrows[0].Click();

            IList<IWebElement> options = BrowserWindow.FindElements(By.XPath("//ul[@class='k-list k-reset']//li"));

            IEnumerator<IWebElement> iterator = options.GetEnumerator();

            foreach (string ddlValue in dropDownValues)
            {
                options.Where(s => s.Text.Equals(dropDownValues[0])).ToList()[options.Where(s => s.Text.Equals(dropDownValues[0])).ToList().Count - 1].Click();
                result = true;
            }

            //foreach (string ddlValue in dropDownValues)
            //{
            //    while (iterator.MoveNext())
            //    {
            //        if (iterator.Current.Text.Equals(ddlValue))
            //        {
            //            iterator.Current.Click();
            //            result = true;
            //            break;
            //        }
            //    }
            //    break;
            //}
            if (!result)
            {
                iterator = options.GetEnumerator();
                foreach (string ddlValue in dropDownValues)
                {
                    while (iterator.MoveNext())
                    {
                        if (iterator.Current.Text.Contains(ddlValue))
                        {
                            iterator.Current.Click();
                            result = true;
                            break;
                        }
                    }
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// Method for verifying if property of any control with the value
        /// </summary>
        /// <param name="Object"></param>
        /// <param name="BrowserWindow"></param>
        /// <param name="TestContext"></param>
        /// <returns></returns>
        public static bool VerifyProperty(string Object, string Value)
        {

            bool Result = false;
            string[] property = Value.Split(',');
            try
            {
                Value = VariableManager.GetVariableValue(Value);
                IWebElement control = FindObject(Object, WebDriverManager.GetWebdriver(), Value);

                if (control.GetAttribute(property[property.Length - 2]) != null)
                {
                    Result = (property[property.Length - 1].ToLower().Trim() == control.GetAttribute(property[property.Length - 2]).ToLower().Trim());
                }
                else
                {
                    if (property[property.Length - 2].ToLower().Contains("enabled"))
                    {
                        Result = (property[property.Length - 1].ToLower().Trim() == control.Enabled.ToString().ToLower().Trim());
                    }
                    else if (property[property.Length - 2].ToLower().Contains("selected"))
                    {
                        Result = (property[property.Length - 1].ToLower().Trim() == control.Selected.ToString().ToLower().Trim());
                    }
                    else if (property[property.Length - 2].ToLower().Contains("text"))
                    {
                        Result = (property[property.Length - 1].ToLower().Trim() == control.Text.ToString().ToLower().Trim());
                    }
                    else if (property[property.Length - 2].ToLower().Contains("displayed"))
                    {
                        Result = (property[property.Length - 1].ToLower().Trim() == control.Displayed.ToString().ToLower().Trim());
                    }
                    else
                    {
                        throw new Exception("Property " + property[property.Length - 2].ToLower() + "doesn't exist for control" + Object);
                    }
                }

                // Print the result
                if (Result)
                {
                    MessageLogger.LogMessage("VerifyProperty : For the control: " + Object + " " + property[0] + " is  " + property[1]);
                }

                else
                {
                    MessageLogger.LogMessage("Failed VerifyProperty : For the control: " + Object + " " + property[0] + " is not " + property[1]);
                }

            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage("Failed VerifyProperty : For the control: " + Object + " " + property[0] + " " + property[1]);
                MessageLogger.LogMessage("Failed VerifyProperty : For the control: " + Object + " " + property[0] + " " + property[1] + " Exception Details: " + e.Message);
            }

            return Result;
        }

        /// <summary>
        /// Method for verifying if property of any control with the value
        /// </summary>
        /// <param name="Object"></param>
        /// <param name="BrowserWindow"></param>
        /// <param name="TestContext"></param>
        /// <returns></returns>
        public static bool VerifyCSSProperty(string Object, string Value)
        {

            bool Result = false;
            string[] property = Value.Split('|');
            try
            {
                IWebElement control = FindObject(Object, WebDriverManager.GetWebdriver(), Value);

                if (control.GetCssValue(property[property.Length - 2]) != null)
                {
                    // string s = control.GetAttribute("class");

                    Result = control.GetCssValue(property[property.Length - 2].ToLower().Trim()).Contains((property[property.Length - 1].ToLower().Trim()));
                }
                else
                {
                    throw new Exception("Property " + property[property.Length - 2].ToLower() + "doesn't exist for control" + Object);
                }


                // Print the result
                if (Result)
                {
                    MessageLogger.LogMessage("VerifyProperty : For the control: " + Object + " " + property[0] + " is  " + property[1]);
                }
                else
                {
                    MessageLogger.LogMessage("Failed VerifyProperty : For the control: " + Object + " " + property[0] + " is not " + property[1]);
                }
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage("Failed VerifyProperty : For the control: " + Object + " " + property[0] + " " + property[1]);
                MessageLogger.LogMessage("Failed VerifyProperty : For the control: " + Object + " " + property[0] + " " + property[1] + " Exception Details: " + e.Message);
            }

            return Result;
        }

        public static bool VerifyText(string uiElement, string uiElementValue)
        {
            bool Result;
            try
            {
                var uielementNew = SetFocus(uiElement, uiElementValue).Text;
                Result = uielementNew.Contains(uiElementValue);
                if (Result)
                {
                    MessageLogger.LogMessage("Verified innertext of the control: " + uiElement + " : " + uiElementValue);
                }
                else
                {
                    MessageLogger.LogMessage("Failed to verify innertext of the control: " + uiElement + " : " + uiElementValue);
                }
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage("Failed to verify innertext of the control: " + uiElement + " : " + uiElementValue);
                MessageLogger.LogMessage("Failed to verify innertext of the control:" + uiElement + " : " + uiElementValue + " Exception Details: " + e.Message);
            }
            return Result;
        }

        public static bool VerifyTextOnPopUp(string uiElement, string uiElementValue)
        {
            bool Result;
            try
            {
                var uielementNew = SetFocus(uiElement, uiElementValue).Text;
                //var windows = BrowserWindow.WindowHandles;
                //var y = BrowserWindow.SwitchTo().Window(windows[0]);
                //var temp = y.SetFocus(uiElement);                
                //var uielementNew = temp.Text;
                //var uielementNew = BrowserWindow.SwitchTo().ActiveElement().FindChildObject(uiElement).Text;
                //LogMessage($"Expected message: {uiElementValue}");
                //LogMessage($"Actual message: {uielementNew}");
                Result = uielementNew.ToLower().Contains(uiElementValue.ToLower());
                if (Result)
                {
                    MessageLogger.LogMessage("Verified innertext of the control: " + uiElement + " : " + uiElementValue);
                }
                else
                {
                    MessageLogger.LogMessage("Failed to verify innertext of the control: " + uiElement + " : " + uiElementValue);

                }
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage("Failed to verify innertext of the control: " + uiElement + " : " + uiElementValue);
                MessageLogger.LogMessage("Failed to verify innertext of the control:" + uiElement + " : " + uiElementValue + " Exception Details: " + e.Message);
            }
            return Result;
        }

        /// <summary>
        /// Method for opening the menus
        /// </summary>
        /// <param name="Object"></param>
        /// <param name="BrowserWindow"></param>
        /// <param name="TestContext"></param>
        /// <returns></returns>
        public static bool HoverMenu(string Object, IWebDriver BrowserWindow, string valueFromExcel)
        {

            bool Result = false;
            try
            {
                IWebElement Link = null;
                //It is for handling the Divs
                if (Object.ToLower().Contains("masterscheduling"))
                {
                    Link = FindObject(Object, BrowserWindow, valueFromExcel);
                }
                //This is for normal clickable menu
                else
                {
                    Link = FindObject(Object, BrowserWindow, valueFromExcel);
                }

                /**/
                String mouseOverScript = "if(document.createEvent){var evObj = document.createEvent('MouseEvents');evObj.initEvent('mouseover', true, false); "
                    + "arguments[0].dispatchEvent(evObj);} else if(document.createEventObject) { arguments[0].fireEvent('onmouseover');}";
                IJavaScriptExecutor js = BrowserWindow as IJavaScriptExecutor;
                js.ExecuteScript(mouseOverScript, Link);

                Result = true;
                //LogMessage("Open Menu: " + Object);
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage("Failed Open Menu: " + Object);
                MessageLogger.LogMessage("Failed Open Menu: " + Object + " Exception Details: " + e.Message);
            }

            return Result;
        }

        /// <summary>
        /// Drags and drops the input Object(and value) to target location
        /// </summary>
        /// <param name="Object"></param>
        /// <param name="Value"></param>
        /// <param name="BrowserWindow"></param>
        /// <param name="TestContext"></param>
        /// <returns></returns>
        public static bool DragAndDrop(string Object, string Value)
        {
            bool Result = false;

            try
            {
                IWebElement sourceControl = FindObject(Object.Split(',')[0], WebDriverManager.GetWebdriver(), Value);
                IWebElement targetControl = FindObject(Object.Split(',')[1], WebDriverManager.GetWebdriver(), Value);

                if (Value.ToLower() != "na")
                {
                    //sourceControl.SearchProperties.Add(new PropertyExpression(HtmlControl.PropertyNames.InnerText, Value));
                    sourceControl = WebDriverManager.GetWebdriver().FindElement(By.LinkText(Value));
                }

                Actions action = new Actions(WebDriverManager.GetWebdriver());
                action.DragAndDrop(sourceControl, targetControl).Build().Perform();

                Result = true;
                //LogMessage("DragAndDrop: " + Object + " : " + Value);
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage("Failed DragAndDrop: " + Object + " : " + Value);
                MessageLogger.LogMessage("Failed DragAndDrop: " + Object + " : " + Value + "Exception Details: " + e.Message);
            }

            return Result;
        }

        public static DataTable ReadTableGrid(string uiElement, string uiElementValue)
        {
            string[] value = GetObject(uiElement).Split(ConfigHelper.ObjectPropertyValueSeparator);
            string tableXPath = value[1];
            string Property = value[0];
            DataTable dt = new DataTable();
            SetFocus(uiElement, uiElementValue);
            try
            {
                var xx = tableXPath + "/div/div/table/thead/tr/th";
                var yy = tableXPath + "/kendo-grid-list/div/table/tbody/tr";
                //Get number of columns In table.     
                var headerElements = WebDriverManager.GetWebdriver().FindElements(By.XPath(xx));
                int columnCount = WebDriverManager.GetWebdriver().FindElements(By.XPath(tableXPath + "/div/div/table/thead/tr/th")).Count();

                IEnumerable<IWebElement> controls = WebDriverManager.GetWebdriver().FindElements(By.XPath(tableXPath + "/div/div/table/thead/tr/th"));
                int count = 0;
                IWebElement control = null;
                foreach (IWebElement item in controls)
                {
                    if (item.Displayed)
                    {
                        control = item;
                        count++;
                    }
                }

                columnCount = count;

                //get number of rows in table
                int rowCount = WebDriverManager.GetWebdriver().FindElements(By.XPath(tableXPath + "/kendo-grid-list/div/table/tbody/tr")).Count();

                controls = WebDriverManager.GetWebdriver().FindElements(By.XPath(tableXPath + "/kendo-grid-list/div/table/tbody/tr"));
                count = 0;

                foreach (IWebElement item in controls)
                {
                    if (item.Displayed)
                    {
                        control = item;
                        count++;
                    }
                }

                rowCount = count;

                //divide xpath in three parts to pass rowCount and columnCount values
                string firstPart = tableXPath + "/kendo-grid-list/div/table/tbody/tr[";
                string secondPart = "]/td[";
                string thirdPart = "]";

                //Loop to read column headers
                for (int i = 1; i <= columnCount; i++)
                {
                    string columnHeader = WebDriverManager.GetWebdriver().FindElement(By.XPath(tableXPath + "/div/div/table/thead/tr/th[" + i + "]")).Text;
                    dt.Columns.Add(columnHeader);
                }

                //Used for loop for number of rows.
                for (int i = 1; i <= rowCount; i++)
                {
                    DataRow dr = dt.NewRow();
                    //Used for loop for number of columns.
                    for (int j = 1; j <= columnCount; j++)
                    {
                        //Prepared final xpath of specific cell as per values of i and j
                        string finalPath = firstPart + i + secondPart + j + thirdPart;
                        string tableData = WebDriverManager.GetWebdriver().FindElement(By.XPath(finalPath)).Text;
                        dr[j - 1] = tableData;
                    }
                    dt.Rows.Add(dr);
                }
            }
            catch (Exception exception)
            {
                MessageLogger.LogMessage("Failed to fetch table data for the object: " + uiElement);
                MessageLogger.LogMessage("Exception message: " + exception.Message);
                MessageLogger.LogMessage("Failed to fetch table data for the object: " + uiElement + ", Exception Details: " + exception.Message);
            }
            return dt;
        }

        public static int ReadTableGridRowCount(string uiElement, string uiElementValue)
        {
            string[] value = GetObject(uiElement).Split(ConfigHelper.ObjectPropertyValueSeparator);
            SetFocus(uiElement, uiElementValue);
            string tableXPath = value[1];
            string Property = value[0];
            int rowCount = 0;
            try
            {
                //get number of rows in table
                rowCount = WebDriverManager.GetWebdriver().FindElements(By.XPath(tableXPath + "/kendo-grid-list/div/table/tbody/tr")).Count();
                //LogMessage($"Row count: {rowCount}");
            }
            catch (Exception exception)
            {
                MessageLogger.LogMessage($"Failed to fetch table data for the object: {uiElement} \n Exception Details: {exception.Message}");
            }
            return rowCount;
        }
        public static int ReadTableGridColumnCount(string uiElement, string uiElementValue)
        {
            string[] value = GetObject(uiElement).Split(ConfigHelper.ObjectPropertyValueSeparator);
            string tableXPath = value[1];
            string Property = value[0];
            int columnCount = 0;
            try
            {
                //Get number of columns In table.                     
                columnCount = WebDriverManager.GetWebdriver().FindElements(By.XPath(tableXPath + "/div/div/table/thead/tr/th")).Count();
                //LogMessage($"Column count: {columnCount}");
            }
            catch (Exception exception)
            {
                MessageLogger.LogMessage($"Failed to fetch table data for the object: {uiElement} \n Exception Details: {exception.Message}");
            }
            return columnCount;
        }

        public static bool CaptureScreenshot(string screenshotName, string screenpath)
        {
            try
            {
                Screenshot screenshot = ((ITakesScreenshot)WebDriverManager.GetWebdriver()).GetScreenshot();
                //string screenshotName = "FailedItemScreenShot_" + System.DateTime.Now.ToString("yyyy-mmm-ddd-hh-mm-ss") + ".png";
                //string screenpath = Path.Combine(@"c:\screenshots", screenshotName);
                screenshot.SaveAsFile(screenpath, ScreenshotImageFormat.Jpeg);
                PageLabelTableManager.GetTestContext().AddResultFile(screenpath);
                return true;
            }
            catch (Exception e)
            {
                MessageLogger.LogMessage("Failed to take screenshot");
                MessageLogger.LogMessage("Exception message: " + e.Message);
                return false;
            }

        }


        public static string ReadControlInnerTextByJs(string Object, string valueFromExcel)
        {
            string innerText = "";
            try
            {
                //innerText = GetValByJs(BrowserWindow,SetFocus(Object));
                innerText = SetFocus(Object, valueFromExcel).GetAttribute("value");
                //LogMessage($"Control inner text: {Object} : {innerText}");
            }
            catch (Exception e)
            {
                MessageLogger.LogMessage($"Failed to find control: { Object} : \n Exception Details: {e.Message}");
                innerText = "NF";
            }
            return innerText;
        }

        public static bool VerifyIfControlExist(string ControlIdentifier, string valueFromExcel)
        {
            bool Result = false;
            IWebElement Control = null;
            try
            {

                Control = FindObject(ControlIdentifier, WebDriverManager.GetWebdriver(), valueFromExcel);
                if (Control == null)
                {
                    Result = false;
                }
                else
                {
                    Result = true;
                }
                //LogMessage($"VerifyIfControlExist: {ControlIdentifier}");

            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed VerifyIfControlExist: {ControlIdentifier} \n Exception Details: {e.Message}");
            }
            return Result;

        }

        public static bool VerifyIfControlHasSomeText(string ControlIdentifier, string valueFromExcel)
        {
            bool Result = false;
            IWebElement Control = null;
            try
            {

                Control = FindObject(ControlIdentifier, WebDriverManager.GetWebdriver(), valueFromExcel);
                if (Control == null)
                {
                    Result = false;
                }
                else if (ReadControlInnerText(ControlIdentifier, valueFromExcel).Equals(String.Empty) && ReadControlInnerTextByJs(ControlIdentifier, valueFromExcel).Equals(String.Empty))
                {
                    Result = false;
                }
                else
                {
                    Result = true;
                }
                //LogMessage($"VerifyIfControlHasSomeText: {ControlIdentifier}");

            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed VerifyIfControlHasSomeText: {ControlIdentifier} \n Exception Details: {e.Message}");
            }
            return Result;
        }

        public static bool VerifyIfControlNotHasSomeText(string ControlIdentifier, string valueFromExcel)
        {
            bool Result = false;
            IWebElement Control = null;
            try
            {

                Control = FindObject(ControlIdentifier, WebDriverManager.GetWebdriver(), valueFromExcel);
                if (Control == null)
                {
                    Result = false;
                }
                else if (ReadControlInnerText(ControlIdentifier, valueFromExcel).Equals(String.Empty) && ReadControlInnerTextByJs(ControlIdentifier, valueFromExcel).Equals(String.Empty))
                {
                    Result = false;
                }
                else
                {
                    Result = !ReadControlInnerText(ControlIdentifier, valueFromExcel).Equals(valueFromExcel) && !ReadControlInnerTextByJs(ControlIdentifier, valueFromExcel).Equals(valueFromExcel);
                }
                MessageLogger.LogMessage($"VerifyIfControlNotHasSomeText: {ControlIdentifier}");

            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed VerifyIfControlNotHasSomeText: {ControlIdentifier} \n Exception Details: {e.Message}");
            }
            return Result;
        }

        public static bool VerifyIfControlNotExist(string ControlIdentifier, string valueFromExcel)
        {
            bool Result = false;
            IWebElement Control = null ;
            try
            {
                Control = FindObject(ControlIdentifier, WebDriverManager.GetWebdriver(), valueFromExcel);
                if (Control == null)
                {
                    Result = true;
                }
                else
                {
                    Result = false;
                }
                //LogMessage($"VerifyIfControlExist: {ControlIdentifier}");

            }
            catch (NoSuchElementException e)
            {
                Result = true;
                MessageLogger.LogMessage($"VerifyIfControlNotExist: {ControlIdentifier}");
            }
            catch (WebDriverTimeoutException e)
            {
                Result = true;
                MessageLogger.LogMessage($"VerifyIfControlNotExist: {ControlIdentifier}");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed VerifyIfControlExist: {ControlIdentifier} \n Exception Details: {e.Message}");
            }
            return Result;

        }

        /// <summary>
        /// Compares inputlist with valuefromexcel
        /// </summary>
        /// <param name="BrowserWindow"></param>
        /// <param name="inputList"></param>
        /// <param name="valueFromExcel"></param>
        /// <returns></returns>
        public static bool CompareListsContains(List<string> inputList, string valueFromExcel)
        {
            bool Result = false;
            //IWebElement Control = null;
            try
            {
                Result = inputList.Contains(valueFromExcel);

                //LogMessage($"CompareListsContains");
            }

            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed CompareListsContains \n Exception Details: {e.Message}");
            }

            return Result;

        }

        /// <summary>
        /// Method to check current login user
        /// </summary>
        /// <param name="BrowserWindow"></param>
        /// <param name="ControlIdentifier"></param>
        /// <param name="valueFromExcel"></param>
        /// <returns></returns>
        public static bool VerifyCurrentUser(string ControlIdentifier, string valueFromExcel)
        {
            bool Result = false;
            //IWebElement Control = null;
            IWebElement Control = null;
            //string currentUser = "";
            try
            {
                try
                {
                    // Control = FindObject(ControlIdentifier, BrowserWindow, valueFromExcel);
                    Control = WebDriverManager.GetWebdriver().FindElement(By.XPath(GetObject("UserNameOnHomePage").Split('|')[1]));
                }
                catch (Exception)
                {
                    Control = null;
                }
                if (Control == null)
                {
                    Result = false;
                }
                else if (!Control.Text.ToLower().Equals(valueFromExcel.ToLower()))
                {
                    FindObject(ControlIdentifier, WebDriverManager.GetWebdriver(), valueFromExcel).Click();
                    FindObject("LogOutButtonOnHomePage", WebDriverManager.GetWebdriver(), valueFromExcel).Click();
                    Thread.Sleep(2500);
                    if (PageLabelTableManager.GetTestContext().Properties["URL"].ToString() == "BO")
                    {
                        NavigateToUrl(ConfigHelper.BOURL);
                    }
                    else
                    {
                        NavigateToUrl(ConfigHelper.CFOURL);
                    }
                    Result = false;
                }
                else
                {
                    Result = true;
                }
                ////LogMessage($"VerifyCurrentUser: {ControlIdentifier}");
                Thread.Sleep(2500);
                try
                {
                    IWebElement useA = WebDriverManager.GetWebdriver().FindElement(By.Id("use_another_account"));
                    if (useA.Displayed || useA.Enabled)
                    {
                        useA.Click();
                    }
                }
                catch (Exception e)
                {
                    Console.Write("Failed to find the use another account link");
                }
            }
            catch (NoSuchElementException e)
            {
                Result = true;
                MessageLogger.LogMessage($"VerifyCurrentUser: {ControlIdentifier}");
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed VerifyCurrentUser: {ControlIdentifier} \n Exception Details: {e.Message}");
            }
            return Result;

        }

        /// <summary>
        /// Method to check current login user
        /// </summary>
        /// <param name="BrowserWindow"></param>
        /// <param name="ControlIdentifier"></param>
        /// <param name="valueFromExcel"></param>
        /// <returns></returns>
        public static bool VerifyCheckBoxChecked(string ControlIdentifier, string valueFromExcel)
        {
            bool Result = false;
            //IWebElement Control = null;
            IWebElement Control = null;
            try
            {
                Control = FindObject(ControlIdentifier, WebDriverManager.GetWebdriver(), valueFromExcel);
                bool outcome = Control.Selected;
                if (valueFromExcel.ToLower().Equals("na") || valueFromExcel.ToLower().Equals("true"))
                {
                    if (outcome)
                    {
                        Result = true;
                    }
                    else
                    {
                        Result = false;
                    }
                }
                else if (valueFromExcel.ToLower().Equals("false"))
                {
                    if (!outcome)
                    {
                        Result = true;
                    }
                    else
                    {
                        Result = false;
                    }
                }
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage($"Failed VerifyCheckBoxChecked : {ControlIdentifier} \n Exception Details: {e.Message}");

            }
            //string currentUser = "";

            return Result;

        }

        /// <summary>
        /// Swiches frame with input parameters
        /// </summary>
        /// <param name="BrowserWindow"></param>
        /// <param name="uiElementIdentifier"></param>
        /// <param name="uiElementValue"></param>
        /// <returns></returns>
        public static bool SwitchToFrame(string uiElementIdentifier, string uiElementValue)
        {
            bool result = false;
            try
            {
                string frameCase = uiElementIdentifier.ToLower().Trim();
                switch (frameCase)
                {
                    case "index":
                        WebDriverManager.GetWebdriver().SwitchTo().Frame(Convert.ToInt16(uiElementValue));
                        result = true;
                        break;
                    case "name":
                        WebDriverManager.GetWebdriver().SwitchTo().Frame(uiElementValue);
                        result = true;
                        break;
                    case "element":
                        WebDriverManager.GetWebdriver().SwitchTo().Frame(FindObject(uiElementIdentifier, WebDriverManager.GetWebdriver(), uiElementIdentifier));
                        result = true;
                        break;
                }

            }
            catch (Exception e)
            {
                MessageLogger.LogMessage($"Failed to switch frame : {uiElementIdentifier} \n Exception Details: {e.Message}");
                result = false;
            }

            return result;

        }

        /// <summary>
        /// Swiches window to defaut context
        /// </summary>
        /// <param name="BrowserWindow"></param>
        /// <param name="uiElementIdentifier"></param>
        /// <param name="uiElementValue"></param>
        /// <returns></returns>
        public static bool SwitchToDefaultContext(string uiElementIdentifier, string uiElementValue)
        {
            bool result = false;
            try
            {
                WebDriverManager.GetWebdriver().SwitchTo().DefaultContent();
                result = true;
            }
            catch (Exception e)
            {
                //LogMessage($"Failed to switch window to default context: {e.Message}");
                result = false;
            }

            return result;

        }


        /// <summary>
        /// Method to click element if it is visible
        /// </summary>
        /// <param name="BrowserWindow"></param>
        /// <param name="ControlIdentifier"></param>
        /// <param name="valueFromExcel"></param>
        /// <returns></returns>
        public static bool ClickIfVisible(string ControlIdentifier, string valueFromExcel)
        {
            bool Result = false;
            //IWebElement Control = null;
            IWebElement Control = null;
            try
            {
                try
                {
                    Control = FindObject(ControlIdentifier, WebDriverManager.GetWebdriver(), valueFromExcel);

                    if (Control.Displayed || Control.Enabled)
                    {
                        Control = FindObject(valueFromExcel, WebDriverManager.GetWebdriver(), ControlIdentifier);
                        Control.Click();
                        Result = true;
                    }
                }
                catch (Exception e)
                {
                    Result = true;
                }
            }
            catch (Exception e)
            {
                Result = false;
                //LogMessage($"Failed Click if Visible : {ControlIdentifier} \n Exception Details: {e.Message}");

            }
            //string currentUser = "";

            return Result;

        }


        #endregion

       

        private static string GetObject(string uiElementValue) => Convert.ToString(ConfigHelper.ControlObjects[uiElementValue]);

        
    }
}
