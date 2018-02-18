using Atom.SeleniumCore.WebDriverManagers;
using Atom.Utilities;
using Atom.Utilities.Reports;
using Atom.Utilities.SqlDataManager;
using Atom.Utilities.TestDataExcel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using StackExchange.Redis;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Atom.KeyWordsManagers
{
    class KeyWordsFactory
    {
        static Dictionary<string, Func<string, string, bool>> KeyMethodPairs = new Dictionary<string, Func<string, string, bool>>();

        public void BuildKeyMethodPairs()
        {
            AddKeyMethodPair("entertext", EnterText)
               .AddKeyMethodPair("entertextonpopup", EnterTextOnPopUp)
               .AddKeyMethodPair("clickcontrol", ClickControl)
               .AddKeyMethodPair("clickcontroljs", ClickControlJs)
               .AddKeyMethodPair("clickcontrolonpopup", ClickControlOnPopUp)
               .AddKeyMethodPair("readinnertext", ReadControlInnerText)
               .AddKeyMethodPair("verifylabel", VerifyLabel)
               .AddKeyMethodPair("verifylabelonpopup", VerifyLabelOnPopUp)
               .AddKeyMethodPair("hoveroncontrol", HoverOnControl)
               .AddKeyMethodPair("enterspecialkeys", EnterSpecialKeys)
               .AddKeyMethodPair("opendropdown", OpenDropDown)
               .AddKeyMethodPair("selectdropdownitem", SelectDropDownItem)
               .AddKeyMethodPair("readdropdown", ReadDropDownItems)
               .AddKeyMethodPair("readtablegrid", ReadTableGrid)
               .AddKeyMethodPair("executenonscalar", ExecuteNonScalar)
               .AddKeyMethodPair("dismisspopupwindow", DismissPopupWindow)
               .AddKeyMethodPair("acceptpopupwindow", AcceptPopupWindow)
               .AddKeyMethodPair("verifymessageonpopupwindow", VerifyMessageOnPopupWindow)
               .AddKeyMethodPair("executejs", ExecuteJS)
               .AddKeyMethodPair("comparelabels", CompareLabels)
               .AddKeyMethodPair("comparelist", CompareLists)
               .AddKeyMethodPair("comparelabelscontains", CompareLabelsContains)
               .AddKeyMethodPair("comparetables", CompareTables)
               .AddKeyMethodPair("executequery", ExecuteQuery)
               .AddKeyMethodPair("executescalar", ExecuteScalar)
               .AddKeyMethodPair("draganddrop", DragAndDrop)
               .AddKeyMethodPair("excuteexcel", ExecuteTestCase)
               .AddKeyMethodPair("verifyproperty", VerifyProperty)
               .AddKeyMethodPair("verifycssproperty", VerifyCSSProperty)
               .AddKeyMethodPair("waitforprocess", WaitForProcess)
               .AddKeyMethodPair("readtablegridcolumncount", ReadTableGridColumnCount)
               .AddKeyMethodPair("readtablegridrowcount", ReadTableGridRowCount)
               .AddKeyMethodPair("readtablegridandverifyrecordcount", ReadTableGridAndVerifyRecordCount)
               .AddKeyMethodPair("refreshpage", RefreshPage)
               .AddKeyMethodPair("navigatetourl", NavigateToUrl)
               .AddKeyMethodPair("readcontrolinnertextbyjs", ReadControlInnerTextByJs)
               .AddKeyMethodPair("verifyifcontrolexist", VerifyIfControlExist)
               .AddKeyMethodPair("verifyifcontrolhassometext", VerifyIfControlHasSomeText)
               .AddKeyMethodPair("verifyifcontrolnothassometext", VerifyIfControlNotHasSomeText)
               .AddKeyMethodPair("navigatetoapplicationurl", NavigateToApplicationUrl)
               .AddKeyMethodPair("genrandomstring", GenRandomString)
               .AddKeyMethodPair("pagesourcecontains", PageSourceContains)
               .AddKeyMethodPair("clickcontrolwithtext", ClickControlWithText)
               .AddKeyMethodPair("setfocus", SetFocus)
               .AddKeyMethodPair("opennewbrowser", OpenNewBrowser)
             .AddKeyMethodPair("readselectedropdownvalue", ReadSelectedDropDownValue)
            .AddKeyMethodPair("unclickcontrolwithtext", UnClickControlWithText)
            .AddKeyMethodPair("setvariable", SetVariable)
            .AddKeyMethodPair("getoptionsonautosugegst", GetOptionsOnAutoSuggest)
            .AddKeyMethodPair("verifyifcontrolnotexist", VerifyIfControlNotExist)
             .AddKeyMethodPair("switchtochildwindow", SwitchToChildWindow)
            .AddKeyMethodPair("verifycurrentloginuser", VerifyCurrentUser)
            .AddKeyMethodPair("comparelistscontains", CompareListsContains)
            .AddKeyMethodPair("verifycheckboxchecked", VerifyCheckBoxChecked)
            .AddKeyMethodPair("comparefirstlistcontainsvaluesfromsecondlist", CompareFirstListContainsValuesFromSecondList)
            .AddKeyMethodPair("clearrediscache", ClearRedisCache)
            .AddKeyMethodPair("gettextonlistofcontrols", GetTextOnListOfControls)
            .AddKeyMethodPair("switchtoframe", SwitchToFrame)
            .AddKeyMethodPair("switchtodefaultcontext", SwitchToDefaultContext)
            .AddKeyMethodPair("clickifvisible", ClickIfVisible);



            //test.Add("entertext", EnterText);
            //test.Add("clickcontrol", ClickControl);
        }



        private bool WaitForProcess(string timeInMilliSeconds = "MediumWait", string comments = "")
        {

            if (timeInMilliSeconds.ToLower() == "minimumwait")
            {
                Wait(ConfigHelper.MinimumWait);
                MessageLogger.LogMessage($"Waiting for the time: {ConfigHelper.MinimumWait} milliseconds");
            }
            else if (timeInMilliSeconds.ToLower() == "mediumwait")
            {
                Wait(ConfigHelper.MediumWait);
                MessageLogger.LogMessage($"Waiting for the time: {ConfigHelper.MediumWait} milliseconds");
            }
            else if (timeInMilliSeconds.ToLower() == "maximumwait")
            {
                Wait(ConfigHelper.MaximumWait);
                MessageLogger.LogMessage($"Waiting for the time: {ConfigHelper.MaximumWait} milliseconds");
            }
            else if (timeInMilliSeconds.ToLower() == "explicitwait")
            {
                Wait(ConfigHelper.ExplicitWait);
                MessageLogger.LogMessage($"Waiting for the time: {ConfigHelper.ExplicitWait} milliseconds");
            }
            else
            {
                var time = Convert.ToInt16(timeInMilliSeconds);
                MessageLogger.LogMessage($"Waiting for the time: {time} milliseconds");
                Wait(time);
            }
            return true;
        }


        public KeyWordsFactory AddKeyMethodPair(string key, Func<string, string, bool> method)
        {
            KeyMethodPairs.Add(key, method);
            return this;
        }


        #region Execute Test case 
        /// <summary>
        /// Method for running test cases directly without master plan file
        /// </summary>
        /// <param name="TestScenarioFileNameWithPath"></param>
        /// <param name="BrowserWindow"></param>
        /// <returns></returns>
        public static bool ExecuteTestCase(string TestScenarioFileNameWithPath, string moduleName = "")
        {
            try
            {
                string enviornmentDetails = string.Empty;
                bool Result = false;

                DataSet TestCaseDS = GetExcelSheetData("", moduleName, TestScenarioFileNameWithPath);
                // HF.LogMessage($"Number tables : {TestCaseDS.Tables.Count}");
                DataTable TestScenario = TestCaseDS.Tables["MasterTestPlan"];
                int TestCasePassCount = 0;

                MessageLogger.LogMessage("*************** Test Case Steps ***************");
                //OpenQA.Selenium.Support.UI.WebDriverWait wait = new OpenQA.Selenium.Support.UI.WebDriverWait(HF.BrowserWindow, TimeSpan.FromSeconds(60));
                foreach (DataRow row in TestScenario.Rows)
                {
                    string Keyword = row[0].ToString().ToLower().Trim();
                    string uiElement = row[1].ToString().Trim();
                    string uiElementValue = row[2].ToString().Trim();

                    MessageLogger.LogMessage($"------executing excel step : {TestCasePassCount + 1}------");

                    if (KeyMethodPairs.ContainsKey(Keyword.ToLower()))
                    {
                        //for (int i = 0; i < 10; i++)
                        //{
                        //try
                        //{
                        WebDriverWait wait = new WebDriverWait(WebDriverManager.GetWebdriver(), TimeSpan.FromSeconds(120)); wait.Until(wd => { try { return (wd as IJavaScriptExecutor).ExecuteScript("return (document.readyState == 'complete' && jQuery.active == 0)"); } catch { return false; } });

                        Wait(ConfigHelper.MinimumWait);
                        Result = KeyMethodPairs[Keyword.ToLower()](uiElement, uiElementValue);
                        if (Keyword.ToLower().Equals("verifycurrentloginuser") && Result)
                        {
                            Result = true;
                            break;
                        }
                        else if (Keyword.ToLower().Equals("verifycurrentloginuser") && !Result)
                        {
                            Result = true;
                        }
                        //    if (!Result)
                        //    { 
                        //        throw new Exception();
                        //    }
                        //    break;
                        //}
                        //catch
                        //{ 
                        //   wait.Until(OpenQA.Selenium.Support.UI.ExpectedConditions.ElementExists(By.XPath("//div[@class='shell-preloader'][(contains(@style, 'display: none'))]")));
                        //     wait.Until(OpenQA.Selenium.Support.UI.ExpectedConditions.ElementExists(By.XPath("//div[@id='loaderNav'][(contains(@style, 'display: none'))]")));
                        //  i++;
                        //}
                        // }

                    }
                    else
                    {
                        MessageLogger.LogMessage("Failed at Keyword: " + Keyword + " for Object " + uiElement + "Value: Null or Empty");
                        Result = false;
                    }

                    if (Result)
                        TestCasePassCount++;
                    else
                    {
                        //This will take a screenshot of the failed item and save the data in test results if the configuration is enabled
                        if (ConfigHelper.IsScreenCaptureEnabledOnStepFailure && (!Keyword.Equals("excuteexcel")))
                        {
                            string datetime = DateTime.Now.ToString("MM-dd-yyyy-hh-mm-ss");
                            string screenshotName = $"{PageLabelTableManager.GetTestContext().Properties["TFSID"].ToString()}_{datetime}.jpg";
                            string screenpath = FolderManager.createScreenshotFolderRunFolder() + "\\" + screenshotName;
                            KeyWordsMethods.CaptureScreenshot(screenshotName, screenpath);
                            PageLabelTableManager.GetTestContext().Properties.Add("ScreenShot", screenshotName);
                            MessageLogger.LogMessage("**************Failed at Keyword: " + Keyword + " for Object " + uiElement + "Value: " + uiElementValue + "************");
                        }

                        return false;
                    }
                }
                return Result;
            }

            catch (Exception exception)
            {
                //BaseTestSuite.captureScreenShotOnFailure();
                MessageLogger.LogMessage($"Failed with exception: {exception.Message}");
                MessageLogger.LogMessage($"Exception Stack trace: {exception.StackTrace}");
                return false;
            }
        }

        #endregion

        #region Excute Test case child methods

        #region Excel methods

        /// <summary>
        /// Gets Excel sheet data from respective input modulename (or) SharedSteps (or) SharedSteps(Common folder)
        /// </summary>
        /// <param name="moduleName"></param>
        /// <param name="value"></param>
        /// <param name="excelSheetName"></param>
        /// <returns></returns>
        public static DataSet GetExcelSheetData(string moduleName, string value, string excelSheetName)
        {

            DataSet TestCaseDS = null;
            TestCaseReader testCaseReaderObj = new TestCaseReader();
            TestCaseDS = testCaseReaderObj.GetExcelSheetData(moduleName, value, excelSheetName);

            return TestCaseDS;
        }


        #endregion // Excel Methods

        #region UT test methods

        public static void Wait(int timeInMilliSeconds)
        {
            Thread.Sleep(TimeSpan.FromMilliseconds(timeInMilliSeconds));
        }
        
        public IWebDriver LaunchAndNavigate(string url)
        {
            return KeyWordsMethods.LaunchAndNavigate(ConfigHelper.BrowserType, url);
        }

        private bool DragAndDrop(string uiElement, string uiElementValue)
        {
            return KeyWordsMethods.DragAndDrop(uiElement, uiElementValue);
        }
        public bool EnterText(string uiElementIdentifier, string uiElementValue)
        {
            if (uiElementValue != null)
            {
                uiElementValue = ReadPageLabels(uiElementValue);
                return KeyWordsMethods.EnterText(uiElementIdentifier, uiElementValue);
            }
            else
            {
                MessageLogger.LogMessage("Failed. Invalid Value for Enter Text. Value is NULL.");
                return false;
            }
        }

        private bool EnterTextOnPopUp(string uiElementIdentifier, string uiElementValue)
        {
            if (uiElementValue != null)
            {
                uiElementValue = ReadPageLabels(uiElementValue);
                return KeyWordsMethods.EnterTextOnPopUp(uiElementIdentifier, uiElementValue);
            }
            else
            {
                MessageLogger.LogMessage("Failed. Invalid Value for Enter Text. Value is NULL.");
                return false;
            }
        }

        private bool EnterSpecialKeys(string uiElementIdentifier, string uiElementValue)
        {
            if (uiElementValue != null)
            {
                return KeyWordsMethods.EnterSpecialKeys(uiElementIdentifier, uiElementValue);
            }
            else
            {
                MessageLogger.LogMessage("Failed to enter special keys: " + uiElementValue);
                return false;
            }
        }
        private bool ClickControl(string uiElementIdentifier, string uiElementValue)
        {
            //bool clickControlSucceeded = HF.BrowserWindow.ClickControl(uiElementIdentifier, uiElementValue);
            //// replacing clickcontrol with clickcontroljs to support IE browser
            //if (!clickControlSucceeded)
            //{
            //clickControlSucceeded = HF.BrowserWindow.ClickControlJs(uiElementIdentifier, uiElementValue);
            // }

            bool clickControlSucceeded = KeyWordsMethods.ClickControlJs(uiElementIdentifier, uiElementValue);

            return clickControlSucceeded;
        }

        public bool ClickControlJs(string uiElementIdentifier, string uiElementValue)
        {
            return KeyWordsMethods.ClickControlJs(uiElementIdentifier, uiElementValue);
        }

        private bool ClickControlOnPopUp(string uiElementIdentifier, string uiElementValue)
        {
            return KeyWordsMethods.ClickControlOnPopUp(uiElementIdentifier, uiElementValue);
        }

        private bool NavigateToApplicationUrl(string uiElement, string uiElementValue)
        {
            uiElementValue = ReadPageLabels(uiElementValue);

            if (PageLabelTableManager.GetTestContext().Properties["URL"].ToString() == string.Empty)
            {
                NavigateToUrl("", uiElementValue);
            }
            else
            {

                if (PageLabelTableManager.GetTestContext().Properties["URL"].ToString() == "BO")
                    KeyWordsMethods.NavigateToUrl(ConfigHelper.BOURL + uiElementValue);
                else if (PageLabelTableManager.GetTestContext().Properties["URL"].ToString() == "CFO")
                    KeyWordsMethods.NavigateToUrl(ConfigHelper.CFOURL + uiElementValue);
            }

            //Thread.Sleep(4000);
            return true;
        }

        private bool GenRandomString(string uiElement, string uiElementValue)
        {
            string currentTimeStamp = DateTime.Now.ToString("ssffff");
            while (currentTimeStamp.StartsWith("0"))
            {
                currentTimeStamp = DateTime.Now.ToString("ssffff");
            }
            this.UpdatePageLabels(uiElementValue, uiElementValue.Replace('{', ' ').Replace('}', ' ').Trim() + currentTimeStamp);
            return true;
        }

        /// <summary>
        /// Method for extract auto suggest list items from auto suggest text box
        /// </summary>
        /// <param name="uiElementIdentifier"></param>
        /// <param name="uiElementValue"></param>
        /// <returns></returns>
        private bool GetOptionsOnAutoSuggest(string uiElementIdentifier, string uiElementValue)
        {
            if (uiElementValue != null)
            {

                try
                {
                    IList<IWebElement> allOptions =
                        WebDriverManager.GetWebdriver().FindElements(By.XPath(uiElementIdentifier));
                    List<String> allText = new List<string>();
                    foreach (var a in allOptions)
                    {
                        allText.Add(a.Text);
                    }

                    UpdatePageLists(uiElementValue, allText);
                    return true;
                }
                catch (Exception e)
                {
                    MessageLogger.LogMessage("Failed. Invalid Value for Enter Text. Value is NULL. " +e.Message);
                    return false;
                }

            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Method to extract text on list items 
        /// </summary>
        /// <param name="uiElementIdentifier"></param>
        /// <param name="uiElementValue"></param>
        /// <returns></returns>
        private bool GetTextOnListOfControls(string uiElementIdentifier, string uiElementValue)
        {
            try
            {
                uiElementValue = ReadPageLabels(uiElementValue);
                UpdatePageLists(uiElementValue, KeyWordsMethods.GetTextFromListOfControls(uiElementIdentifier, uiElementValue));
                return true;

            }
            catch (Exception e)
            {
                MessageLogger.LogMessage("Failed. Invalid Value for Enter Text. Value is NULL. " + e.Message);
                return false;
            }

        }

        /// <summary>
        /// Swiches frame with input parameters
        /// </summary>
        /// <param name="uiElementIdentifier"></param>
        /// <param name="uiElementValue"></param>
        /// <returns></returns>
        private bool SwitchToFrame(string uiElementIdentifier, string uiElementValue)
        {
            return KeyWordsMethods.SwitchToFrame(uiElementIdentifier, uiElementValue);
        }

        /// <summary>
        /// Swiches window to default context
        /// </summary>
        /// <param name="uiElementIdentifier"></param>
        /// <param name="uiElementValue"></param>
        /// <returns></returns>
        private bool SwitchToDefaultContext(string uiElementIdentifier, string uiElementValue)
        {
            return KeyWordsMethods.SwitchToDefaultContext(uiElementIdentifier, uiElementValue);
        }


        private bool SwitchToChildWindow(string uiElementIdentifier, string uiElementValue)
        {
            uiElementValue = ReadPageLabels(uiElementValue);
            return KeyWordsMethods.SwitchToChildWindow(uiElementIdentifier, uiElementValue);
        }



        private bool PageSourceContains(string uiElement, string uiElementValue)
        {
            bool pageSourceContains = (uiElement.ToLower() == "true" || uiElement.ToLower() == "na" || String.IsNullOrEmpty(uiElement.ToLower())) ? true : false;

            bool result = false;

            int i = 1;
            while (uiElementValue.Contains('}') || uiElementValue.Contains('{'))
            {
                int startIndex = uiElementValue.IndexOf('{');
                int endIndex = uiElementValue.IndexOf('}');

                string dynamicPart = uiElementValue.Substring(startIndex, endIndex - uiElementValue.IndexOf('{') + 1);
                string dynamicPartValue = ReadPageLabels(dynamicPart);
                uiElementValue = uiElementValue.Replace(dynamicPart, dynamicPartValue);




                //if (uiElementValue.Contains('}'))
                //{
                //    if (uiElementValue.IndexOf('{') == 0)
                //    {
                //        string staticValue = uiElementValue.Substring(uiElementValue.IndexOf('}') + 1);

                //        string dynamicValue = uiElementValue.Substring(uiElementValue.IndexOf('{'), uiElementValue.IndexOf('}') - uiElementValue.IndexOf('{') + 1);

                //        dynamicValue = ReadPageLabels(dynamicValue);

                //        uiElementValue = dynamicValue + staticValue;
                //    }
                //    else
                //    {
                //        string staticValue = uiElementValue.Substring(0, uiElementValue.IndexOf('{'));

                //        string dynamicValue = uiElementValue.Substring(uiElementValue.IndexOf('{'), uiElementValue.IndexOf('}') - uiElementValue.IndexOf('{') + 1);

                //        dynamicValue = ReadPageLabels(dynamicValue);

                //        uiElementValue = staticValue + dynamicValue;
                //    }
                //}

            }

            while (!result && i < 30)
            {
                if (pageSourceContains)
                {
                    result = WebDriverManager.GetWebdriver().PageSource.ToLower().Contains(uiElementValue.ToLower());
                }
                else
                {
                    result = !WebDriverManager.GetWebdriver().PageSource.ToLower().Contains(uiElementValue.ToLower());
                }

                i++;
                Thread.Sleep(500);
            }


            return result;
        }

        private bool ClickControlWithText(string uiElement, string uiElementValue)
        {
            uiElementValue = ReadPageLabels(uiElementValue);
            return KeyWordsMethods.ClickControlWithText(uiElement, uiElementValue);
        }

        private bool UnClickControlWithText(string uiElement, string uiElementValue)
        {
            uiElementValue = ReadPageLabels(uiElementValue);
            return KeyWordsMethods.UnClickControlWithText(uiElement, uiElementValue);
        }

        /// <summary>
        /// Sets variable from uiElementValue into uiElement
        /// </summary>
        /// <param name="uiElement"></param>
        /// <param name="uiElementValue"></param>
        /// <returns></returns>
        private bool SetVariable(string uiElement, string uiElementValue)
        {
            uiElementValue = ReadPageLabels(uiElementValue);
            UpdatePageLabels(uiElement, uiElementValue);
            return true;
        }

        private bool VerifyIfControlNotExist(string uiElement, string uiElementValue)
        {
            return KeyWordsMethods.VerifyIfControlNotExist(uiElement, uiElementValue);
        }

        private bool VerifyCurrentUser(string uiElement, string uiElementValue)
        {
            uiElement = ReadPageLabels(uiElement);
            uiElementValue = ReadPageLabels(uiElementValue);
            return KeyWordsMethods.VerifyCurrentUser(uiElement, uiElementValue);
        }

        private bool CompareListsContains(string inputList, string uiElementValue)
        {
            string[] values = inputList.Split(ConfigHelper.ObjectPropertyValueSeparator);
            List<string> label1;
            string label2;

            label1 = this.ReadPageLists(values[0]);

            label2 = this.ReadPageLabels(values[1]);

            bool flag = false;

            if (uiElementValue.ToLower() == "true" || uiElementValue.ToLower() == "na")
            {
                flag = KeyWordsMethods.CompareListsContains(label1, label2);
            }
            else
            {
                flag = !KeyWordsMethods.CompareListsContains(label1, label2);
            }

            return flag;
        }


        private bool ReadControlInnerText(string uiElement, string uiElementValue)
        {
            string innerText = KeyWordsMethods.ReadControlInnerText(uiElement, uiElementValue);
            bool result = true;
            if (innerText == "NF")
            {
                MessageLogger.LogMessage("Failed. unable to find control");
                result = false;
            }
            else
            {
                UpdatePageLabels(uiElementValue, innerText);
            }
            return result;
        }



        private bool VerifyProperty(string uiElement, string uiElementValue)
        {
            return KeyWordsMethods.VerifyProperty(uiElement, uiElementValue);
        }
        private bool VerifyCSSProperty(string uiElement, string uiElementValue)
        {
            return KeyWordsMethods.VerifyCSSProperty(uiElement, uiElementValue);
        }

        private bool ExecuteJS(string script, string uiElementValue) => KeyWordsMethods.ExecuteJS(script, uiElementValue);


        private bool SelectDropDownItem(string uiElement, string uiElementValue)
        {
            string[] dropdownValues = uiElementValue.Split(ConfigHelper.ObjectPropertyValueSeparator);
            dropdownValues = ReadPageLabels(dropdownValues);
            return KeyWordsMethods.SelectDropDownItem(uiElement, dropdownValues);
        }

        public bool ReadDropDownItems(string uiElement, string uiElementValue)
        {
            UpdatePageLists(uiElementValue, KeyWordsMethods.ReadDropDownItem(uiElement, uiElementValue));
            return true;
        }

        public bool ReadSelectedDropDownValue(string uiElement, string uiElementValue)
        {
            UpdatePageLabels(uiElementValue, KeyWordsMethods.ReadSelectedDropDownValue(uiElement, uiElementValue));
            return true;
        }

        private bool OpenDropDown(string uiElement, string uiElementValue) => KeyWordsMethods.OpenDropDown(uiElement, uiElementValue);

        private bool VerifyLabel(string uiElement, string uiElementValue)
        {
            bool Result = false;
            if (uiElementValue != null)
            {
                uiElementValue = ReadPageLabels(uiElementValue);
                Result = KeyWordsMethods.VerifyText(uiElement, uiElementValue);
            }

            else
            {
                MessageLogger.LogMessage("Failed. Invalid Value for Verify Label. Value is NULL.");
                Result = false;
            }
            return Result;
        }

        private bool VerifyLabelOnPopUp(string uiElement, string uiElementValue)
        {
            bool Result = false;
            if (uiElementValue != null)
            {
                uiElementValue = ReadPageLabels(uiElementValue);
                Result = KeyWordsMethods.VerifyTextOnPopUp(uiElement, uiElementValue);
            }

            else
            {
                MessageLogger.LogMessage("Failed. Invalid Value for Verify Label. Value is NULL.");
                Result = false;
            }
            return Result;
        }

        private bool ReadTableGrid(string uiElement, string uiElementValue)
        {
            DataTable dt = KeyWordsMethods.ReadTableGrid(uiElement, uiElementValue);
            this.UpdatePageTables(uiElementValue, dt);
            return true;
        }

        private bool SetFocus(string uiElement, string uiElementValue)
        {
            if (KeyWordsMethods.SetFocus(uiElement, uiElementValue) == null)
            {
                return false;
            }

            return true;

        }

        private bool OpenNewBrowser(string uiElement, string uiElementValue)
        {
            IWebDriver driver = null;
            bool result = true;
            WebDriverManager.GetWebdriver().Close();

            if (uiElement.Trim().ToLower() == "newsession")
            {
                WebDriverFactory webDriverFactoryObj = new WebDriverFactory();
                driver = webDriverFactoryObj.CreateDriverInstance(ConfigHelper.BrowserType);
                WebDriverManager.SetDriver(driver);
                uiElementValue = ReadPageLabels(uiElementValue);
                PageLabelTableManager.GetTestContext().Properties["URL"] = string.Empty;
                LaunchAndNavigate(uiElementValue);
            }
            else
            {
                result = ExecuteJS("window.open()", string.Empty);
                WebDriverManager.GetWebdriver().SwitchTo().Window(string.Empty);
                WebDriverManager.GetWebdriver().SwitchTo().Window(WebDriverManager.GetWebdriver().WindowHandles.Last());
            }

            return result;
        }

        private bool ReadTableGridRowCount(string uiElement, string uiElementValue)
        {
            int count = KeyWordsMethods.ReadTableGridRowCount(uiElement, uiElementValue);
            this.UpdatePageLabels(uiElementValue, count.ToString());
            return true;
        }

        private bool ReadTableGridColumnCount(string uiElement, string uiElementValue)
        {
            int count = KeyWordsMethods.ReadTableGridColumnCount(uiElement, uiElementValue);
            this.UpdatePageLabels(uiElementValue, count.ToString());
            return true;
        }

        private bool HoverOnControl(string uiElement, string uiElementValue) => KeyWordsMethods.HoverOnControl(uiElement, uiElementValue);

        private bool DismissPopupWindow(string uiElement, string uiElementValue) => KeyWordsMethods.DismissPopupWindow();

        private bool AcceptPopupWindow(string uiElement, string uiElementValue) => KeyWordsMethods.AcceptPopupWindow();

        private bool VerifyMessageOnPopupWindow(string uiElementValue, string temp) => KeyWordsMethods.VerifyMessageOnPopupWindow(uiElementValue);


        private bool CompareLabelsContains(string value, string flagResult)
        {

            string[] values = value.Split(ConfigHelper.ObjectPropertyValueSeparator);
            var dblist = values[0].Replace('{', ' ').Replace('}', ' ').Trim();
            var uilist = values[1].Replace('{', ' ').Replace('}', ' ').Trim();
            MessageLogger.LogMessage($"Comparing lists: {uilist} , {uilist}");
            var pageResults = this.ReadPageLabels(values[0]);
            var dbResults = this.ReadPageLabels(values[1]);
            bool flag = (flagResult.ToLower() == "true") ? true : false;
            //HF.LogMessage($"Total count of :{dblist} is {pageResults.Count} ,\n Total count of :{uilist} is {dbResults.Count}");
            //bool result = (pageResults.Count().Equals(dbResults.Count())) == flag ? true : false;
            //if (result == flag && result == true)
            //{
            pageResults = pageResults.ToLower();
            dbResults = dbResults.ToLower();
            if ((flagResult.ToLower() == "true") ? dbResults.ToLower().Contains(pageResults.ToLower()) : !dbResults.ToLower().Contains(pageResults.ToLower()))
            {
                MessageLogger.LogMessage($"Total count from both lists are matching");
                return true;
            }

            return false;
        }

        private bool ReadTableGridAndVerifyRecordCount(string uiElement, string uiElementValue)
        {
            DataTable dt = KeyWordsMethods.ReadTableGrid(uiElement, uiElementValue);
            this.UpdatePageTables(uiElementValue, dt);
            if (dt.Rows.Count > 0)
            {
                return true;
            }
            else
                return false;
        }

        private bool RefreshPage(string uiElement, string uiElementValue)
        {
            WebDriverManager.GetWebdriver().Navigate().Refresh();
            Thread.Sleep(4000);
            return true;
        }

        private bool NavigateToUrl(string uiElement, string uiElementValue)
        {
            uiElementValue = ReadPageLabels(uiElementValue);
            KeyWordsMethods.NavigateToUrl(uiElementValue);
            //Thread.Sleep(4000);
            return true;
        }

        private bool ReadControlInnerTextByJs(string uiElement, string uiElementValue)
        {
            string innerText = KeyWordsMethods.ReadControlInnerTextByJs(uiElement, uiElementValue);
            bool result = true;
            if (innerText == "NF")
            {
                MessageLogger.LogMessage("Failed. unable to find control");
                result = false;
            }
            else
            {
                UpdatePageLabels(uiElementValue, innerText);
            }
            return result;
        }

        private bool VerifyIfControlExist(string uiElementIdentifier, string uiElementValue)
        {
            uiElementValue = ReadPageLabels(uiElementValue);
            return KeyWordsMethods.VerifyIfControlExist(uiElementIdentifier, uiElementValue);
        }

        private bool VerifyIfControlHasSomeText(string uiElementIdentifier, string uiElementValue)
        {
            uiElementValue = ReadPageLabels(uiElementValue);
            return KeyWordsMethods.VerifyIfControlHasSomeText(uiElementIdentifier, uiElementValue);
        }

        private bool VerifyIfControlNotHasSomeText(string uiElementIdentifier, string uiElementValue)
        {
            uiElementValue = ReadPageLabels(uiElementValue);
            return KeyWordsMethods.VerifyIfControlNotHasSomeText(uiElementIdentifier, uiElementValue);
        }
        #endregion

        #region Database Methods

        /// <summary>
        /// Executes query based on parameters
        /// </summary>
        /// <param name="sqlFileName">sql file name</param>
        /// <param name="value">Contains parameters</param>
        /// <returns>bool</returns>
        public bool ExecuteQuery(string sqlFileName, string value)
        {
            bool Result = false;
            string[] values = value.Split(ConfigHelper.ObjectPropertyValueSeparator);
            string parameters = "";
            if (values.Count() > 1)
                parameters = values[1];
            try
            {
                string queryToExecute = GetFormattedQuery(sqlFileName, value);
                SQLDataProvider sQLDataProviderObj = new SQLDataProvider();
                int effectedRowsCount = sQLDataProviderObj.ExecuteNonQuery(ConfigHelper.ConnectionString, queryToExecute, PageLabelTableManager.GetTestContext());
                //Console.WriteLine("\n Query for this file:" + sqlFileName + " is " + queryToExecute + "\n");
                MessageLogger.LogMessage($"Executed required query: {sqlFileName}: effected row count: {effectedRowsCount}");
                Result = true;
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage("Failed to execute query => " + sqlFileName + "\n Exception Details: " + e.Message);
            }

            return Result;
        }

        /// <summary>
        /// Gives single value as result and will get saved in pagelables dictionary to use further
        /// </summary>
        /// <param name="sqlFileName">Sql file name</param>
        /// <param name="value">{keyname}|value1,value2,value3</param>
        /// <returns>bool</returns>
        private bool ExecuteScalar(string sqlFileName, string value)
        {
            bool Result = false;
            string[] values = value.Split(ConfigHelper.ObjectPropertyValueSeparator);
            string parameters = "";
            if (values.Count() > 1)
                parameters = values[1];
            string key = values[0];
            try
            {
                string queryToExecute = GetFormattedQuery(sqlFileName, parameters);
                SQLDataProvider sQLDataProviderObj = new SQLDataProvider();
                string output = sQLDataProviderObj.ExecuteScalar(ConfigHelper.ConnectionString, queryToExecute, PageLabelTableManager.GetTestContext());
                //Console.WriteLine("\n Query for this file:" + sqlFileName + " is " + queryToExecute + "\n");
                if (output != null && output != "")
                {
                    Result = true;
                    // HF.LogMessage("Executed required query" + sqlFileName);
                    // HF.LogMessage("Executed required query result" + output);
                    this.UpdatePageLabels(key, output);
                }
                else
                {
                    Result = false;
                    MessageLogger.LogMessage("Failed to execute query =>" + sqlFileName);
                }
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage("Failed to execute query => " + sqlFileName + "Exception Details: " + e.Message);
            }

            return Result;
        }

        /// <summary>
        /// Gives data table as result after executing the query and will get saved in PageTables dictionary to use further
        /// </summary>
        /// <param name="sqlFileName">Sql file name</param>
        /// <param name="value">{keyname}|value1,value2,value3</param>
        /// <returns>bool</returns>
        public bool ExecuteNonScalar(string sqlFileName, string value)
        {
            bool Result = false;
            string[] values = value.Split(ConfigHelper.ObjectPropertyValueSeparator);
            string parameters = "";
            if (values.Count() > 1)
                parameters = values[1];
            string key = values[0];
            try
            {
                string queryToExecute = GetFormattedQuery(sqlFileName, parameters);
                SQLDataProvider sQLDataProviderObj = new SQLDataProvider();
                DataTable output = sQLDataProviderObj.ExecuteNonScalar(ConfigHelper.ConnectionString, queryToExecute, PageLabelTableManager.GetTestContext());
                //Console.WriteLine("\n Query for this file:" + sqlFileName + " is " + queryToExecute + "\n");
                if (output != null && output.Rows.Count > 0)
                {
                    Result = true;
                    // HF.LogMessage("Executed required query" + sqlFileName);
                    // HF.LogMessage("Executed required query result" + output);
                    if (output.Columns.Count == 1)
                    {
                        var outputList = this.ConvertTableToList(output);
                        // HF.LogMessage($"Total list of records fetched: {outputList.Count}");
                        this.UpdatePageLists(key, outputList);
                    }

                    else
                        this.UpdatePageTables(key, output);
                }
                else
                {
                    Result = false;
                    // HF.LogMessage("Failed to execute query =>" + sqlFileName);
                }
            }
            catch (Exception e)
            {
                Result = false;
                MessageLogger.LogMessage("Failed to execute query => " + sqlFileName + "Exception Details: " + e.Message);
            }

            return Result;
        }



        private List<string> ConvertTableToList(DataTable output)
        {
            List<string> list = new List<string>();
            foreach (DataRow dr in output.Rows)
            {
                list.Add(dr[0].ToString().Trim());
            }
            return list;
        }

        /// <summary>
        /// Compare tables based on key values provided 
        /// </summary>
        /// <param name="value">{SQLDatatableKey}|{UIDatatableKey}</param>
        /// <param name="comments">comments</param>
        /// <returns>bool</returns>
        private bool CompareTables(string value, string comments)
        {

            string[] values = value.Split(ConfigHelper.ObjectPropertyValueSeparator);
            DataTable SQLresults = this.ReadPageTables(values[1]);
            DataTable UIResults = this.ReadPageTables(values[0]);
            bool Result = false;
            Console.WriteLine("====================Starting comparing SQL data and Search Result data...=========================");
            int MatchingResultCount = 0;
            if ((SQLresults.Rows.Count == UIResults.Rows.Count) && (SQLresults.Columns.Count == UIResults.Columns.Count))
            {
                foreach (DataRow SQLRow in SQLresults.Rows)
                {
                    foreach (DataRow UIRow in UIResults.Rows)
                    {
                        //var SQLData = SQLRow.ItemArray;
                        var SQLData = SQLRow.ItemArray.Select(i => i == null ? string.Empty : i.ToString().Trim().ToLower()).ToArray();

                        //var UIDate = UIRow.ItemArray;
                        var UIDate = UIRow.ItemArray.Select(i => i == null ? string.Empty : i.ToString().Trim().ToLower()).ToArray();

                        if (SQLData.SequenceEqual(UIDate))
                        {
                            Console.WriteLine("\nData from SQL Query: ");
                            for (int SqlItem = 0; SqlItem < SQLData.Length; SqlItem++)
                                Console.WriteLine(SQLData[SqlItem].ToString());

                            Console.WriteLine("\nData from UI : ");
                            for (int SqlItem = 0; SqlItem < SQLData.Length; SqlItem++)
                                Console.WriteLine(UIDate[SqlItem].ToString());

                            MatchingResultCount++;
                            Console.WriteLine("Found the Required Data in Results!");
                            break;
                        }
                    }
                }
            }

            if (UIResults.Rows.Count == MatchingResultCount)
                Result = true;
            else
            {
                Result = false;
                Console.WriteLine("Not Found all Required Data in Results!");
            }
            Console.WriteLine("======================Finished the Comparision==================!");
            return Result;
        }

        private bool CompareLists(string value, string flagResult)
        {

            string[] values = value.Split(ConfigHelper.ObjectPropertyValueSeparator);
            var dblist = values[0].Replace('{', ' ').Replace('}', ' ').Trim();
            var uilist = values[1].Replace('{', ' ').Replace('}', ' ').Trim();
            MessageLogger.LogMessage($"Comparing lists: {uilist} , {uilist}");
            var pageResults = this.ReadPageLists(values[0]);
            var dbResults = this.ReadPageLists(values[1]);
            bool flag = (flagResult.ToLower() == "true") ? true : false;
            MessageLogger.LogMessage($"Total count of :{dblist} is {pageResults.Count} ,\n Total count of :{uilist} is {dbResults.Count}");
            bool result = (pageResults.Count().Equals(dbResults.Count())) == flag ? true : false;
            if (result == flag && result == true)
            {
                pageResults = pageResults.ConvertAll(d => d.ToLower());
                dbResults = dbResults.ConvertAll(d => d.ToLower());
                var firstNotSecond = pageResults.Except(dbResults).ToList();
                var secondNotFirst = dbResults.Except(pageResults).ToList();
                result = !firstNotSecond.Any() && !secondNotFirst.Any();
                MessageLogger.LogMessage($"Total count from both lists are matching");
            }

            return result;
        }
        /// <summary>
        /// Tests whether the values in the second dropdown list are present in the first dropdown list
        /// </summary>
        /// <param name="value"></param>
        /// <param name="flagResult"></param>
        /// <returns></returns>
        private bool CompareFirstListContainsValuesFromSecondList(string value, string flagResult)
        {

            string[] values = value.Split(ConfigHelper.ObjectPropertyValueSeparator);
            var dblist = values[0].Replace('{', ' ').Replace('}', ' ').Trim();
            var uilist = values[1].Replace('{', ' ').Replace('}', ' ').Trim();
            MessageLogger.LogMessage($"Comparing lists: {uilist} , {uilist}");
            var pageResults = this.ReadPageLists(values[0]);
            var dbResults = this.ReadPageLists(values[1]);
            bool flag = (flagResult.ToLower() == "true") ? true : false;
            MessageLogger.LogMessage($"Total count of :{dblist} is {pageResults.Count} ,\n Total count of :{uilist} is {dbResults.Count}");
            bool result = (pageResults.Count().Equals(dbResults.Count())) == flag ? true : false;
            if (result == true && flag == true) // checking that first list and second list both are same
            {
                pageResults = pageResults.ConvertAll(d => d.ToLower());
                dbResults = dbResults.ConvertAll(d => d.ToLower());
                var firstNotSecond = pageResults.Except(dbResults).ToList();
                var secondNotFirst = dbResults.Except(pageResults).ToList();
                result = !firstNotSecond.Any() && !secondNotFirst.Any();
                MessageLogger.LogMessage($"Total count from both lists are matching");
            }
            else if (((result == true) || (result == false)) && flag == false) // checking that first list doesnt contain any value from second list
            {
                pageResults = pageResults.ConvertAll(d => d.ToLower());
                var count = pageResults.Count;
                dbResults = dbResults.ConvertAll(d => d.ToLower());
                var firstNotSecond = pageResults.Except(dbResults).ToList();
                var secondNotFirst = dbResults.Except(pageResults).ToList();
                result = firstNotSecond.Count == count; // first list doesnt contain any element from second list, hence count should not change

            }

            else if (result == false && flag == true) // checking that first list contains some elements from second list
            {
                pageResults = pageResults.ConvertAll(d => d.ToLower());
                dbResults = dbResults.ConvertAll(d => d.ToLower());
                var secondNotFirst = dbResults.Except(pageResults).ToList();
                result = !secondNotFirst.Any();
            }

            return result;
        }


        /// <summary>
        /// Clears redis cache with input connection string
        /// </summary>
        /// <param name="environment"></param>
        /// <returns></returns>
        private bool ClearRedisCache(string s, string s1)
        {
            try
            {
                var redisConnectionString = String.Format("{0},abortConnect=false,ssl=true,ConnectTimeout=30000,SyncTimeout=30000,password={1},allowAdmin=true",
                                                      PageLabelTableManager.GetTestContext().Properties["RedisCacheServer_Name"].ToString(),
                                                     PageLabelTableManager.GetTestContext().Properties["RedisCacheServer_Password"].ToString());

                var _connectionMultiplexer = ConnectionMultiplexer.Connect(redisConnectionString);
                var endpoints = _connectionMultiplexer.GetEndPoints(true);
                foreach (var endpoint in endpoints)
                {
                    var server = _connectionMultiplexer.GetServer(endpoint);
                    server.FlushAllDatabases();
                }
            }

            catch (Exception ex)
            {
                MessageLogger.LogMessage("Failed to clear Redis cache" + ex.Message);
                return false;
            }

            return true;
        }


        private bool CompareLabels(string value, string flagResult)
        {
            if (flagResult.ToLower().Equals("na"))
            {
                flagResult = "true";
            }
            string[] values = value.Split(ConfigHelper.ObjectPropertyValueSeparator);
            string label1, label2;
            if (values[0].ToString().Contains("{") && values[0].ToString().Contains("}"))
            {
                label1 = values[0].Replace('{', ' ').Replace('}', ' ').Trim();
                label1 = this.ReadPageLabels(values[0]);
            }
            else
            {
                label1 = values[0].ToString();
            }
            if (values[1].ToString().Contains("{") && values[1].ToString().Contains("}"))
            {
                label2 = values[1].Replace('{', ' ').Replace('}', ' ').Trim();
                label2 = this.ReadPageLabels(values[1]);
            }
            else
            {
                label2 = values[1].ToString();
            }
            MessageLogger.LogMessage($"Comparing labels: {label1} ; {label2}");
            bool result = label1.TrimEnd().Equals(label2.TrimEnd());

            if (result.ToString().ToLower().TrimEnd().Equals(flagResult.ToLower().TrimEnd()))
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        /// <summary>
        /// generate n digit random number
        /// </summary>
        /// <param name="digCount"></param>
        /// <returns></returns>
        public static String GetRandomNumber(int digCount)
        {
            CommonMethods commonUtilitiesObj = new CommonMethods();
            return commonUtilitiesObj.GetRandomNumber(digCount);
        }

        /// <summary>
        /// random between 2 nos
        /// </summary>
        /// <param name="digCount"></param>
        /// <returns></returns>
        public String RandomNumberBetween2Nos(int low, int high)
        {
            Random r = new Random();
            int Low = low;
            int High = high;
            int Result = r.Next(High - Low) + Low;
            return Result.ToString();
        }


        #endregion

        #region Saving key, value pairs 
        private void UpdatePageTables(string key, DataTable value)
        {
            key = key.Replace("{", "").Replace("}", "");
            PageLabelTableManager.GetPageTables().Add(key, value);
        }

        private void UpdatePageLabels(string key, string value)
        {
            key = key.Replace("{", "").Replace("}", "");

            if (!PageLabelTableManager.GetPageLabels().ContainsKey(key))
            {
                PageLabelTableManager.GetPageLabels().Add(key, value);
            }
            else
            {
                PageLabelTableManager.GetPageLabels()[key] = value;
            }
        }

        private void UpdatePageLists(string key, List<string> outputList)
        {
            key = key.Replace("{", "").Replace("}", "");
            PageLabelTableManager.GetPageList().Add(key, outputList);
        }
        private string ReadPageLabels(string key)
        {
            int i = 0;
            while (key.Contains("{"))
            {
                i++;
                if (key.Contains("{") && key.Contains("}"))
                {
                    var tempKey = this.ExtractDynamicPart(key);

                    if (tempKey.ToUpper().StartsWith("G_") || tempKey.ToUpper().StartsWith("M_"))
                    {
                        key = key.Replace("{" + tempKey + "}", VariableManager.ContainsVariableValue("{" + tempKey + "}") ? VariableManager.GetVariableValue("{" + tempKey + "}") : key);
                    }
                    else
                    {
                        key = key.Replace("{" + tempKey + "}", PageLabelTableManager.GetPageLabels().ContainsKey(tempKey) ? PageLabelTableManager.GetPageLabels()[tempKey] : key);
                    }
                    if (i == 20)
                    {
                        break;
                    }
                }
            }

            return key;
        }

        private string ExtractDynamicPart(string value)
        {
            if (value.Contains("{"))
            {
                value = value.Substring(value.IndexOf("{") + 1, value.IndexOf("}") - value.IndexOf("{") - 1);
            }

            return value;
        }

        private string[] ReadPageLabels(string[] values)
        {
            for (int i = 0; i < values.Count(); i++)
            {
                values[i] = ReadPageLabels(values[i]);
            }
            return values;
        }

        private List<string> ReadPageLists(string key)
        {
            string tempKey = "";
            if (key.StartsWith("{") && key.EndsWith("}"))
            {
                tempKey = key.Replace("{", "").Replace("}", "");
            }
            return PageLabelTableManager.GetPageList().ContainsKey(tempKey) ? PageLabelTableManager.GetPageList()[tempKey] : null;
        }
        private DataTable ReadPageTables(string key)
        {
            string tempKey = "";
            if (key.StartsWith("{") && key.EndsWith("}"))
            {
                tempKey = key.Replace("{", "").Replace("}", "");
            }
            return PageLabelTableManager.GetPageTables().ContainsKey(tempKey) ? PageLabelTableManager.GetPageTables()[tempKey] : null;
        }

        /// <summary>
        /// Returns formatted query from the input SQLQuery
        /// </summary>
        /// <param name="txtFileName">SQL Query file name</param>
        /// <param name="parameters">Comma seperated list of SQL query input parameters</param>
        /// <returns>Formatted SQL query</returns>
        public string GetFormattedQuery(string txtFileName, string parameters)
        {
            string Query = this.GetQueryFromTextFile(txtFileName);
            string[] parametersArray = parameters.Split(',');
            parametersArray = ReadPageLabels(parametersArray);
            for (int i = 0; i < parametersArray.Length; i++)
            {
                Query = Query.Replace("{" + i + "}", parametersArray[i].Trim());
            }
            return Query;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="QueryFileName"></param>
        /// <returns></returns>
        private string GetQueryFromTextFile(string QueryFileName)
        {
            string Query = "";
            try
            {
                StreamReader reader = new StreamReader(Path.Combine(ConfigHelper.DBQueryFolderPath, QueryFileName));
                Query = reader.ReadToEnd();
                reader.Close();
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.InnerException.ToString());
            }
            return Query;
        }
        private bool VerifyCheckBoxChecked(string uiElement, string uiElementValue)
        {
            return KeyWordsMethods.VerifyCheckBoxChecked(uiElement, uiElementValue);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="uiElement"> element whose visiblity has to check</param>
        /// <param name="uiElementValue">element who needs to be click</param>
        /// <returns></returns>
        private bool ClickIfVisible(string uiElement, string uiElementValue)
        {
            return KeyWordsMethods.ClickIfVisible(uiElement, uiElementValue);
        }
        #endregion
    }
}
#endregion