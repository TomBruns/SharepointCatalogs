using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;

using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

using SeleniumExtras.WaitHelpers;

using SpreadsheetGear;

using FIS.USESA.POC.Sharepoint.Selinium.Entities;
using static FIS.USESA.POC.Sharepoint.Selinium.Constants;

namespace FIS.USESA.POC.Sharepoint.Selenium
{
    /// <summary>
    /// This class uses Selenium to drive a browser session to auotmate entering data into the Business Process list in our Sharepoint Catalog
    /// Credit goes to Ramesh who suggested this approach.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            WriteToConsole("Load Business Process Catalog...");

            string excelFilePathName = @".\ProcessesOwners20200310_Ramesh updated v2.xlsx";
            string worksheetName = @"Processes with Sites grouped by";
            string browserLocation = @"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
            string sharepointURL = @"https://gsp.worldpay.com/sites/ITStrategyandArchitecture/SitePages/Home.aspx";

            #region Step 0.1: Load Business Processes from Excel
            WriteToConsole(@"Step 0.1: Load Business Processes from Excel");

            Dictionary<string, BusinessProcessBE> newBusinessProcesses = LoadBusinessProcessesFromExcel(excelFilePathName, worksheetName);

            var rtoFilter = new List<string>() { @"1", @"2", @"4", @"24" };
            var filteredNewBusinessProcesses = newBusinessProcesses.Values
                                                .Where(v => rtoFilter.Contains(v.RTO))
                                                .OrderBy(v => v.RTONum).ThenBy(v => v.ShortDescription)
                                                .ToList();
            #endregion

            WriteToConsole(@"Step 1.0: Open the browser");
            var edgeOptions = new EdgeOptions()
            {
                UseChromium = true,
                BinaryLocation = browserLocation
            };

            string edgeDriverDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            // Use the EdgeDriver class provided with Selenium.
            using (EdgeDriver driver = new EdgeDriver(edgeDriverDirectory, edgeOptions))
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                driver.Navigate().GoToUrl(sharepointURL);

                #region ==== Step 1.1: Pick account to use to signin => Vantiv, now worldpay
                WriteToConsole(@"Step 1.1: Click on Vantiv, now Worldpay");

                // find the span that has the correct text label
                var vantivAccount = driver.FindElementByXPath("//span[text()='Vantiv, now Worldpay']");

                // click event will bubble up to a parent element that has an onclick handler defined
                vantivAccount.Click();
                #endregion

                #region ==== Step 1.2: Navigate to the Catalogs Page
                WriteToConsole(@"Step 1.2: Navigate to the Catalogs Page");

                var catalogsButton = wait.Until(ExpectedConditions.ElementExists(By.XPath("//img[@alt='catalogSmall.png']")));

                catalogsButton.Click();
                #endregion

                #region ==== Step 1.3: Navigate to the Business Processes Page
                var businessProcessesLink = wait.Until(ExpectedConditions.ElementExists(By.XPath("//a[text()='CAT-010 - Business Process']")));
                //var businessProcessesLink = driver.FindElementByXPath("//a[text()='CAT-010 - Business Process']");

                businessProcessesLink.Click();

                WriteToConsole(@"Step 1.3: Navigate to the Business Processes Page");
                #endregion

                // create a dictionary to hold the current entries
                Dictionary<string, BusinessProcessBE> existingBusinessProcesses = new Dictionary<string, BusinessProcessBE>();

                #region Step 2.1: Load all of the existing Business Processes
                bool isMorePages = true;
                int pageCount = 1;
                while (isMorePages)
                {
                    WriteToConsole($"Step 2.1: Load page [{pageCount}] of existing Business Processes");

                    isMorePages = LoadExisitingBusinessProcesses(driver, wait, existingBusinessProcesses);
                    pageCount++;
                }
                #endregion

                int largestCodePartSuffix = -1;
                #region Step 2.2 Find the highest Code number used so far

                System.Console.WriteLine(@"Step 2.2: Find the highest Code number used so far");

                string[] codeParts;

                int codePartSuffix = -1;
                foreach (var existingBusinessProcess in existingBusinessProcesses)
                {
                    codeParts = existingBusinessProcess.Value.Code.Split("-");

                    codePartSuffix = Int32.Parse(codeParts[1]);

                    if(codePartSuffix > largestCodePartSuffix)
                    {
                        largestCodePartSuffix = codePartSuffix;
                    }
                }

                #endregion

                #region Step 3 Upload new Business Processes
                foreach (var filteredNewBusinessProcess in filteredNewBusinessProcesses)
                {
                    if (!existingBusinessProcesses.ContainsKey(filteredNewBusinessProcess.ShortDescription))
                    {
                        // Step 3.1 Find New Item Link
                        WriteToConsole(@"Step 3.1 Find New Item Link");

                        // find the span that has the correct text label
                        //var newItem = driver.FindElementByXPath("//span[text()='new item']");
                        var newItem = wait.Until(ExpectedConditions.ElementExists(By.XPath("//span[text()='new item']")));

                        newItem.Click();

                        // Step 3.2 Upload new Business Processes
                        WriteToConsole(@"Step 3.2 Upload new Business Processes");

                        largestCodePartSuffix++;

                        var newBusinessProcess = new BusinessProcessBE()
                        {
                            Code = $"BPC-{largestCodePartSuffix:D3}",
                            ShortDescription = filteredNewBusinessProcess.ShortDescription,
                            Location = string.Empty,
                            Description = filteredNewBusinessProcess.Description,
                            RTO = filteredNewBusinessProcess.RTO,
                            Owner = !string.IsNullOrEmpty(filteredNewBusinessProcess.Owner) ? filteredNewBusinessProcess.Owner : @"TBD",
                            Status = "Requested"
                        };

                        UploadNewBusinessProcess(driver, wait, newBusinessProcess);
                    }
                }
                #endregion
            }
        }

        /// <summary>
        /// Load Business Processes From Excel File
        /// </summary>
        /// <param name="filePathName"></param>
        /// <param name="worksheetName"></param>
        /// <returns></returns>
        private static Dictionary<string, BusinessProcessBE> LoadBusinessProcessesFromExcel(string filePathName, string worksheetName)
        {
            Dictionary<string, BusinessProcessBE> newBusinessProcesses = new Dictionary<string, BusinessProcessBE>();

            // get the workbook.
            SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(filePathName);
            SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets[worksheetName];
            SpreadsheetGear.IRange usedRange = worksheet.UsedRange;

            string processName;
            string finalRTOHrs;
            string processManager;
            string processDescription;

            for (int rowIndex = 1; rowIndex <= usedRange.RowCount; rowIndex++)
            {
                processName = worksheet.Cells[rowIndex, (int)EXCEL_COLS.PROCESS_NAME_2].Text;
                finalRTOHrs = worksheet.Cells[rowIndex, (int)EXCEL_COLS.FINAL_RTO_HOURS].Text;
                processManager = worksheet.Cells[rowIndex, (int)EXCEL_COLS.PROCESS_MANAGER].Text;
                processDescription = worksheet.Cells[rowIndex, (int)EXCEL_COLS.PROCESS_DESCRIPTION].Text;

                if (!newBusinessProcesses.ContainsKey(processName))
                {
                    newBusinessProcesses.Add(processName, new BusinessProcessBE()
                    {
                        Description = processDescription,
                        Owner = processManager,
                        RTO = finalRTOHrs,
                        ShortDescription = processName
                    });
                }
            }

            return newBusinessProcesses;
        }

        /// <summary>
        /// Pages thru the result grid and loads all of the existing Business Proecesses
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="wait"></param>
        /// <param name="existingBusinessProcesses"></param>
        /// <returns></returns>
        private static bool LoadExisitingBusinessProcesses(EdgeDriver driver, WebDriverWait wait, Dictionary<string, BusinessProcessBE> existingBusinessProcesses)
        {
            // Get the main table
            var mainBizProcessTable = wait.Until(ExpectedConditions.ElementExists(By.XPath("//table[@summary='CAT-010 - Business Process']")));
            //var mainBizProcessTable = driver.FindElementByXPath("//table[@summary='CAT-010 - Business Process']");

            //var mainBizProcessTableBody = wait.Until(ExpectedConditions.ElementExists(By.XPath("child::tbody")));
            var mainBizProcessTableBody = mainBizProcessTable.FindElement(By.XPath("child::tbody"));

            //System.Console.WriteLine($"......Pausing 10 secs to let the DOM settle after dynamically building the table");
            //Thread.Sleep(10000);

            // workaround because xpath axis queries do not work in wait.Until
            //var tableRowsTest = wait.Until(ExpectedConditions.ElementExists(By.XPath("child::tr")));
            IReadOnlyCollection<IWebElement> tableRows = null;
            for (int loopctr = 1; loopctr <= 5; loopctr++)
            {
                try
                {
                    tableRows = mainBizProcessTableBody.FindElements(By.XPath("child::tr"));
                    break;
                }
                catch
                {
                    // exception is thrown if FindElements target does not exist, sleep for 1 sec
                    Thread.Sleep(1000);
                }
            }

            if(tableRows==null)
            {
                throw new ApplicationException($"Error getting row collection in table");
            }

            // declare outside of loop to reduce gc pressure
            IWebElement codeCell;
            IWebElement shortDescriptionCell;
            IWebElement locationCell;
            IWebElement descriptionCell;
            IWebElement rtoCell;
            IWebElement ownerCell;
            IWebElement statusCell;

            // loop thru each each row in the table
            foreach (var tableRow in tableRows)
            {
                // occasional stale element
                var tableRowCells = tableRow.FindElements(By.XPath("child::td"));

                // code
                codeCell = tableRowCells[(int)BUSINESS_PROCESS_GRID_COLS.CODE];
                var code = codeCell.FindElement(By.XPath("./div/a[text()]")).Text;

                // shortDesription
                shortDescriptionCell = tableRowCells[(int)BUSINESS_PROCESS_GRID_COLS.SHORT_DESCRIPTION];
                var shortDescription = shortDescriptionCell.Text;

                WriteToConsole($"...... Loading code [{code}] [{shortDescription}]");

                // location
                locationCell = tableRowCells[(int)BUSINESS_PROCESS_GRID_COLS.LOCATION];
                //var location = locationCell.Text;

                // desription
                descriptionCell = tableRowCells[(int)BUSINESS_PROCESS_GRID_COLS.DESCRIPTION];
                //var description = descriptionCell.Text;

                // rto
                rtoCell = tableRowCells[(int)BUSINESS_PROCESS_GRID_COLS.RTO];
                //var rto = rtoCell.Text;

                // owner
                ownerCell = tableRowCells[(int)BUSINESS_PROCESS_GRID_COLS.OWNER];
                //var owner = ownerCell.Text;

                // status
                statusCell = tableRowCells[(int)BUSINESS_PROCESS_GRID_COLS.STATUS];
                //var status = statusCell.FindElement(By.XPath("./a[text()]")).Text;

                // add an entry to the collection
                existingBusinessProcesses.Add(shortDescriptionCell.Text, new BusinessProcessBE()
                {
                    Code = codeCell.FindElement(By.XPath("./div/a[text()]")).Text,
                    ShortDescription = shortDescriptionCell.Text,
                    Location = locationCell.Text,
                    Description = descriptionCell.Text,
                    RTO = rtoCell.Text,
                    Owner = ownerCell.Text,
                    Status = statusCell.FindElement(By.XPath("./a[text()]")).Text

                });
            }

            try
            {
                // look for the "NEXT" page button
                var nextPageLabel = driver.FindElementByXPath("//img[@alt='Next']");

                var nextPageButton = nextPageLabel.FindElement(By.XPath("parent::span/parent::a"));

                nextPageButton.Click();

                WriteToConsole($".....Navigating to next page of data");

                // There are more pages
                return true;
            }
            catch(Exception ex)
            {
                // The Next button does not exist
                return false;
            }
        }

        /// <summary>
        /// Uploads a new Business Process
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="wait"></param>
        /// <param name="newBusinessProcess"></param>
        private static void UploadNewBusinessProcess(EdgeDriver driver, WebDriverWait wait, BusinessProcessBE newBusinessProcess)
        {
            WriteToConsole($"...... Uploading Code: [{newBusinessProcess.Code}] [{newBusinessProcess.ShortDescription}]");

            var codeTextInputField = wait.Until(ExpectedConditions.ElementExists(By.XPath("//input[@title='Code Required Field']")));
            //var codeTextInputField = driver.FindElementByXPath("//input[@title='Code Required Field']");

            // this field was randomly dropping characters, I assume this is due to something unique about this field
            // so I used this character by character approach to slow the entry down
            codeTextInputField.Clear();
            for (int i = 0; i < newBusinessProcess.Code.Length; i++)
            {
                string letter = newBusinessProcess.Code[i].ToString();
                codeTextInputField.SendKeys(letter);
            }
            //codeTextInputField.SendKeys(newBusinessProcess.Code);

            var shortDescriptionTextInputField = driver.FindElementByXPath("//input[@title='Short Description (Name) Required Field']");
            shortDescriptionTextInputField.SendKeys(newBusinessProcess.ShortDescription);

            var locationField = driver.FindElementByXPath("//select[@title='Location']");
            var locationSelectField = new SelectElement(locationField);
            locationSelectField.SelectByText(newBusinessProcess.Location);

            var descriptionTextInputField = driver.FindElementByXPath("//input[@title='Description']");
            descriptionTextInputField.SendKeys(newBusinessProcess.Description);

            var rtoField = driver.FindElementByXPath("//select[@title='RTO Required Field']");
            var rtoSelectField = new SelectElement(rtoField);
            rtoSelectField.SelectByText(newBusinessProcess.RTO);

            var ownerTextInputField = driver.FindElementByXPath("//input[@title='Owner Required Field']");
            ownerTextInputField.SendKeys(newBusinessProcess.Owner);

            var statusField = driver.FindElementByXPath("//select[@title='Status Required Field']");
            var statusSelectField = new SelectElement(statusField);
            statusSelectField.SelectByText(newBusinessProcess.Status);

            // I tried alot of options to click the save button until i found 1 that worked!
            //System.Console.WriteLine($"......Pausing 5 secs to let the DOM settle and for the Save button to become interactable");
            //Thread.Sleep(5000);

            // timeout error
            //var saveButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//input[@value='Save']")));

            // element not interactable
            //var saveButton = driver.FindElementByXPath("//input[@type='button' and @value='Cancel']");
            //saveButton.Click();

            // javascript error: Failed to execute 'elementsFromPoint' on 'Document': The provided double value is non-finite.
            //new Actions(driver).MoveToElement(saveButton).Click().Perform();

            new Actions(driver).KeyDown(Keys.Alt).SendKeys("O").Perform();

        }

        /// <summary>
        /// Write a log message to the console
        /// </summary>
        /// <param name="message"></param>
        private static void WriteToConsole(string message)
        {
            ConsoleColor defaultColor = System.Console.ForegroundColor;
            System.Console.ForegroundColor = ConsoleColor.Blue;
            System.Console.WriteLine(message);
            System.Console.ForegroundColor = defaultColor;
        }
    }
}
