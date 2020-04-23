using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;

using Microsoft.Extensions.Configuration;

using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

using SeleniumExtras.WaitHelpers;

using SpreadsheetGear;

using FIS.USESA.POC.Sharepoint.Selinium.Entities;
using static FIS.USESA.POC.Sharepoint.Selinium.Constants;
using FIS.USESA.POC.Sharepoint.Selenium.Entities;

namespace FIS.USESA.POC.Sharepoint.Selenium
{
    /// <summary>
    /// This class uses Selenium to drive a browser session to auotmate entering data into a list in our Sharepoint Catalog
    /// Credit goes to Ramesh who suggested this approach.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            // load plug-in specific configuration from appsettings.json file copied into the plug-in specific subfolder 
            IConfigurationRoot configuration = new ConfigurationBuilder()
                .SetBasePath(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location))
                .AddJsonFile("appsettings.json", false)
                .Build();

            var loadProcessConfig = configuration.GetSection("loadConfig").Get<LoadProcessConfigBE>();

            //CATALOG_TYPES catalogType = CATALOG_TYPES.BUSINESS_PROCESSES;
            //string excelFilePathName = @".\ProcessesOwners20200310_Ramesh updated v2.xlsx";
            //var rtoFilter = new List<string>() { @"1", @"2", @"4", @"24" };

            //string worksheetName = @"Processes with Sites grouped by";
            //string browserLocation = @"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
            //string sharepointURL = @"https://gsp.worldpay.com/sites/ITStrategyandArchitecture/SitePages/Home.aspx";

            Utilities.WriteToConsole(@"-------------------------------------------------------------------");
            Utilities.WriteToConsole(@" Note: You can ignore the log messages from Selenium (white text)");
            Utilities.WriteToConsole(@"-------------------------------------------------------------------");

            Utilities.WriteToConsole(@"Step 1.0: Open the browser");
            var edgeOptions = new EdgeOptions()
            {
                UseChromium = true,
                BinaryLocation = loadProcessConfig.BrowserLocation
            };

            string edgeDriverDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            // Use the EdgeDriver class provided with Selenium.
            using (EdgeDriver driver = new EdgeDriver(edgeDriverDirectory, edgeOptions))
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                driver.Navigate().GoToUrl(loadProcessConfig.SharepointURL);

                #region ==== Step 1.1: Pick account to use to signin => Vantiv, now worldpay
                Utilities.WriteToConsole(@"Step 1.1: Click on Vantiv, now Worldpay");

                // find the span that has the correct text label
                var vantivAccount = driver.FindElementByXPath("//span[text()='Vantiv, now Worldpay']");

                // click event will bubble up to a parent element that has an onclick handler defined
                vantivAccount.Click();
                #endregion

                #region ==== Step 1.2: Navigate to the Catalogs Page
                Utilities.WriteToConsole(@"Step 1.2: Navigate to the Catalogs Page");

                var catalogsButton = wait.Until(ExpectedConditions.ElementExists(By.XPath("//img[@alt='catalogSmall.png']")));

                catalogsButton.Click();
                #endregion

                // call the appropriate Upload method
                switch (loadProcessConfig.CatalogType)
                {
                    case CATALOG_TYPES.BUSINESS_PROCESSES:
                        Catalogs.BusinessProcess.Upload(loadProcessConfig.ExcelFilePathName, loadProcessConfig.WorksheetName, loadProcessConfig.RtoFilter, driver, wait);
                        break;
                }
            }
        }

    }
}
