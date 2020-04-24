# [Sharepoint Catalogs](https://github.com/TomBruns/SharepointCatalogs)

The purpose of this project is to automate loading a Sharepoint list from data in an Excel file.

---
## Technologies Leveraged
|Technology | Description | Addl Info |
|---- | ------------ | ------- |
| Console App  | Automation Logic (C#) | .Net Core v3.1 |
| Microsoft Edge (Chromium) | Browser | v81.0.416.58 |
| Microsoft Edge Driver | Webdriver compatible driver | https://msedgewebdriverstorage.z22.web.core.windows.net/ (included in solution) |
| Selenium Client | Brower Automation API | 4.0.0-alpha05 (via nuGet) |
| Selenium WebDriver Language Bindings | C# language-specific client driver | 4.0.0-alpha05 (via nuGet) |
| SpreadsheetGear | Excel (xlsx) file integration API | https://www.spreadsheetgear.com/ (via nuget)

> **Note**: To use Selenium you need both a **WebDriver** (that matches the version of browser you are automating and provided by the browser maker) and a **Language Binding** assembly (that matches the automation language you are using and provided by Selenium).

> **Note**: Alternate Browsers could be used with the corresponding browser version specific driver (nothing in the code is Microsoft Edge Browser specific).  The automation code uses the W3C Webdriver API to control the behavior of a web browser.  Each browser manufacturer typically supplies a Webdriver compatible driver.

---
## Solution Architecture

A .Net Core console app reads the data out of an excel file and uses selenium to automate a browser session.

![CSProj Changes](images/Architecture.jpg?raw=true)

> **Note**: The associated Excel file is **NOT** included in the Source Code Repo!

> **Note**: Since we are actually driving a Browser Session, we will authenticate to Sharepoint as the current user.
---
## Logical Flow

Here is logically how the functionality works

![CSProj Changes](images/Logical_Process.jpg?raw=true)

---
## Configuration Options

The following configuration options are available in the `appSettings.json` file:

```json
{
  "loadConfig": {
    "catalogType": "BUSINESS_PROCESSES",
    "excelFilePathName": ".\\ProcessesOwners20200310_Ramesh updated v2.xlsx",
    "rtoFilter": [
      1,
      2,
      4,
      24
    ],
    "worksheetName": "Processes with Sites grouped by",
    "browserLocation": "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
    "sharepointURL": "https://gsp.worldpay.com/sites/ITStrategyandArchitecture/SitePages/Home.aspx"
  }
}
```

| Config Parameter | Description | Options |
|---- | ------------ | ------- |
| catalogType | Which Catalog (Sharepoint List) to load | BUSINESS_PROCESSES |
| excelFilePathName | Pathname to the excel file (xlsx) containing the data to load | |
| rtoFilter | RTO Values used to filter rows in the excel file | 0.25, 0.5, 1, 2, 4, 24, 48, 72, 120, 168, 336, 504 |
| worksheetName | Name of the worksheet in the Excel file containing the data to load| |
| browserLocation | Pathname to the browser we are automating | |
| sharepointURL | URL of the EA homepage on the Sharepoint Site | |

---
## Interesting Challenges

* The Sharepoint pages are built dynamically so the page elements have random names.  This made selecting page elements by ID not feasible.  The alternative approach used was selecting by XPath:

```csharp
   var businessProcessesLink = driver.FindElementByXPath("//a[text()='CAT-010 - Business Process']")
```
* The Code field had some unusual behavior when using sendkeys to set the entire string (it was randomly dropping the 3rd character).  I worked around it by sending each character of the string separately
```csharp
   // this field was randomly dropping characters, I assume this is due to something unique about this field
   // so I used this character by character approach to slow the entry down
   codeTextInputField.Clear();
   for (int i = 0; i < newBusinessProcess.Code.Length; i++)
   {
       string letter = newBusinessProcess.Code[i].ToString();
       codeTextInputField.SendKeys(letter);
   }
   //codeTextInputField.SendKeys(newBusinessProcess.Code);
```
* The `Save` button on the `New Item` page seems to only be enabled based on some difficult to automate interaction between the user's mouse and the field on the screen. It is executed instead using this button's alternate access method `ALT-O`

```csharp
   // funny enough, the Save Button is mapped to ALT-O on this page
   new Actions(driver).KeyDown(Keys.Alt).SendKeys("O").Perform();
```

* Since Sharepoint builds the pages dynamically, sometimes the automation needs to wait until the target page element is available in the DOM.

```csharp
   var codeTextInputField = wait.Until(ExpectedConditions.ElementExists(By.XPath("//input[@title='Code Required Field']")));
```

* The Sharepoint site uses Federated Authentication.  I was not successful trying to automate this interaction (This would have enabled the Sharepoint Client APIs to be used).
