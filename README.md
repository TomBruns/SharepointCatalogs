# [Sharepoint Catalogs](https://github.com/TomBruns/SharepointCatalogs)

The purpose of this project is to automate loading a Sharepoint list from data in an Excel file.

> **Note**: The normal methods of leveraging Sharepoint's native Excel support or the Sharepoint API were not available.

---
## Technologies Leveraged
|Site | Description | Other |
|---- | ------------ | ------- |
| .Net Core Console App | Procedural Logic | v3.1 |
| Microsoft Edge (Chromium) | Browser | v81.0.416.58 |
| Microsoft Edge Driver | Excel Automation | https://msedgewebdriverstorage.z22.web.core.windows.net/ |
| Selenium Client | Brower Automation | 4.0.0-alpha05 (nuGet) |
| Selenium WebDriver Language Bindings | C# language-specific client driver | 4.0.0-alpha05 (nuGet) |
| SpreadsheetGear | Excel Automation | https://www.spreadsheetgear.com/ (nuget})

> **Note**: To use Selenium you need both a **WebDriver** (matching the version of browser you are automating and provided by the browser maker) and a **Language Binding** assembly (matching the automation language you are using and provided by Selenium).

> **Note**: Alternate Browsers could be used with the corresponding browser specific driver.
---
## Solution Architecture

A .Net Core console app reads the data out of an excel file and uses selenium to drive a browser session.

![CSProj Changes](images/Architecture.jpg?raw=true)

> **Note**: The associated Excel file is **NOT** included in the Source Code Repo!

> **Note**: Since we are actually driving a Browser Session, we will authenticate to Sharepoint as the current user.
---
## Logical Flow

Here is logically how the functionality works

![CSProj Changes](images/Logical_Process.jpg?raw=true)

---
## Interesting Challenges

* The Sharepoint pages are built dynamically so the page elements have random names.  This made selecting page elements by ID not feasible.  The approach used was selecting by XPath instead.
* The `submit` button on the `New Item` page seems to only be enabled based on some difficult to automate interaction with the mouse. It is execute instead using the alternate access method `ALT-O`
* The Sharepoint site uses Federated Authentication.  I was not successful trying to automate this interaction so that the Sharepoint Client APIs could be used.