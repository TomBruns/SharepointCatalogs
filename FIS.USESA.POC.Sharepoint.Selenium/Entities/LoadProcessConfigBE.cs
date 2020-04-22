using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json.Serialization;

using static FIS.USESA.POC.Sharepoint.Selinium.Constants;

namespace FIS.USESA.POC.Sharepoint.Selenium.Entities
{
    public class LoadProcessConfigBE
    {
        [JsonPropertyName("catalogType")]
        public CATALOG_TYPES CatalogType { get; set; }

        [JsonPropertyName("excelFilePathName")]
        public string ExcelFilePathName { get; set; }

        [JsonPropertyName("rtoFilter")]
        public List<string> RtoFilter { get; set; }

        [JsonPropertyName("worksheetName")]
        public string WorksheetName { get; set; }

        [JsonPropertyName("browserLocation")]
        public string BrowserLocation { get; set; }

        [JsonPropertyName("sharepointURL")]
        public string SharepointURL { get; set; }
    }
}
