using System;
using System.Collections.Generic;
using System.Text;

namespace FIS.USESA.POC.Sharepoint.Selinium
{
    static class Constants
    {
        public enum CATALOG_TYPES
        {
            UNKNOWN = 0,
            BUSINESS_PROCESSES
        }

        public enum BUSINESS_PROCESS_GRID_COLS
        {
            CHECKBOX = 0,
            CODE,
            MORE_OPTIONS,
            SHORT_DESCRIPTION,
            LOCATION,
            DESCRIPTION,
            RTO,
            OWNER,
            STATUS
        }

        public enum BUSINESS_PROCESS_EXCEL_COLS
        {
            PLAN_TIER = 0,
            PLAN_NAME,
            BUSINESS_PLAN_OWNER,
            PROCESS_NAME_1,
            PROCESS_NAME_2,
            SITE_NAME,
            NO_OF_STAFF,
            FINAL_RTO_HOURS,
            FINAL_TIER,
            PROCESS_MANAGER,
            SUBJECT_MATTER_EXPERT,
            PROCESS_DESCRIPTION
        }
    }
}
