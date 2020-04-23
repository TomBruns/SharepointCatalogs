using System;
using System.Collections.Generic;
using System.Text;

namespace FIS.USESA.POC.Sharepoint.Selinium.Entities
{
    public class BusinessProcessBE
    {
        public string Code { get; set; }
        public string ShortDescription {get; set;}
        public string Location { get; set; }
        public string Description { get; set; }
        public string RTO { get; set; }

        public decimal? RTONum 
        {
            get 
            {
                if (decimal.TryParse(this.RTO, out decimal result))
                {
                    return result;
                }
                else
                {
                    return (decimal?)null;
                }
            }
        }

        public string Owner { get; set; }
        public string Status { get; set; }
    }
}
