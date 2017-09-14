using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace openXML_CSV_ServiceStack
{
    public class CSPInvoice
    {
        public string CustomerName { get; set; }
        public string OfferName { get; set; }
        public DateTime ChargeStartDate { get; set; }
        public DateTime ChargeEndDate { get; set; }
        public string ChargeType { get; set; }
        public string UnitPrice { get; set; }
        public string Quantity { get; set; }
        public int ChargeDaterange { get; set; }
        public int ChargeStartMonthDays { get; set; }
        
        //塞入資料
        public void addvalue(string propName, string value)
        {
            if (propName.Equals("CustomerName"))
            {
                CustomerName = value;
            }
            if (propName.Equals("OfferName"))
            {
                OfferName = value;
            }
            if (propName.Equals("ChargeStartDate"))
            {
                ChargeStartDate = Convert.ToDateTime(value);
            }
            if (propName.Equals("ChargeEndDate"))
            {
                ChargeEndDate = Convert.ToDateTime(value);
            }
            if (propName.Equals("ChargeType"))
            {
                ChargeType = value;
            }
            if (propName.Equals("UnitPrice"))
            {
                UnitPrice = value;
            }
            if (propName.Equals("Quantity"))
            {
                Quantity = value;
            }
            if (propName.Equals("ChargeDaterange"))
            {
                ChargeDaterange = Convert.ToInt32(value);
            }
            if (propName.Equals("ChargeStartMonthDays"))
            {
                ChargeStartMonthDays = Convert.ToInt32(value);
            }
        }
    }
}
