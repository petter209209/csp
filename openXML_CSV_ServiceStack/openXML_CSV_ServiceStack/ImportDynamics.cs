using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace openXML_CSV_ServiceStack
{
    public class ImportDynamics
    {
        List<CSPInvoice> invoices;
        static IOrganizationService service;
        const int CUSTOMER_NAME_INDEX = 2;
        string filename;
        public ImportDynamics(List<CSPInvoice> ImportInvoices,string ImportFilename)
        {
            invoices = ImportInvoices;
            filename = ImportFilename;
        }
        public void startImportDynamics()
        {
            var connectionString = ConfigurationManager.ConnectionStrings["myConnectionString"].ToString();
            CrmServiceClient CSC = new CrmServiceClient(connectionString);
            service = (IOrganizationService)CSC.OrganizationWebProxyClient != null ? (IOrganizationService)CSC.OrganizationWebProxyClient : (IOrganizationService)CSC.OrganizationServiceProxy;
            
            foreach (CSPInvoice invoice in invoices)
            {
                Console.WriteLine(CreateCSPTitle(invoice.CustomerName));
                //新增資料
                Console.WriteLine(CreateCSPInvoice(invoice));
            }
        }
        private string CreateCSPTitle(String name)
        {
            name = name + "-" + filename;
            try
            {
                selectCSPInvoiceGuid CSPInvoiceGuid = new selectCSPInvoiceGuid(service, name);
                CSPInvoiceGuid.startdate();
                if (!CSPInvoiceGuid.message.Equals("Successfully crawl Guid"))
                {
                    Entity CSPTitle = new Entity("new_csp_cloud_invoice");
                    CSPTitle.Attributes["new_name"] = name;
                    service.Create(CSPTitle);
                    return $"匯入成功{name,-30}";
                }
                else
                {
                    return $"已建立{name,-30}";
                }
            }
            catch
            {
                return $"匯入失敗{name,-30}";
            }
        }
        private string CreateCSPInvoice(CSPInvoice invoice)
        {
            try
            {
                Entity CSPInvoice = new Entity("new_csp_cloud_invoicedetail");
                CSPInvoice.Attributes["new_name"] = invoice.OfferName;
                CSPInvoice.Attributes["new_charge_startdate"] = invoice.ChargeStartDate;
                CSPInvoice.Attributes["new_charge_enddate"] = invoice.ChargeEndDate;
                CSPInvoice.Attributes["new_month_days"] = invoice.ChargeStartMonthDays;
                CSPInvoice.Attributes["new_charge_days"] = invoice.ChargeDaterange;
                CSPInvoice.Attributes["new_quantity"] =Convert.ToInt32(invoice.Quantity);
                selectCSPInvoiceGuid CSPInvoiceGuid = new selectCSPInvoiceGuid(service, invoice.CustomerName + "-" + filename);
                CSPInvoiceGuid.startdate();
                if (!CSPInvoiceGuid.message.Substring(1, 5).Equals("Error"))
                {
                    Console.WriteLine(CSPInvoiceGuid.myGuid);
                    var CSPInvoicedetailValue = new EntityReference("new_csp_cloud_invoicedetail", CSPInvoiceGuid.myGuid);
                    CSPInvoice.Attributes["new_csp_cloud_invoicedetail"] = CSPInvoicedetailValue;
                    service.Create(CSPInvoice);
                    return $"匯入成功{invoice.CustomerName,-20}";
                }
                else
                {
                    return $"匯入失敗{invoice.CustomerName,-20}{CSPInvoiceGuid.message,-20}";
                }
            }
            catch
            {
                return $"匯入失敗{invoice.CustomerName,-20}";
            }
        }

    }
}
