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
            //連線到Dynamicd365
            var connectionString = ConfigurationManager.ConnectionStrings["myConnectionString"].ToString();
            CrmServiceClient CSC = new CrmServiceClient(connectionString);
            service = (IOrganizationService)CSC.OrganizationWebProxyClient != null ? (IOrganizationService)CSC.OrganizationWebProxyClient : (IOrganizationService)CSC.OrganizationServiceProxy;
            Console.WriteLine("----------------------------------------------");
            foreach (CSPInvoice invoice in invoices)
            {
                
                Console.WriteLine(CreateCSPInvoice(invoice.CustomerName));
                //新增
                Console.WriteLine(CreateCSPInvoicedetail(invoice));
                Console.WriteLine("----------------------------------------------");
            }
        }
        private string CreateCSPInvoice(String name)
        {
            String invoiceName = name + "-" + filename;
            try
            {
                selectCSPInvoiceGuid CSPInvoiceGuid = new selectCSPInvoiceGuid(service, invoiceName);
                CSPInvoiceGuid.startdate();
                if (!CSPInvoiceGuid.message.Equals("Successfully crawl Guid"))
                {
                    Entity CSPTitle = new Entity("new_csp_cloud_invoice");
                    CSPTitle.Attributes["new_name"] = invoiceName;
                    selectAccountGuid accountGuid = new selectAccountGuid(service, name);
                    accountGuid.startdate();
                    if (!accountGuid.message.Substring(1, 5).Equals("Error"))
                    {
                        var CSPInvoicedetailValue = new EntityReference("new_csp_cloud_invoice", accountGuid.myGuid);
                        CSPTitle.Attributes["new_account"] = CSPInvoicedetailValue;
                        service.Create(CSPTitle);
                        return $"{"invoice",-15}{"建立成功",-10}{invoiceName,-30}";
                    }
                    else
                    {
                        return $"{"invoice",-15}{"建立失敗",-10}{invoiceName,-30}";
                    }
                }
                else
                {
                    return $"{"invoice",-15}{"已建立",-11}{invoiceName,-30}";
                }
            }
            catch
            {
                return $"{"invoice",-15}{"建立失敗",-10}{invoiceName,-30}";
            }
        }
        private string CreateCSPInvoicedetail(CSPInvoice invoice)
        {
            try
            {
                Entity CSPInvoice = new Entity("new_csp_cloud_invoicedetail");
                /*
                 * Dynamics-new_csp_cloud_invoicedetail
                 *                       欄位名稱                      CSPInvoice對應名稱             查詢欄位
                 * 品項                  new_name                      invoice.OfferName
                 * 數量                  new_quantity                  invoice.Quantity
              @  * 單價                  new_unitprice                 invoice.UnitPrice
                 * CSP雲端服務訂單明細   new_csp_cloud_invoicedetail                                  CSPInvoicedetailValue
                 * 計費開始時間          new_charge_startdate          invoice.ChargeStartDate
                 * 計費開始結束時間      new_charge_enddate            invoice.ChargeEndDate
                 * 計費開始月份總天數    new_month_days                invoice.ChargeStartMonthDays
                 * 計費天數              new_charge_days               invoice.ChargeDaterange
                */
                CSPInvoice.Attributes["new_name"] = invoice.OfferName;
                CSPInvoice.Attributes["new_charge_startdate"] = invoice.ChargeStartDate;
                CSPInvoice.Attributes["new_charge_enddate"] = invoice.ChargeEndDate;
                CSPInvoice.Attributes["new_month_days"] = invoice.ChargeStartMonthDays;
                CSPInvoice.Attributes["new_charge_days"] = invoice.ChargeDaterange;
                CSPInvoice.Attributes["new_quantity"] =Convert.ToInt32(invoice.Quantity);

                CSPInvoice.Attributes["new_unitprice"] = Convert.ToDecimal(invoice.UnitPrice);
                Console.WriteLine(Convert.ToDecimal(invoice.UnitPrice));

                //抓取對應訂單(Invoice)的Guid
                selectCSPInvoiceGuid CSPInvoiceGuid = new selectCSPInvoiceGuid(service, invoice.CustomerName + "-" + filename);
                CSPInvoiceGuid.startdate();
                if (!CSPInvoiceGuid.message.Substring(1, 5).Equals("Error"))
                {
                    //Console.WriteLine(CSPInvoiceGuid.myGuid);
                    var CSPInvoicedetailValue = new EntityReference("new_csp_cloud_invoicedetail", CSPInvoiceGuid.myGuid);
                    CSPInvoice.Attributes["new_csp_cloud_invoicedetail"] = CSPInvoicedetailValue;
                    service.Create(CSPInvoice);
                    return $"{"invoicedetail",-15}{"建立成功",-10}{invoice.OfferName,-30}";
                }
                else
                {
                    return $"{"invoicedetail",-15}{"建立失敗",-10}{invoice.OfferName,-30}{CSPInvoiceGuid.message,-20}";
                }
            }
            catch
            {
                return $"{"invoicedetail",-15}{"建立失敗",-10}{invoice.OfferName,-30}";
            }
        }

    }
}
