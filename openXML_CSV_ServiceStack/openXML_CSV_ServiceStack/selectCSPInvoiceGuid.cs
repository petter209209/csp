using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace openXML_CSV_ServiceStack
{
    class selectCSPInvoiceGuid
    {
        public string nweName;
        public static IOrganizationService service;
        public Guid myGuid { get; set; }
        public string message { get; set; }
        public selectCSPInvoiceGuid(IOrganizationService importService,string improtnweName)
        {
            nweName = improtnweName;
            service = importService;
            message = "Error#Not ready";
        }
        public void startdate()
        {
            string FetchXml_nweName =
              @"<fetch>
                   <entity name='new_csp_cloud_invoice'>
                          <attribute name='new_csp_cloud_invoiceid'/>   
                          <attribute name='new_name'/>
                              <filter>
                                <condition attribute='new_name' operator='eq' value='{{nweName}}'/>
                              </filter>
                    </entity>
               </fetch>";

            FetchXml_nweName = FetchXml_nweName.Replace("{{nweName}}", nweName);
            FetchExpression nweName_result = new FetchExpression(FetchXml_nweName);
            //result中斷點爪ID資訊
            EntityCollection result = service.RetrieveMultiple(nweName_result);
            if (result.Entities.Count() == 0)
            {
                message = "Error#1001";
            }
            else
            {
                message = "Successfully crawl Guid";
                myGuid = (Guid)result.Entities[0].Attributes["new_csp_cloud_invoiceid"];
            }
        }
    }
}
