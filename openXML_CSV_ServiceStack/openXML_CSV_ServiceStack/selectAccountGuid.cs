using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace openXML_CSV_ServiceStack
{
    class selectAccountGuid
    {
        public string Name;
        public static IOrganizationService service;
        public Guid myGuid { get; set; }
        public string message { get; set; }
        public selectAccountGuid(IOrganizationService importService, string improtnweName)
        {
            Name = improtnweName;
            service = importService;
            message = "Error#Not ready";
        }
        public void startdate()
        {
            string FetchXml_nweName =
              @"<fetch>
                  <entity name='account'>
                     <attribute name='name'/>
                     <attribute name='accountid'/>
                     <filter>
                         <condition attribute='name' operator='eq' value='{{name}}'/>
                     </filter>
                   </entity>
                </fetch> ";

            FetchXml_nweName = FetchXml_nweName.Replace("{{name}}", Name);
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
                myGuid = (Guid)result.Entities[0].Attributes["accountid"];
                for (int r = 0; r < result.Entities.Count; r++)
                {
                    Console.WriteLine(result.Entities[r]);
                }
            }
        }
    }
}
