using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using System.Net;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Query;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk.Messages;

namespace Connect2Dynamics365Online
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create IOrgannization Service Object

            IOrganizationService service = ConnectD35OnlineUsingOrgSvc();
            var myemail = "alex@treyresearch.net";
            var fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='contact'>
                   <attribute name='fullname' />
                     <attribute name='telephone1' />
                         <attribute name='contactid' />
                        <order attribute='fullname' descending='false' />
                              <filter type='and'>
                         <condition attribute='emailaddress1' operator='eq' value=(0) />
                          </filter>
                                </entity>
                          </fetch>";
            fetchXml = String.Format(fetchXml, myemail);
            EntityCollection contacts = service.RetrieveMultiple(new FetchExpression(fetchXml));
            Console.WriteLine("Total record: " + contacts.Entities.Count);
            foreach (var contact in contacts.Entities)
            {
                Console.WriteLine("Id" + contact.Id);
                if (contact.Contains("fullname") && contact["fullname"] != null)
                {
                    Console.WriteLine("Contact Name: " + contact["fullname"]);
                }
            }
  
            Console.ReadLine();
        }

        public static IOrganizationService ConnectD35OnlineUsingOrgSvc()
        {
            IOrganizationService organizationService = null;

            String username = "SHAHZAIBALI@SHAHZAIBSAFDAR1.onmicrosoft.com";
            String password = "safdar786ALI!";

            String url = "https://org666f01ac.crm.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=dashboard&id=2701de60-8f2a-48a4-8262-4a35ca7441fa&type=system&_canOverride=true";
            try
            {
                ClientCredentials clientCredentials = new ClientCredentials();
                clientCredentials.UserName.UserName = username;
                clientCredentials.UserName.Password = password;

                // For Dynamics 365 Customer Engagement V9.X, set Security Protocol as TLS12
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                organizationService = (IOrganizationService)new OrganizationServiceProxy(new Uri(url), null, clientCredentials, null);

                if (organizationService != null)
                {
                    Guid gOrgId = ((WhoAmIResponse)organizationService.Execute(new WhoAmIRequest())).OrganizationId;
                    if (gOrgId != Guid.Empty)
                    {
                        Console.WriteLine("Connection Established Successfully...");
                    }
                }
                else
                {
                    Console.WriteLine("Failed to Established Connection!!!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception occured - " + ex.Message);
            }
            return organizationService;

        }
    }
}
