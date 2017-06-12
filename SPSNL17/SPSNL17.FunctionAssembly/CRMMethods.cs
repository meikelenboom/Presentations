using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPSNL17.FunctionAssembly
{
    public static class CRMMethods
    {
        /// <summary>
        /// Creates a SharePointSite record in CRM en SharePointDocumentLocations
        /// </summary>
        /// <param name="username">User name</param>
        /// <param name="password">Password</param>
        /// <param name="crmUrl">The url of the CRM instance</param>
        /// <param name="urlPart">The url part used for creating the web</param>
        /// <param name="entityName">The entity name of the parent record</param>
        /// <param name="recordId">The record is of the parent record</param>
        public static void CreateDocumentLocations(string username, string password, string crmUrl, string urlPart, string entityName, string recordId)
        {
            //Creating a connection to CRM
            string connString = string.Format("AuthType=Office365;Url={0}; Domain=CONTOSO; Username={1}; Password={2}", crmUrl, username, password);
            var connStringSetting = new ConnectionStringSettings("xrm", connString);
            var serviceClient = new CrmServiceClient(connString);
            var serviceProxy = serviceClient.OrganizationServiceProxy;

            //create the sharepoint web and use the already configured site collection as parent site
            var site = new Entity("sharepointsite");
            site["name"] = string.Format("Site-{0}", urlPart);
            site["relativeurl"] = urlPart;
            site["parentsite"] = new EntityReference("sharepointsite", new Guid("11a3359d-1043-e711-8100-5065f38b0301")); //this shouldn't be hard coded, but queried
            var siteRef = serviceProxy.Create(site);

            //create the document locations with the web as a parent
            var general = new Entity("sharepointdocumentlocation");
            general["name"] = SPMethods.GeneralDocuments.Title;
            general["relativeurl"] = SPMethods.GeneralDocuments.CustomUrl;
            general["regardingobjectid"] = new EntityReference(entityName, new Guid(recordId));
            general["parentsiteorlocation"] = new EntityReference("sharepointsite", siteRef);
            serviceProxy.Create(general);

            var quotes = new Entity("sharepointdocumentlocation");
            quotes["name"] = SPMethods.QuotesDocuments.Title;
            quotes["relativeurl"] = SPMethods.QuotesDocuments.CustomUrl;
            quotes["regardingobjectid"] = new EntityReference(entityName, new Guid(recordId));
            quotes["parentsiteorlocation"] = new EntityReference("sharepointsite", siteRef);
            serviceProxy.Create(quotes);

            var projects = new Entity("sharepointdocumentlocation");
            projects["name"] = SPMethods.ProjectDocuments.Title;
            projects["relativeurl"] = SPMethods.ProjectDocuments.CustomUrl;
            projects["regardingobjectid"] = new EntityReference(entityName, new Guid(recordId));
            projects["parentsiteorlocation"] = new EntityReference("sharepointsite", siteRef);
            serviceProxy.Create(projects);
        }
    }
}
