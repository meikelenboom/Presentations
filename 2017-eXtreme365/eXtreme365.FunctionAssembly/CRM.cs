using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Configuration;

namespace eXtreme365.FunctionAssembly
{
    public static class CRM
    {
        public static void CreateDocumentLocations(string crmurl, string user, string password, string urlpart, string entityName, string recordId)
        {
            //build the 'new style' connection string and create a proxy.
            string connString = string.Format("AuthType=Office365;Url={0}; Domain=CONTOSO; Username={1}; Password={2}", crmurl, user, password);
            var connStringSetting = new ConnectionStringSettings("xrm", connString);
            var serviceClient = new CrmServiceClient(connString);
            var serviceProxy = serviceClient.OrganizationServiceProxy;

            //create the sharepoint web and use the already configured site collection as parent site
            var site = new Entity("sharepointsite");
            site["name"] = string.Format("Site-{0}", urlpart);
            site["relativeurl"] = urlpart;
            site["parentsite"] = new EntityReference("sharepointsite", new Guid("518DDDEE-0FF6-E611-80FE-5065F38B9561")); //this shouldn't be hard coded, but queried
            var siteRef = serviceProxy.Create(site);

            //create the document locations with the web as a parent
            var general = new Entity("sharepointdocumentlocation");
            general["name"] = "General";
            general["relativeurl"] = "General";
            general["regardingobjectid"] = new EntityReference(entityName, new Guid(recordId));
            general["parentsiteorlocation"] = new EntityReference("sharepointsite", siteRef);
            var generalId = serviceProxy.Create(general);

            var quotes = new Entity("sharepointdocumentlocation");
            quotes["name"] = "Quotes";
            quotes["relativeurl"] = "Quotes";
            quotes["regardingobjectid"] = new EntityReference(entityName, new Guid(recordId));
            quotes["parentsiteorlocation"] = new EntityReference("sharepointsite", siteRef);
            var quitesId = serviceProxy.Create(quotes);

            var projects = new Entity("sharepointdocumentlocation");
            projects["name"] = "Projects";
            projects["relativeurl"] = "Projects";
            projects["regardingobjectid"] = new EntityReference(entityName, new Guid(recordId));
            projects["parentsiteorlocation"] = new EntityReference("sharepointsite", siteRef);
            var projectsId = serviceProxy.Create(projects);
        }
    }
}
