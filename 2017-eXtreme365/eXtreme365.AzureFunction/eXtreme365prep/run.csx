#r "Newtonsoft.Json"
#r "eXtreme365.FunctionAssembly.dll"

using System;
using System.Configuration;
using System.Net;
using Newtonsoft.Json;
using eXtreme365.FunctionAssembly;


public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info($"Running function");

    try
    {
        //getting app settings
        var user = ConfigurationManager.AppSettings["user"];
        var password = ConfigurationManager.AppSettings["password"];
        var crmurl = ConfigurationManager.AppSettings["crmurl"];

        //reading json data
        string jsonContent = await req.Content.ReadAsStringAsync();
        dynamic data = JsonConvert.DeserializeObject(jsonContent);

        //create sharepoint web
        SharePoint.CreateSharePointWeb($"{data.url}", user, password, $"{data.title}", $"{data.urlpart}");

        //construct the url for the new sharepoint web
        var weburl = $"{data.url}/{data.urlpart}";
        //create the content on the new sharepoint web
        SharePoint.CreateSharePointWebContent(weburl, user, password);

        //create the document locations in CRM for this SharePoint site and its libraries
        CRM.CreateDocumentLocations(crmurl, user, password, $"{data.urlpart}", $"{data.entityname}", $"{data.recordid}");

        log.Info($"Function succeeded");
        //create the success response for the CRM workflow
        return req.CreateResponse(HttpStatusCode.OK, new
        {
            result = $"Success"
        });
    }
    catch (Exception ex)
    {
        log.Info($"{ex.Message}");
        //return the exception to CRM 
        return req.CreateResponse(HttpStatusCode.InternalServerError, new
        {
            error = ex.Message
        });
    }

    //sample json 
    //{
    //    "url": "https://oak3crm.sharepoint.com/sites/Dynamics365",
    //    "title": "Test site",
    //    "urlpart": "12345",
    //    "entityname": "account",
    //    "recordid": "C221AD5B-11F6-E611-80FE-5065F38B9561"
    //}
}