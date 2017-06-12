using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using System;
using SPSNL17.FunctionAssembly;

namespace SPSNL17.Function
{
    public static class CreateLocations
    {
        [FunctionName("CreateLocations")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            //get the request body
            dynamic data = await req.Content.ReadAsAsync<object>();

            //get the values from the request body
            string accountName = data.accountName;
            string accountNumber = data.accountNumber;
            string entityName = data.entityName;
            string recordId = data.recordId;

            //log the request
            log.Info($"Account Name: {accountName} | Account number: {accountNumber} | EntityName: {entityName} | RecordId: {recordId}");

            //get the environment variables
            string user = Environment.GetEnvironmentVariable("user", EnvironmentVariableTarget.Process);
            string password = Environment.GetEnvironmentVariable("password", EnvironmentVariableTarget.Process);
            string baseUrl = Environment.GetEnvironmentVariable("sharepointBaseUrl", EnvironmentVariableTarget.Process);
            string crmUrl = Environment.GetEnvironmentVariable("crmurl", EnvironmentVariableTarget.Process);

            log.Info("Creating web");
            SPMethods.CreateWeb(user, password, baseUrl, accountNumber, accountName);

            log.Info("Creating web content");
            SPMethods.CreateWebContent(user, password, $"{baseUrl}/{accountNumber}");

            log.Info("Creating document locations");
            CRMMethods.CreateDocumentLocations(user, password, crmUrl, accountNumber, entityName, recordId);

            return req.CreateResponse(HttpStatusCode.OK, $"{accountName}-{accountNumber}");
        }
    }
}