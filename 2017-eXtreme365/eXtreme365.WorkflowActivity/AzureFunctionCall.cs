using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.IO;
using System.Net;
using System.Text;

namespace eXtreme365.WorkflowActivity
{
    public class AzureFunctionCall : CodeActivity
    {
        [RequiredArgument]
        [Input("Azure Function URL")]
        [Default("")]
        public InArgument<String> FunctionUrl { get; set; }

        [RequiredArgument]
        [Input("SharePoint Site Collection Url")]
        [Default("")]
        public InArgument<String> SiteUrl { get; set; }

        [RequiredArgument]
        [Input("SharePoint Web Title")]
        [Default("")]
        public InArgument<String> WebTitle { get; set; }

        [RequiredArgument]
        [Input("SharePoint Web Url Part")]
        [Default("")]
        public InArgument<String> WebUrlPart { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                //get the input parameters
                var functionUrl = this.FunctionUrl.Get(executionContext);
                var siteUrl = this.SiteUrl.Get(executionContext);
                var webTitle = this.WebTitle.Get(executionContext);
                var webUrlPart = this.WebUrlPart.Get(executionContext);

                //get entity name and record id
                var recordId = context.PrimaryEntityId.ToString();
                var entityName = context.PrimaryEntityName;

                //creating the json post data by parsing a string
                var postdata = string.Format("{{ \"url\" : \"{0}\", \"title\" : \"{1}\", \"urlpart\" : \"{2}\", \"recordid\" : \"{3}\", \"entityname\" : \"{4}\"  }}", 
                    siteUrl, webTitle, webUrlPart, recordId, entityName);
                //convert the data to a byte array for the http post call
                var data = Encoding.ASCII.GetBytes(postdata);

                //setting the web request properties
                WebRequest request = WebRequest.Create(functionUrl);
                request.Method = "POST";
                request.ContentType = "application/json";
                request.ContentLength = data.Length;

                //writing the request stream
                using (var stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }

                //reading the response
                try
                {
                    var response = (HttpWebResponse)request.GetResponse();
                    var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                }
                catch (WebException webex)
                {
                    WebResponse errResp = webex.Response;
                    using (Stream respStream = errResp.GetResponseStream())
                    {
                        StreamReader reader = new StreamReader(respStream);
                        string text = reader.ReadToEnd();

                        throw new InvalidPluginExecutionException(text);
                    }
                }

            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }
    }
}
