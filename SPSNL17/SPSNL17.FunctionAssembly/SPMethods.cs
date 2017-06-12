using Microsoft.SharePoint.Client;
using SPMeta2.CSOM.Services;
using SPMeta2.Definitions;
using SPMeta2.Enumerations;
using SPMeta2.Models;
using SPMeta2.Syntax.Default;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SPSNL17.FunctionAssembly
{
    public static class SPMethods
    {
        /// <summary>
        /// Creates the SharePoint web
        /// </summary>
        /// <param name="username">User name</param>
        /// <param name="password">Password</param>
        /// <param name="sitecollectionurl">The url of the site collection</param>
        /// <param name="urlPart">The sub url / url part for the web</param>
        /// <param name="title">The title for the web</param>
        public static void CreateWeb(string username, string password, string sitecollectionurl, string urlPart, string title)
        {
            using (var ctx = new ClientContext(sitecollectionurl))
            {
                //set the context credentials
                ctx.Credentials = new SharePointOnlineCredentials(username, GetSecurePassword(password));

                //create the web and set the properties
                WebCreationInformation information = new WebCreationInformation();
                information.WebTemplate = "STS#0";
                information.Description = title;
                information.Title = title;
                information.Url = urlPart;
                information.Language = 1033;

                //add the web to the site collection
                Web newWeb = null;
                newWeb = ctx.Web.Webs.Add(information);
                //execute the request
                ctx.ExecuteQuery();
            }
        }
        
        /// <summary>
        /// Deploys the SharePoint artefacts to a SharePoint a web
        /// </summary>
        /// <param name="username">User name</param>
        /// <param name="password">Password</param>
        /// <param name="webUrl">The url of the web</param>
        public static void CreateWebContent(string username, string password, string webUrl)
        {
            using (var ctx = new ClientContext(webUrl))
            {
                //set the context credentials
                ctx.Credentials = new SharePointOnlineCredentials(username, GetSecurePassword(password));

                //create a new SPMeta2 provisioning service
                var provisionService = new CSOMProvisionService();
                //deploy the model with SPMeta2
                provisionService.DeployWebModel(ctx, GetListModel());

                //delete the quick launch items
                DeleteQuickLaunchNodes(ctx);

                //then create the new quick launch items
                provisionService.DeployWebModel(ctx, GetNavigationModel());

                //set homepage
                ctx.Web.RootFolder.WelcomePage = "General";
                ctx.Web.RootFolder.Update();
                ctx.ExecuteQuery();
            }
        }

        #region models

        /// <summary>
        /// Web model for SharePoint
        /// </summary>
        /// <returns></returns>
        public static ModelNode GetListModel()
        {
            var model = SPMeta2Model
                .NewWebModel(web =>
                {
                    web
                        .AddList(GeneralDocuments)
                        .AddList(QuotesDocuments)
                        .AddList(ProjectDocuments);
                });

            return model;
        }

        /// <summary>
        /// List model for the General document library
        /// </summary>
        public static ListDefinition GeneralDocuments = new ListDefinition
        {
            Title = "General",
            Description = "Library to store general documents",
            TemplateType = BuiltInListTemplateTypeId.DocumentLibrary,
            CustomUrl = "General"
        };

        /// <summary>
        /// List model for the Quotes document library
        /// </summary>
        public static ListDefinition QuotesDocuments = new ListDefinition
        {
            Title = "Quotes",
            Description = "Library to store quote documents",
            TemplateType = BuiltInListTemplateTypeId.DocumentLibrary,
            CustomUrl = "Quotes"
        };

        /// <summary>
        /// List model for the Project document library
        /// </summary>
        public static ListDefinition ProjectDocuments = new ListDefinition
        {
            Title = "Projects",
            Description = "Library to store project documents",
            TemplateType = BuiltInListTemplateTypeId.DocumentLibrary,
            CustomUrl = "Projects"
        };

        /// <summary>
        /// The quick launch nodes model
        /// </summary>
        /// <returns></returns>
        public static ModelNode GetNavigationModel()
        {
            var model = SPMeta2Model
                .NewWebModel(web =>
                {
                    web
                        .AddQuickLaunchNavigationNode(
                            new QuickLaunchNavigationNodeDefinition
                            {
                                Title = GeneralDocuments.Title,
                                Url = GeneralDocuments.CustomUrl
                            }
                        )
                        .AddQuickLaunchNavigationNode(
                            new QuickLaunchNavigationNodeDefinition
                            {
                                Title = QuotesDocuments.Title,
                                Url = QuotesDocuments.CustomUrl
                            }
                        )
                        .AddQuickLaunchNavigationNode(
                            new QuickLaunchNavigationNodeDefinition
                            {
                                Title = ProjectDocuments.Title,
                                Url = ProjectDocuments.CustomUrl
                            }
                        );
                });

            return model;
        }

        #endregion


        #region helpers

        /// <summary>
        /// Deletes the quick launch nodes on a web
        /// </summary>
        /// <param name="ctx">SharePoint clientcontext</param>
        public static void DeleteQuickLaunchNodes(ClientContext ctx)
        {
            NavigationNodeCollection qlNodes = ctx.Web.Navigation.QuickLaunch;
            ctx.Load(qlNodes);
            ctx.ExecuteQuery();

            var nodes = qlNodes.ToList();
            for (var i = nodes.Count - 1; i > -1; i--)
            {
                nodes[i].DeleteObject();
            }
            ctx.ExecuteQuery();
        }


        /// <summary>
        /// Converts a string input to a secure string
        /// </summary>
        /// <param name="input">string value to be converted to a secure string</param>
        /// <returns>A secure string</returns>
        private static SecureString GetSecurePassword(string input)
        {
            var secure = new SecureString();
            foreach (var c in input)
            {
                secure.AppendChar(c);
            }

            return secure;
        }

        #endregion
    }
}
