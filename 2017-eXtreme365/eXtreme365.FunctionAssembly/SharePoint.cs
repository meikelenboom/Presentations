using Microsoft.SharePoint.Client;
using SPMeta2.CSOM.Services;
using SPMeta2.Definitions;
using SPMeta2.Enumerations;
using SPMeta2.Models;
using SPMeta2.Syntax.Default;
using System.Linq;
using System.Security;

namespace eXtreme365.FunctionAssembly
{
    public class SharePoint
    {
        /// <summary>
        /// Creates a SharePoint web object in a site collection
        /// </summary>
        /// <param name="sitecollectionurl">The url of the root site collection</param>
        /// <param name="username">Username for authentication</param>
        /// <param name="password">Password for authentication</param
        /// <param name="title">The title of the web object</param>
        /// <param name="urlpart">The sub url for the web object</param>
        public static void CreateSharePointWeb(string sitecollectionurl, string username, string password, string title, string urlpart)
        {
            using (var ctx = new ClientContext(sitecollectionurl))
            {
                ctx.Credentials = new SharePointOnlineCredentials(username, GetSecurePassword(password));

                //create the web and set the properties
                WebCreationInformation information = new WebCreationInformation();
                information.WebTemplate = "STS#0";
                information.Description = title;
                information.Title = title;
                information.Url = urlpart;
                information.Language = 1033;

                //add the web to the site collection
                Web newWeb = null;
                newWeb = ctx.Web.Webs.Add(information);
                //make sure the web inherits permissions from its parent
                newWeb.ResetRoleInheritance();
                //execute the request
                ctx.ExecuteQuery();
            }
        }

        /// <summary>
        /// Creates the content for the web object. E.g. List, libraries, permissions etc.
        /// </summary>
        /// <param name="weburl">The url of the web object (sitecollection + url part)</param>
        /// <param name="username">Username for authentication</param>
        /// <param name="password">Password for authentication</param>
        public static void CreateSharePointWebContent(string weburl, string username, string password)
        {
            using (var ctx = new ClientContext(weburl))
            {
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

        #region SPMeta2 model
        /// <summary>
        /// The list model to be deployed to the web
        /// </summary>
        /// <returns>SPMeta2 ModelNode</returns>
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

        //below are the list definitions for the document libraries
        public static ListDefinition GeneralDocuments = new ListDefinition
        {
            Title = "General",
            Description = "Library to store general documents",
            TemplateType = BuiltInListTemplateTypeId.DocumentLibrary,
            CustomUrl = "General"
        };

        public static ListDefinition QuotesDocuments = new ListDefinition
        {
            Title = "Quotes",
            Description = "Library to store quote documents",
            TemplateType = BuiltInListTemplateTypeId.DocumentLibrary,
            CustomUrl = "Quotes"
        };

        public static ListDefinition ProjectDocuments = new ListDefinition
        {
            Title = "Projects",
            Description = "Library to store project documents",
            TemplateType = BuiltInListTemplateTypeId.DocumentLibrary,
            CustomUrl = "Projects"
        };

        /// <summary>
        /// The quick launch navigation model for the web
        /// </summary>
        /// <returns>SPMeta2 ModelNode</returns>
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

        /// <summary>
        /// Converts a string to a secure string to be used for SharePoint authentication
        /// </summary>
        /// <param name="input">String to be converted</param>
        /// <returns>Inpute converted into a secure string</returns>
        private static SecureString GetSecurePassword(string input)
        {
            var secure = new SecureString();
            foreach (var c in input)
            {
                secure.AppendChar(c);
            }

            return secure;
        }

        /// <summary>
        /// Iterates over all Quick Launch nodes on a web and deletes them
        /// </summary>
        /// <param name="ctx">SharePoint context</param>
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


    }
}
