using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Threading;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using System.IO;
using System.Web;
using System.Web.Script.Serialization;
using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Configuration;
using Newtonsoft.Json;
using Microsoft.WindowsAzure.Storage.Queue;
using Microsoft.WindowsAzure.Storage;
using System.Reflection;

namespace GxDcCPSSitesCreationsfnc
{
    public static class CreateSite
    {

        static readonly string PNP_TEMPLATE_FILE = "template.xml";

        static string siteRelativePath = "teams/scw";
       
        static string listTitle = "space requests";
       
        static string appOnlyId = ConfigurationManager.AppSettings["AppOnlyID"];
        static string appOnlySecret = ConfigurationManager.AppSettings["AppOnlySecret"];
        static string CLIENT_ID = ConfigurationManager.AppSettings["CLIENT_ID"];
        static string CLIENT_SECERET = ConfigurationManager.AppSettings["CLIENT_SECRET"];
        

        [FunctionName("CreateSite")]
        public static void Run([QueueTrigger("ccapplication", Connection = "")] CCApplication myQueueItem,TraceWriter log, Microsoft.Azure.WebJobs.ExecutionContext functionContext)
        {
            try
            {
                log.Info($"C# Queue trigger function processed: {myQueueItem.name}");
                string TENANT_ID = ConfigurationManager.AppSettings["TENANT_ID"];
                string TENANT_NAME = ConfigurationManager.AppSettings["TENANT_NAME"];
                string hostname = $"{TENANT_NAME}.sharepoint.com";
                string TEAMS_INIT_USERID = ConfigurationManager.AppSettings["TEAMS_INIT_USERID"];
                var displayName = myQueueItem.name;
                var description = myQueueItem.description;
                var mailNickname = myQueueItem.mailNickname;
                var itemId = myQueueItem.itemId;
                var emails = myQueueItem.emails;
                var requesterName = myQueueItem.requesterName;
                var requesterEmail = myQueueItem.requesterEmail;

                string targetSiteUrl = $"https://{TENANT_NAME}.sharepoint.com/teams/{mailNickname}";

                var authResult = GetOneAccessToken(TENANT_ID);
                var graphClient = GetGraphClient(authResult);

                var siteId = GetSiteId(graphClient, log, siteRelativePath, hostname).GetAwaiter().GetResult();
                var listId = GetSiteListId(graphClient, siteId, listTitle).GetAwaiter().GetResult();

                var groupId = CreateGroupAndSite(graphClient, log, description, displayName, mailNickname).GetAwaiter().GetResult();
                log.Info($"Group id is {groupId}");
                AddLicensedUserToGroup(graphClient, log, groupId, TEAMS_INIT_USERID);

                log.Info("Wait 3 minutes for site setup.");
                Thread.Sleep(3 * 60 * 1000);

                var siteDescriptions = GetSiteDescriptions(graphClient, siteId, listId, itemId).GetAwaiter().GetResult();

                ClientContext ctx = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(targetSiteUrl, appOnlyId, appOnlySecret);
                ApplyProvisioningTemplate(ctx, log, functionContext, siteDescriptions);
                UpdateStatus(graphClient, log, itemId, siteId, listId);

                //send message to create-tems queue
                var connectionString = ConfigurationManager.AppSettings["AzureWebJobsStorage"];
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
                CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                CloudQueue queue = queueClient.GetQueueReference("create-teams");
                InsertMessageAsync(queue, itemId, targetSiteUrl, groupId, displayName, emails, requesterName, requesterEmail, log).GetAwaiter().GetResult();
                log.Info($"Sent request to queue successful.");
            }
            catch (Exception ex)
            {
                log.Info($"error message: {ex}");
            }

        }
        /// <summary>
        /// This method will send message to queue.
        /// </summary>
        /// <param name="theQueue"></param>
        /// <param name="itemId"></param>
        /// <param name="siteUrl"></param>
        /// <param name="groupId"></param>
        /// <param name="displayName"></param>
        /// <param name="emails"></param>
        /// <param name="requesterName"></param>
        /// <param name="requesterEmail"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        public static async Task InsertMessageAsync(CloudQueue theQueue, string itemId, string siteUrl, string groupId, string displayName, string emails, string requesterName, string requesterEmail, TraceWriter log)
        {
            SiteInfo siteInfo = new SiteInfo();
            siteInfo.itemId = itemId;
            siteInfo.siteUrl = siteUrl;
            siteInfo.groupId = groupId;
            siteInfo.displayName = displayName;
            siteInfo.emails = emails;
            siteInfo.requesterName = requesterName;
            siteInfo.requesterEmail = requesterEmail;

            string serializedMessage = JsonConvert.SerializeObject(siteInfo);
            if (await theQueue.CreateIfNotExistsAsync())
            {
                log.Info("The queue was created.");
            }

            CloudQueueMessage message = new CloudQueueMessage(serializedMessage);
            await theQueue.AddMessageAsync(message);
        }
        /// <summary>
        /// This method will get graph client.
        /// </summary>
        /// <param name="authResult"></param>
        /// <returns></returns>
        public static GraphServiceClient GetGraphClient(string authResult)
        {
            GraphServiceClient graphClient = new GraphServiceClient(
                 new DelegateAuthenticationProvider(
            async (requestMessage) =>
            {
                requestMessage.Headers.Authorization =
                    new AuthenticationHeaderValue("bearer",
                    authResult);
            }));
            return graphClient;
        }
        /// <summary>
        /// This method will create an Office 365 group
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="log"></param>
        /// <param name="description"></param>
        /// <param name="displayName"></param>
        /// <param name="mailNickname"></param>
        /// <returns></returns>
        public static async Task<string> CreateGroupAndSite(GraphServiceClient graphClient, TraceWriter log, string description, string displayName, string mailNickname)
        {

            var o365Group = new Microsoft.Graph.Group
            {
                Description = description,
                DisplayName = $@"{displayName}",
                GroupTypes = new List<String>()
                    {
                        "Unified"
                    },
                MailEnabled = true,
                MailNickname = mailNickname,
                SecurityEnabled = false,
                Visibility = "Private"

            };

            var result = await graphClient.Groups
            .Request()
            .AddAsync(o365Group);
            log.Info($"Site and Office 365 {displayName} created successfully.");

            return result.Id;
        }
        /// <summary>
        /// This method will update requests status in SharePoint list.
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="log"></param>
        /// <param name="itemId"></param>
        public static async void UpdateStatus(GraphServiceClient graphClient, TraceWriter log, string itemId, string siteId, string listId)
        {
            var fieldValueSet = new FieldValueSet();
            var field = new Dictionary<string, object>()
                              {
                                {"_Status", "Site Created" },
                              };
            fieldValueSet.AdditionalData = field;
            var result = await graphClient.Sites[siteId].Lists[listId].Items[itemId].Fields
                .Request()
                .UpdateAsync(fieldValueSet);
            log.Info("Update status successfully.");
        }
        /// <summary>
        /// This method will apply PNP template to a SharePoint site.
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="log"></param>
        /// <param name="functionContext"></param>
        public static void ApplyProvisioningTemplate(ClientContext ctx, TraceWriter log, Microsoft.Azure.WebJobs.ExecutionContext functionContext, string description)
        {    try
            {
            ctx.RequestTimeout = Timeout.Infinite;
            Web web = ctx.Web;
            ctx.Load(web, w => w.Title);
            ctx.ExecuteQueryRetry();

            log.Info($"Successfully connected to site: {web.Title}");

            DirectoryInfo dInfo;
            var schemaDir = "";
            string currentDirectory = functionContext.FunctionDirectory;
            if (currentDirectory == null)
            {
                string workingDirectory = Environment.CurrentDirectory;
                currentDirectory = System.IO.Directory.GetParent(workingDirectory).Parent.Parent.FullName;
                dInfo = new DirectoryInfo(currentDirectory);
                schemaDir = dInfo + "\\GxDcCPS-SitesCreations-fnc\\bin\\Debug\\net461\\Templates\\GenericTemplatev2";
            }
            else
            {
                dInfo = new DirectoryInfo(currentDirectory);
                schemaDir = dInfo.Parent.FullName + "\\Templates\\GenericTemplatev2";
            }

            log.Info($"schemaDir is {schemaDir}");
            XMLTemplateProvider sitesProvider = new XMLFileSystemTemplateProvider(schemaDir, "");
        
            ProvisioningTemplate template = sitesProvider.GetTemplate(PNP_TEMPLATE_FILE); 
            log.Info($"Successfully found template with ID '{template.Id}'");
          
      

            ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation
            {
                ProgressDelegate = (message, progress, total) =>
                {
                    log.Info(string.Format("{0:00}/{1:00} - {2} : {3}", progress, total, message, web.Title));
                }
            };
            FileSystemConnector connector = new FileSystemConnector(schemaDir, "");

            template.Connector = connector;

            string[] descriptions = description.Split('|');

            template.Parameters.Add("descEN", descriptions[0]);
            template.Parameters.Add("descFR", descriptions[1]);

            web.ApplyProvisioningTemplate(template, ptai);

            log.Info($"Site {web.Title} apply template successfully.");      
              }
            catch (ReflectionTypeLoadException ex)
            {
                foreach (var item in ex.LoaderExceptions)
                {
                    log.Info(item.Message);
                }
            }
        }
        /// <summary>
        /// This method will get AAD access token.
        /// </summary>
        /// <returns></returns>
        public static string GetOneAccessToken(string TENANT_ID)
        {
            string token = "";

            string TOKEN_ENDPOINT = "";
            string MS_GRAPH_SCOPE = "";
            string GRANT_TYPE = "";

            try
            {
                TOKEN_ENDPOINT = "https://login.microsoftonline.com/" + TENANT_ID + "/oauth2/v2.0/token";
                MS_GRAPH_SCOPE = "https://graph.microsoft.com/.default";
                GRANT_TYPE = "client_credentials";
            }
            catch (Exception e)
            {
                Console.WriteLine("A error happened while search config file");
            }
            try
            {
                HttpWebRequest request = WebRequest.Create(TOKEN_ENDPOINT) as HttpWebRequest;
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                StringBuilder data = new StringBuilder();
                data.Append("client_id=" + HttpUtility.UrlEncode(CLIENT_ID));
                data.Append("&scope=" + HttpUtility.UrlEncode(MS_GRAPH_SCOPE));
                data.Append("&client_secret=" + HttpUtility.UrlEncode(CLIENT_SECERET));
                data.Append("&GRANT_TYPE=" + HttpUtility.UrlEncode(GRANT_TYPE));
                byte[] byteData = UTF8Encoding.UTF8.GetBytes(data.ToString());
                request.ContentLength = byteData.Length;
                using (Stream postStream = request.GetRequestStream())
                {
                    postStream.Write(byteData, 0, byteData.Length);
                }

                // Get response
                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
                {

                    using (var reader = new StreamReader(response.GetResponseStream()))
                    {
                        JavaScriptSerializer js = new JavaScriptSerializer();
                        var objText = reader.ReadToEnd();
                        LgObject myojb = (LgObject)js.Deserialize(objText, typeof(LgObject));
                        token = myojb.access_token;
                    }

                }
                return token;
            }
            catch (Exception e)
            {
                Console.WriteLine("A error happened while connect to server please check config file");
                return "error";
            }
        }

        //Add licensed user
        public static async void AddLicensedUserToGroup(GraphServiceClient graphClient, TraceWriter log, string groupId, string TEAMS_INIT_USERID)
        {
            var directoryObject = new DirectoryObject
            {
                Id = TEAMS_INIT_USERID //teamcreator
            };

            await graphClient.Groups[groupId].Owners.References
                  .Request()
                  .AddAsync(directoryObject);
            log.Info($"Licensed add to owner of {groupId} successfully.");
        }
        /// <summary>
        /// Get scw site id
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="log"></param>
        /// <param name="siteRelativePath"></param>
        /// <param name="hostname"></param>
        /// <returns></returns>
        public static async Task<string> GetSiteId(GraphServiceClient graphClient, TraceWriter log, string siteRelativePath, string hostname)
        {
            // get site id
            var site = await graphClient.Sites.GetByPath(siteRelativePath, hostname).Request().Select("id").GetAsync();
            var siteId = site.Id;
            var hostLength = hostname.Length;
            return siteId = siteId.Remove(0, hostLength + 1);
        }

        /// <summary>
        /// get space requests list id
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="siteId"></param>
        /// <param name="listTitle"></param>
        /// <returns></returns>
        public static async Task<string> GetSiteListId(GraphServiceClient graphClient, string siteId, string listTitle)
        {
            //get list id
            var lists = await graphClient.Sites[siteId].Lists.Request()
                                                    .Select("id")
                                                    .Filter($@"displayName eq '{listTitle}'")
                                                    .GetAsync();
            var listId = "";
            foreach (var list in lists)
            {
                listId = list.Id;
                break;
            }
            return listId;
        }

        /// <summary>
        /// Get English | French description from space requests
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="siteId"></param>
        /// <param name="listId"></param>
        /// <param name="itemId"></param>
        /// <returns></returns>
        public static async Task<string> GetSiteDescriptions(GraphServiceClient graphClient, string siteId, string listId, string itemId)
        {

            var items = await graphClient.Sites[siteId].Lists[listId].Items[itemId].Fields
                .Request()
                .Select("Space_x0020_Description_x0020__x, Space_x0020_Description_x0020__x0")
                .GetAsync();

            var descEn = items.AdditionalData["Space_x0020_Description_x0020__x"];
            var descfr = items.AdditionalData["Space_x0020_Description_x0020__x0"];

            return descEn + "|" + descfr;
        }
    }
}
