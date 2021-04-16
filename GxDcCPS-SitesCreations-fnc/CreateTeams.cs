using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using RoleDefinition = Microsoft.SharePoint.Client.RoleDefinition;
using RoleAssignment = Microsoft.SharePoint.Client.RoleAssignment;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Web;
using System.IO;
using System.Web.Script.Serialization;
using System.Configuration;
using System.Text.RegularExpressions;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using Newtonsoft.Json;
using File = Microsoft.SharePoint.Client.File;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Reflection;
using System.Linq;

namespace GxDcCPSSitesCreationsfnc
{
    public static class CreateTeams
    {

        //static string siteId = "dc4145ad-428b-4e9f-a411-627ab525c06b,4b447531-c890-4345-b486-fdc317b95e03";
        //static string listId = "2ef1680d-b577-4ec1-9332-b6cc3ffc306a";
        static string servicePrincipalName = "Graph get teams info";
        static string adminSiteUrl = "https://tbssctdev-admin.sharepoint.com/";
        static string hubSiteUrl = "https://tbssctdev.sharepoint.com/sites/GCXCollab";
        
        static string siteRelativePath = "teams/scw";
        static string hostname = "tbssctdev.sharepoint.com";
        static string listTitle = "space requests";

        static string CLIENT_ID = ConfigurationManager.AppSettings["CLIENT_ID"];
        static string CLIENT_SECERET = ConfigurationManager.AppSettings["CLIENT_SECRET"];
        static string appOnlyId = ConfigurationManager.AppSettings["AppOnlyID"];
        static string appOnlySecret = ConfigurationManager.AppSettings["AppOnlySecret"];
        [FunctionName("CreateTeams")]
        public static async void Run([QueueTrigger("create-teams", Connection = "")] SiteInfo myQueueItem, TraceWriter log)
        {
            log.Info($"C# Queue trigger function processed: {myQueueItem.displayName}");

            var itemId = myQueueItem.itemId;
            var siteUrl = myQueueItem.siteUrl;
            var groupId = myQueueItem.groupId;
            var displayName = myQueueItem.displayName;
            var emails = myQueueItem.emails;
            var requesterName = myQueueItem.requesterName;
            var requesterEmail = myQueueItem.requesterEmail;
            var authResult = GetOneAccessToken();
            var graphClient = GetGraphClient(authResult);
            
            var siteId = GetSiteId(graphClient, log, siteRelativePath, hostname).GetAwaiter().GetResult();
            var listId = GetSiteListId(graphClient, siteId, listTitle).GetAwaiter().GetResult();
            
            var teamsUrl = CreateSCWTeams(graphClient, log, groupId).GetAwaiter().GetResult();
            if (teamsUrl == "")
            {
                UpdateStatus(graphClient, log, itemId, "Team Creation Failed", siteId, listId);
                throw new ApplicationException("Failure");
            }
            RemoveServicePrincipalFromOwnerGroup(graphClient, groupId, servicePrincipalName).GetAwaiter().GetResult();
            AddUserToGroup(graphClient, log, groupId, emails).GetAwaiter().GetResult();
            RemoveTeamCreatorUser(graphClient, log, groupId).GetAwaiter().GetResult();

            ClientContext ctx = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(siteUrl, appOnlyId, appOnlySecret);

            UpdateNavigation(ctx, log, teamsUrl, displayName, "microsoft teams", 4);
            UpdateNavigation(ctx, log, teamsUrl, displayName, "request to join", 5);
            RemoveSitePage(ctx, log);
            RemoveUserFromSiteAdmin(ctx, log, groupId);

            var status = "Team Created";
            //   UpdateStatus(graphClient, log, itemId, "Team Created");
            UpdateStatus(graphClient, log, itemId, status, siteId, listId);

            ClientContext adminCtx = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(adminSiteUrl, appOnlyId, appOnlySecret);
            AddTopNavigationBar(adminCtx, log, siteUrl, hubSiteUrl);

            
            //send message to queue
            var connectionString = ConfigurationManager.AppSettings["AzureWebJobsStorage"];
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
            CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
            CloudQueue queue = queueClient.GetQueueReference("email-info");
            InsertMessageAsync(queue, siteUrl, displayName, emails, status, requesterName, requesterEmail, log).GetAwaiter().GetResult();
            log.Info($"Sent request to queue successful.");

        }
        /// <summary>
        /// This method will delete service principal from owner group of office 365 group
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="groupId"></param>
        /// <param name="servicePrincipalName"></param>
        /// <returns></returns>
        public static async Task RemoveServicePrincipalFromOwnerGroup(GraphServiceClient graphClient, string groupId, string servicePrincipalName)
        {
            var servicePrincipals = await graphClient.ServicePrincipals
                                                    .Request()
                                                    .Select("id")
                                                    .Filter($@"displayName eq '{servicePrincipalName}'")
                                                    .GetAsync();
            var servicePrincipalId = "";
            foreach (var item in servicePrincipals)
            {
                servicePrincipalId = item.Id;
                break;
            }

            await graphClient.Groups[groupId].Owners[servicePrincipalId].Reference
                            .Request()
                            .DeleteAsync();
        }
        /// <summary>
        /// Send message to queue.
        /// </summary>
        /// <param name="theQueue"></param>
        /// <param name="siteUrl"></param>
        /// <param name="displayName"></param>
        /// <param name="emails"></param>
        /// <param name="status"></param>
        /// <param name="requesterName"></param>
        /// <param name="requesterEmail"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        public static async Task InsertMessageAsync(CloudQueue theQueue, string siteUrl, string displayName, string emails, string status, string requesterName, string requesterEmail, TraceWriter log)
        {
            SiteInfo siteInfo = new SiteInfo();

            siteInfo.siteUrl = siteUrl;
            siteInfo.status = status;
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
        /// Get access token from AAD
        /// </summary>
        /// <returns></returns>
        public static string GetOneAccessToken()
        {
            string token = "";

            string TENAT_ID = "ddbd240e-11ba-47a6-abeb-e1a6be847a17";
            string TOKEN_ENDPOINT = "";
            string MS_GRAPH_SCOPE = "";
            string GRANT_TYPE = "";

            try
            {

                TOKEN_ENDPOINT = "https://login.microsoftonline.com/" + TENAT_ID + "/oauth2/v2.0/token";
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
        /// <summary>
        /// Get graph client
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
        /// Update navigation of SharePoint site with teams url.
        /// </summary>
        /// <param name="contextNavigation"></param>
        /// <param name="log"></param>
        /// <param name="teamsUrl"></param>
        /// <param name="displayName"></param>
        /// <param name="navigationNode"></param>
        /// <param name="navigationIndex"></param>
        public static void UpdateNavigation(ClientContext contextNavigation, TraceWriter log, string teamsUrl, string displayName, string navigationNode, int navigationIndex)
        {
            System.Collections.Generic.List<int> navIds = new System.Collections.Generic.List<int>();
            Web web = contextNavigation.Web;
            contextNavigation.Load(web, w => w.Navigation);
            contextNavigation.Load(web);
            contextNavigation.ExecuteQuery();

            //Start working on Navigation
            NavigationNodeCollection lefthandNav = web.Navigation.QuickLaunch;
            contextNavigation.Load(lefthandNav);
            contextNavigation.ExecuteQuery();
            NavigationNodeCreationInformation nodeToCreate = new NavigationNodeCreationInformation();
            NavigationNode navNode = lefthandNav[navigationIndex];
            contextNavigation.Load(navNode);
            contextNavigation.ExecuteQuery();

            if (navNode.Title.ToLower().Contains(navigationNode))
            {
                nodeToCreate.PreviousNode = navNode;
                nodeToCreate.Title = navNode.Title;
                nodeToCreate.Url = teamsUrl;
                nodeToCreate.IsExternal = true;
                navIds.Add(navNode.Id); //For deleting the existing node after creating the new one
            }

            if (nodeToCreate.Title != "")
            {
                lefthandNav.Add(nodeToCreate);
            }

            contextNavigation.ExecuteQuery();

            foreach (int id in navIds)
            {
                NavigationNode nodeToDelete = web.Navigation.GetNodeById(id);
                contextNavigation.Load(nodeToDelete);
                contextNavigation.ExecuteQuery();
                nodeToDelete.DeleteObject();
                contextNavigation.ExecuteQuery();
            }
            log.Info($"Update navigation node {navigationNode} successfully.");

        }
        /// <summary>
        /// Remove SharePoint owners group from SharePoint admin collection
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="log"></param>
        /// <param name="groupId"></param>
        public static void RemoveUserFromSiteAdmin(ClientContext ctx, TraceWriter log, string groupId)
        {
            var user = ctx.Site.RootWeb.SiteUsers.GetByLoginName($"c:0o.c|federateddirectoryclaimprovider|{groupId}_o");
            ctx.Load(user);
            ctx.ExecuteQuery();
            user.IsSiteAdmin = false;
            user.Update();
            ctx.Load(user);
            ctx.ExecuteQuery();
            log.Info("Remove owner group from site admin successful.");
        }
        /// <summary>
        /// Remove duplicate page from site pages library.
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="log"></param>
        public static void RemoveSitePage(ClientContext ctx, TraceWriter log)
        {
            Microsoft.SharePoint.Client.List sitePagesList = ctx.Web.Lists.GetByTitle("Site pages");
            Microsoft.SharePoint.Client.ListItem item = sitePagesList.GetItemById(3);
            File targetFile = item.File;
            targetFile.DeleteObject();
            ctx.ExecuteQuery();
            log.Info("Remove site page successful.");
        }
        /// <summary>
        /// Add top navigation bar to SharePoint Site.
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="log"></param>
        /// <param name="targetSiteUrl"></param>
        /// <param name="hubSiteUrl"></param>
        private static void AddTopNavigationBar(ClientContext ctx, TraceWriter log, string targetSiteUrl, string hubSiteUrl)
        {
            var tenant = new Tenant(ctx);
            tenant.ConnectSiteToHubSite(targetSiteUrl, hubSiteUrl);
            ctx.ExecuteQuery();
        }
        /// <summary>
        /// Create Teams
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="log"></param>
        /// <param name="groupId"></param>
        /// <returns></returns>
        public static async Task<string> CreateSCWTeams(GraphServiceClient graphClient, TraceWriter log, string groupId)
        {
            var teamsUrl = "";
            try
            {
                var team = new Team
                {
                    MemberSettings = new TeamMemberSettings
                    {
                        AllowCreateUpdateChannels = true,
                        ODataType = null
                    },
                    MessagingSettings = new TeamMessagingSettings
                    {
                        AllowUserEditMessages = true,
                        AllowUserDeleteMessages = true,
                        ODataType = null
                    },
                    FunSettings = new TeamFunSettings
                    {
                        AllowGiphy = true,
                        GiphyContentRating = GiphyRatingType.Strict,
                        ODataType = null
                    },
                    ODataType = null,
                };

                var createdTeam = await graphClient.Groups[groupId].Team
                 .Request()
                 .PutAsync(team);


                //get channel Id
                var channels = await graphClient.Teams[createdTeam.Id].Channels
                        .Request()
                        .GetAsync();
                var channelId = "";
                foreach (var channel in channels)
                {
                    channelId = channel.Id;
                }
                teamsUrl = $@"https://teams.microsoft.com/l/team/{channelId}/conversations?groupId={createdTeam.Id}&tenantId=ddbd240e-11ba-47a6-abeb-e1a6be847a17";
                log.Info($"Teams created successfully. Team url is {teamsUrl}");
            }
            catch (Exception ex)
            {
                log.Info($"create teams error is: {ex}");
            }
            return teamsUrl;
        }
        /// <summary>
        /// Update status to "Teams Created" in SharePoint list
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="log"></param>
        /// <param name="itemId"></param>
        public static async void UpdateStatus(GraphServiceClient graphClient, TraceWriter log, string itemId, string status, string siteId, string listId)
        {




            var fieldValueSet = new FieldValueSet();
            var field = new Dictionary<string, object>()
                              {
                                {"_Status", status },
                              };
            fieldValueSet.AdditionalData = field;
            var result = await graphClient.Sites[siteId].Lists[listId].Items[itemId].Fields
                .Request()
                .UpdateAsync(fieldValueSet);
            log.Info("Update status successfully.");
        }

        //Remove Licensed user teamcreator
        public static async Task RemoveTeamCreatorUser(GraphServiceClient graphClient, TraceWriter log, string groupId)
        {

            var Id = "552b16be-8a50-460c-ba24-907f45376ac1"; //teamcreator

            await graphClient.Groups[groupId].Owners[Id].Reference
                    .Request()
                    .DeleteAsync();
            log.Info($"Licensed was remove from owner of {groupId} successfully.");

            var conversationMember = new AadUserConversationMember
            {
                Roles = new List<String>()
                {
                    ""
                }
            };
        }

        /// <summary>
        /// This method will add users to group as owners and members.
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="log"></param>
        /// <param name="groupId"></param>
        /// <param name="emails"></param>
        public static async Task AddUserToGroup(GraphServiceClient graphClient, TraceWriter log, string groupId, string emails)
        {
            List<string> emailList = emails.Split(',').ToList<string>();
            try
            {


                foreach (var i in emailList)
                {
                    var userId = "";
                    IGraphServiceUsersCollectionPage users;

                    if (i.Contains($"#EXT#"))
                    {
                        var str = i.Remove(i.IndexOf("#"));
                        users = await graphClient.Users
                                .Request()
                                .Filter($@"startsWith(userPrincipalName, '{str}')")
                                .Select("id")
                                .GetAsync();
                    }
                    else
                    {
                        users = await graphClient.Users
                        .Request()
                        .Filter($@"mail eq '{i}'")
                        .Select("id")
                        .GetAsync();
                    }

                    foreach (var user in users)
                    {
                        userId = user.Id;
                    }
                    var directoryObject = new DirectoryObject
                    {
                        Id = userId
                    };

                    await graphClient.Groups[groupId].Owners.References
                        .Request()
                        .AddAsync(directoryObject);
                    log.Info($"{i} add to owner of {groupId} successfully.");

                    await graphClient.Groups[groupId].Members.References
                        .Request()
                        .AddAsync(directoryObject);
                    log.Info($"{i} add to member of {groupId} successfully.");
                }
            }
            catch (Exception e)
            {
                log.Info($"error message is : {e}");
            }
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
    }
}
