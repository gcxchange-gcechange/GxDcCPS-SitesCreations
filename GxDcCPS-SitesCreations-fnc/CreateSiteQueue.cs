using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Script.Serialization;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using Newtonsoft.Json;

namespace GxDcCPSSitesCreationsfnc
{
    public static class CreateSiteQueue
    {
        [FunctionName("CreateSiteQueue")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            // parse query parameter
            string name = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "name", true) == 0)
                .Value;
            string description = "";
            string mailNickname = "";
            string itemId = "";
            string emails = "";
            string requesterName = "";
            string requesterEmail = "";
            string TENANT_NAME = ConfigurationManager.AppSettings["TENANT_NAME"];


            string targetSiteUrl = $"https://{TENANT_NAME}.sharepoint.com/teams/{mailNickname}";
            if (name == null)
            {
                // Get request body
                dynamic data = await req.Content.ReadAsAsync<object>();
                name = data?.name;
                description = data?.description;
                mailNickname = data?.mailNickname;
                itemId = data?.itemId;
                emails = data?.emails;
                requesterName = data?.requesterName;
                requesterEmail = data?.requesterEmail;
            }

            //send message to queue
            var connectionString = ConfigurationManager.AppSettings["AzureWebJobsStorage"];
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
            CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
            CloudQueue queue = queueClient.GetQueueReference("ccapplication");
            InsertMessageAsync(queue, itemId, name, description, mailNickname, emails, requesterName, requesterEmail, log).GetAwaiter().GetResult();
            log.Info($"Sent request to queue successful.");

            return name == null
                ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
                : req.CreateResponse(HttpStatusCode.OK, "Hello " + name);
        }

        private static async Task InsertMessageAsync(CloudQueue theQueue, string itemId, string dispalyName, string description, string mailNickname, string emails, string requesterName, string requesterEmail, TraceWriter log)
        {
            CCApplication siteInfo = new CCApplication();

            siteInfo.itemId = itemId;
            siteInfo.name = dispalyName;
            siteInfo.description = description;
            siteInfo.mailNickname = mailNickname;
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
    }
}
