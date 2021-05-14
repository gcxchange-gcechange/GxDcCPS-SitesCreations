using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GxDcCPSSitesCreationsfnc
{
    public class LgObject
    {
        public string token_type { get; set; }
        public string scope { get; set; }
        public string expires_in { get; set; }
        public string ext_expires_in { get; set; }
        public string access_token { get; set; }
    }

    public class SiteInfo
    {
        public string itemId { get; set; }
        public string siteUrl { get; set; }
        public string groupId { get; set; }
        public string displayName { get; set; }

        public string emails { get; set; }
        public string comments { get; set; }

        public string status { get; set; }

        public string requesterName { get; set; }
        public string requesterEmail { get; set; }
    }
    public class CCApplication
    {
        public string name { get; set; }
        public string description { get; set; }
        public string mailNickname { get; set; }
        public string itemId { get; set; }
        public string emails { get; set; }
        public string requesterName { get; set; }
        public string requesterEmail { get; set; }
    }

}
