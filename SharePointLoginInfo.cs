using System;
using System.Runtime.Serialization;
using System.Configuration;

namespace Adil.DAL
{
    [DataContract]
    public sealed class SharePointLoginInfo
    {
        public Uri SiteURL
        {
            get
            {
                return new Uri(ConfigurationManager.AppSettings[Constant.Configuration.NeonPortalURL]);
            }
        }

        public Guid WebID { get; set; }

        public Guid ListID { get; set; }

        [DataMember]
        public string UserName { get; set; }

        public string DecryptedUserName 
        { 
            get
            {
                return RijndaelEncryptor.Decrypt(UserName);
            }
        }

        [DataMember]
        public string Password { get; set; }

        public string DecryptedPassword
        {
            get
            {
                return RijndaelEncryptor.Decrypt(Password);
            }
        }

        [DataMember]
        public string Domain { get; set; }

        public string UserLogin
        {
            get
            {
                return string.Format("{0}\\{1}", this.Domain, this.DecryptedUserName);
            }

        }
    }
}
