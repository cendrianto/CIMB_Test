using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Configuration;
using System.Text;
using System.Security.Cryptography;

namespace CIMB_trial
{
    public class security
    {
        public static string MD5hash(string s)
        {
            string cryptoKey = WebConfigurationManager.AppSettings["PwdCryptoKey"];
            StringBuilder hash = new StringBuilder();
            MD5CryptoServiceProvider md5prov = new MD5CryptoServiceProvider();
            byte[] bytes = md5prov.ComputeHash(new UTF8Encoding().GetBytes(s));

            for (int i = 0; i < bytes.Length; i++)
            {
                hash.Append(bytes[i].ToString(cryptoKey));
            }
            return hash.ToString();
        }
    }
}