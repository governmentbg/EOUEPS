using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace EPayments.Common.Helpers
{
    public class HmacRequestHelper
    {
        public static string CalculateHmac(string secretCode, string content)
        {
            string hmac;

            var secretCodeBytes = Encoding.UTF8.GetBytes(secretCode);
            var contentBytes = Encoding.UTF8.GetBytes(content);

            using (var hmacSHA256 = new HMACSHA256(secretCodeBytes))
            {
                var hash = hmacSHA256.ComputeHash(contentBytes);
                hmac = Convert.ToBase64String(hash);
            }

            return hmac;
        }

        public static Tuple<string, string> GetHmacAndCryptData(object data, string secretKey, bool returnUrlEncodedData)
        {
            var jsonData = JsonConvert.SerializeObject(data);

            var jsonDataBytes = Encoding.UTF8.GetBytes(jsonData);

            var base64Data = Convert.ToBase64String(jsonDataBytes);

            var hmac = CalculateHmac(secretKey, base64Data);

            if (returnUrlEncodedData)
            {
                return new Tuple<string,string>(HttpUtility.UrlEncode(hmac), HttpUtility.UrlEncode(base64Data));
            }
            else
            {
                return new Tuple<string,string>(hmac, base64Data);
            }
        }
    }
}
