using System;
namespace EPayments.Common.DataObjects
{
    public class AuthRequestDO
    {
        public string ClientId { get; set; }
        public string Hmac { get; set; }
        public string Data { get; set; }
    }
}
