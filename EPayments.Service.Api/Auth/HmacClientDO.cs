using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Service.Api.Auth
{
    public class HmacClientDO
    {
        public string ClientId { get; set; }
        public string SecretKey { get; set; }
        public bool IsActive { get; set; }
    }
}
