using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Service.Api.Controllers.v1.DataObjects
{
    internal class AcceptedReceiptJsonDO
    {
        public string Id { get; set; }
        public DateTime RegistrationTime { get; set; }
    }
}
