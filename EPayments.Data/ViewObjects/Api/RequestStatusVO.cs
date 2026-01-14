using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.ViewObjects.Api
{
    public class RequestStatusVO
    {
        public string Id { get; set; }
        public PaymentRequestStatus Status { get; set; }
        public DateTime ChangeTime { get; set; }
    }
}
