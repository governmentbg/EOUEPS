using EPayments.Data.ViewObjects.Admin;
using System.Collections.Generic;

namespace EPayments.Admin.Models.PaymentRequests
{
    public class PrintPaymentRequest
    {
        public List<PaymentRequestVO> PaymentRequests { get; set; } = new List<PaymentRequestVO>();
    }
}