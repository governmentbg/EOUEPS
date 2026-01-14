using System;

namespace EPayments.Distributions.Models
{
    public class RecTransaction
    {
        public DateTime TransactionDate { get; set; }

        public string AuthorizationCode { get; set; }

        public string Tid { get; set; }

        public string RRN { get; set; }

        public string OrderId { get; set; }

        public decimal Sum { get; set; }
    }
}
