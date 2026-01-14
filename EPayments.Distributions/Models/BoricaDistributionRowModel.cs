using System;

namespace EPayments.Distributions.Models
{
    public class BoricaDistributionRowModel
    {
        public Guid Gid { get; set; }

        public string AisName { get; set; }

        public string AccountBank { get; set; }

        public string AccountBIC { get; set; }

        public string AccountIBAN { get; set; }

        public decimal PaymentAmount { get; set; }

        public int DistributionType { get; set; }
    }
}
