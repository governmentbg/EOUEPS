using System;
using System.Collections.Generic;

namespace EPayments.Distributions.Models
{
    public class DistributionRevenueAgency
    {
        public string Agency { get; set; }

        public DateTime DistributedDate { get; set; }

        public string DistributedDateAsString
        {
            get
            {
                return this.DistributedDate.ToString("yyyyMMdd");
            }
        }

        public int Num { get; set; }

        public decimal TotalSum { get; set; }

        public string DistributionType { get; set; }

        public ICollection<RecTransaction> RecTransactions { get; set; } = new HashSet<RecTransaction>();

        public ICollection<RecDistributedAmount> DistributedAmounts { get; set; } = new HashSet<RecDistributedAmount>();
    }
}
