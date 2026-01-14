using EPayments.Admin.Models.Shared;
using EPayments.Data.ViewObjects.Admin;
using System.Collections.Generic;

namespace EPayments.Admin.Models.Distributions
{
    public class PaymentVM
    {
        public DistributionRevenueVO DistributionRevenue { get; set; }

        public List<DistributedPaymentRequestVO> Payments { get; set; }

        public List<DistribtutionTypeVO> DistribtutionTypes { get; set; }

        public PaymentSearchDO SearchDO { get; set; }
    }
}