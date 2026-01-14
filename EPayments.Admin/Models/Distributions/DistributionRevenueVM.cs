using EPayments.Admin.Models.Shared;
using EPayments.Data.ViewObjects.Admin;
using System.Collections.Generic;

namespace EPayments.Admin.Models.Distributions
{
    public class DistributionRevenueVM
    {
        public List<DistributionRevenueVO> DistributionRevenues { get; set; }

        public PagingVM RequestsPagingOptions { get; set; }

        public DistributionRevenueSearchDO SearchDO { get; set; }

        public List<DistribtutionTypeVO> DistribtutionTypes { get; set; }
    }
}