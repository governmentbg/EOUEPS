using EPayments.Model.Models;
using System.Collections.Generic;
using System.Threading;

namespace EPayments.Distributions.Interfaces
{
    public interface IDistributionRevenueCreatable
    {
        List<DistributionRevenue> Distribute(CancellationToken token);
    }
}
