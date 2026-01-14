using EPayments.CVPosTransaction.CVPosService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.CVPosTransaction.Manager
{
    public interface ICVPosRegisterManager
    {
        Task<transactionsForDateResponse> GetTransactionPerDateAsync(string agency, string @event, DateTime dateEvent, string tid);
        recEventTransaction[] GetTransactionPerDate(string agency, string @event, DateTime dateEvent, string tid);
        void DistributionRevenueAgencies(string agency, DateTime distributedDate, int num, recTransaction[] transactions, decimal totalSum, recDistributedAmount[] distributedAmount);
    }
}
