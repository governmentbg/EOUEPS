using EPayments.Data.ViewObjects.Admin;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace EPayments.Data.Repositories.Interfaces
{
    public interface IDistributionRepository
    {
        Task<List<ObligationType>> GetAllObligationTypes();

        Task<List<DistributionRevenueVO>> GetDistributionRevenues(int currentPage,
            int pageLength,
            DateTime? startDate,
            DateTime? endDate,
            int? distributionTypeId,
            int? distributionRevenueId,
            string sortByPropertyName,
            bool isDescending);

        Task<int> CountDistributionRevenues(DateTime? startDate,
            DateTime? endDate,
            int? distributionTypeId);

        Task<DistributionRevenueVO> GetDistributionById(int id);

        Task<DistributionRevenue> GetDistribution(int id);

        Task<List<DistributedPaymentRequestVO>> GetDistributionPaymentRequests(int id,
            string sortByPropertyName,
            bool isDescending);

        Task<List<DistributedPaymentRequestVO>> GetAllDistributionPaymentRequests(int id,
            string sortByPropertyName,
            bool isDescending, 
            int takeCount = int.MaxValue);

        Task<int> CountDistributionPaymentRequests(int id);

        Task<List<DistribtutionTypeVO>> GetDistributionTypes();

        Task<List<PaymentRequest>> GetDistributionPaymentRequests(int id);

        Task<List<DistributionRevenuePayment>> GetDistributionRevenuePayments(int id);

        void Save();
    }
}
