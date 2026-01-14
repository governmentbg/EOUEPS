using EPayments.Common.Data;
using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Data.ViewObjects.Admin;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

namespace EPayments.Data.Repositories.Implementations
{
    public class DistributionRepository : IDistributionRepository
    {
        private readonly IUnitOfWork UnitOfWork;

        public DistributionRepository(IUnitOfWork unitOfWork)
        {
            UnitOfWork = unitOfWork ?? throw new ArgumentNullException("unitOfWork is null");
        }

        public async Task<int> CountDistributionRevenues(DateTime? startDate,
            DateTime? endDate,
            int? distributionTypeId)
        {
            return await this.GetWhereClause(this.UnitOfWork.DbContext.Set<DistributionRevenue>(), startDate, endDate, distributionTypeId, null)
                .CountAsync();
        }

        public async Task<DistributionRevenueVO> GetDistributionById(int id)
        {
            DistributionRevenue distributionRevenue = await this.UnitOfWork.DbContext.Set<DistributionRevenue>()
                .Include(dr => dr.DistributionErrors)
                .SingleOrDefaultAsync(dr => dr.DistributionRevenueId == id);

            if (distributionRevenue != null)
            {
                return DistributionRevenueVO.Map.Compile()(distributionRevenue);
            }

            return null;
        }

        public async Task<List<DistributedPaymentRequestVO>> GetDistributionPaymentRequests(int id,
            string sortByPropertyName,
            bool isDescending)
        {
            return await this.SortPaymentsClause(this.UnitOfWork.DbContext.Set<DistributionRevenuePayment>()
                .Include(drp => drp.PaymentRequest.EserviceClient)
                .Include(drp => drp.EserviceClient)
                .Where(drp => drp.DistributionRevenueId == id), sortByPropertyName, isDescending)
                .Select(DistributedPaymentRequestVO.Map)
                .ToListAsync();
        }

        public async Task<List<DistributedPaymentRequestVO>> GetAllDistributionPaymentRequests(int id,
            string sortByPropertyName,
            bool isDescending,
            int takeCount = int.MaxValue)
        {
            return await this.SortPaymentsClause(this.UnitOfWork.DbContext.Set<DistributionRevenuePayment>()
                .Include(drp => drp.PaymentRequest.EserviceClient)
                .Include(drp => drp.EserviceClient)
                .Where(drp => drp.DistributionRevenueId == id), sortByPropertyName, isDescending)
                .Select(DistributedPaymentRequestVO.Map)
                .Take(takeCount)
                .ToListAsync();
        }

        public async Task<int> CountDistributionPaymentRequests(int id)
        {
            return await this.UnitOfWork.DbContext.Set<DistributionRevenuePayment>()
                .Include(drp => drp.PaymentRequest.EserviceClient)
                .Include(drp => drp.EserviceClient)
                .Where(drp => drp.DistributionRevenueId == id)
                .CountAsync();
        }

        public async Task<DistributionRevenue> GetDistribution(int id)
        {
            return await this.UnitOfWork.DbContext.Set<DistributionRevenue>()
                .Include(dr => dr.DistributionErrors)
                .SingleOrDefaultAsync(dr => dr.DistributionRevenueId == id);
        }

        public async Task<List<ObligationType>> GetAllObligationTypes()
        {
            return await this.UnitOfWork.DbContext.Set<ObligationType>().ToListAsync();
        }

        public async Task<List<DistributionRevenueVO>> GetDistributionRevenues(int currentPage,
            int pageLength,
            DateTime? startDate,
            DateTime? endDate,
            int? distributionTypeId,
            int? distributionRevenueId,
            string sortByPropertyName,
            bool isDescending)
        {

            var baseQuery = this.GetWhereClause(this.UnitOfWork.DbContext.Set<DistributionRevenue>(),
                                      startDate, endDate, distributionTypeId, distributionRevenueId);

            baseQuery = this.SortClause(baseQuery, sortByPropertyName, isDescending);

            var pagedRevenueIds = await baseQuery
                .AsNoTracking()
                .Select(dr => dr.DistributionRevenueId)
                .Skip((currentPage - 1) * pageLength)
                .Take(pageLength)
                .ToListAsync();

            var revenues = await this.UnitOfWork.DbContext.Set<DistributionRevenue>()
                .AsNoTracking()
                .Where(dr => pagedRevenueIds.Contains(dr.DistributionRevenueId))
                .ToListAsync();

            var errors = await this.UnitOfWork.DbContext.Set<DistributionError>()
                .AsNoTracking()
                .Where(e => pagedRevenueIds.Contains(e.DistributionRevenueId))
                .ToListAsync();

            var errorsByRevenueId = errors
                .GroupBy(e => e.DistributionRevenueId)
                .ToDictionary(g => g.Key, g => g.Select(e => e.Error).ToList());

            var result = revenues.Select(dr => new DistributionRevenueVO
            {
                DistributionRevenueId = dr.DistributionRevenueId,
                CreatedAt = dr.CreatedAt,
                DistributedDate = dr.DistributedDate,
                IsDistributed = dr.IsDistributed,
                TotalSum = dr.TotalSum,
                DistributionType = dr.DistributionTypeId,
                IsFileGenerated = dr.IsFileGenerated,
                FileName = dr.FileName,
                Errors = errorsByRevenueId.ContainsKey(dr.DistributionRevenueId)
                    ? errorsByRevenueId[dr.DistributionRevenueId]
                    : new List<string>()
            })
            .OrderBy(dr => pagedRevenueIds.IndexOf(dr.DistributionRevenueId))
            .ToList();

            return result;
        }

        public async Task<List<PaymentRequest>> GetDistributionPaymentRequests(int id)
        {
            return await this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .Where(pr => pr.DistributionRevenuePayment.DistributionRevenueId == id)
                .ToListAsync();
        }

        public async Task<List<DistribtutionTypeVO>> GetDistributionTypes()
        {
            return await this.UnitOfWork.DbContext.Set<DistributionType>()
                .Select(DistribtutionTypeVO.Map)
                .ToListAsync();
        }

        public async Task<List<DistributionRevenuePayment>> GetDistributionRevenuePayments(int id)
        {
            return await UnitOfWork.DbContext.Set<DistributionRevenuePayment>()
                .Include(drp => drp.PaymentRequest)
                .Include(drp => drp.EserviceClient.Department)
                .Include(drp => drp.BoricaTransaction)
                .Where(drp => drp.DistributionRevenueId == id)
                .ToListAsync();
        }

        public void Save()
        {
            this.UnitOfWork.Save();
        }

        private IQueryable<DistributionRevenue> GetWhereClause(
            IQueryable<DistributionRevenue> query,
            DateTime? startDate,
            DateTime? endDate,
            int? distributionTypeId,
            int? distributionRevenueId)
        {
            if (startDate != null)
            {
                query = query.Where(dr => dr.CreatedAt >= startDate);
            }

            if (endDate != null)
            {
                query = query.Where(dr => dr.CreatedAt <= endDate);
            }

            if (distributionTypeId != null)
            {
                int distribtutionId = (int)distributionTypeId;

                query = query.Where(dr => dr.DistributionTypeId == distribtutionId);
            }

            if (distributionRevenueId != null)
            {
                int refid = (int)distributionRevenueId;

                query = query.Where(dr => dr.DistributionRevenueId == refid);
            }

            return query;
        }

        private IQueryable<DistributionRevenue> SortClause(IQueryable<DistributionRevenue> query,
            string sortByPropertyName,
            bool isDescending)
        {
            switch (sortByPropertyName)
            {
                case nameof(DistributionRevenue.CreatedAt):
                    query = isDescending == false ?
                        query.OrderBy(dr => dr.CreatedAt) :
                        query.OrderByDescending(dr => dr.CreatedAt);
                    break;
                case nameof(DistributionRevenue.DistributedDate):
                    query = isDescending == false ?
                        query.OrderBy(dr => dr.DistributedDate) :
                        query.OrderByDescending(dr => dr.DistributedDate);
                    break;
                case nameof(DistributionRevenue.IsDistributed):
                    query = isDescending == false ?
                        query.OrderBy(dr => dr.IsDistributed) :
                        query.OrderByDescending(dr => dr.IsDistributed);
                    break;
                case nameof(DistributionRevenue.IsFileGenerated):
                    query = isDescending == false ?
                        query.OrderBy(dr => dr.IsFileGenerated) :
                        query.OrderByDescending(dr => dr.IsFileGenerated);
                    break;
                case nameof(DistributionRevenue.FileName):
                    query = isDescending == false ?
                        query.OrderBy(dr => dr.FileName) :
                        query.OrderByDescending(dr => dr.FileName);
                    break;
                default:
                    query = isDescending == false ?
                        query.OrderBy(dr => dr.CreatedAt < CurrencyHelper.EuroAcceptanceDate ? dr.TotalSum / CurrencyHelper.BgnToEuroRate : dr.TotalSum) :
                        query.OrderByDescending(dr => dr.CreatedAt < CurrencyHelper.EuroAcceptanceDate ? dr.TotalSum / CurrencyHelper.BgnToEuroRate : dr.TotalSum);
                    break;
            }

            return query;
        }

        private IQueryable<DistributionRevenuePayment> SortPaymentsClause(
            IQueryable<DistributionRevenuePayment> query,
            string sortByPropertyName,
            bool isDescending)
        {
            switch (sortByPropertyName)
            {
                case "PaymentRequestIdentifier":
                    query = isDescending == false ?
                        query.OrderBy(drp => drp.PaymentRequest.PaymentRequestIdentifier) :
                        query.OrderByDescending(drp => drp.PaymentRequest.PaymentRequestIdentifier);
                    break;
                case "PaymentReason":
                    query = isDescending == false ?
                        query.OrderBy(drp => drp.PaymentRequest.PaymentReason) :
                        query.OrderByDescending(drp => drp.PaymentRequest.PaymentReason);
                    break;
                case "PaymentAmount":
                    query = isDescending == false ?
                        query.OrderBy(drp => drp.PaymentRequest.PaymentAmount) :
                        query.OrderByDescending(drp => drp.PaymentRequest.PaymentAmount);
                    break;
                case "EServiceClientName":
                    query = isDescending == false ?
                        query.OrderBy(drp => drp.PaymentRequest.EserviceClient.AisName) :
                        query.OrderByDescending(drp => drp.PaymentRequest.EserviceClient.AisName);
                    break;
                case "TargetEServiceClientName":
                    query = isDescending == false ?
                        query.OrderBy(drp => drp.EserviceClient.AisName) :
                        query.OrderByDescending(drp => drp.EserviceClient.AisName);
                    break;
                case "ApplicantName":
                    query = isDescending == false ?
                        query.OrderBy(drp => drp.PaymentRequest.ApplicantName) :
                        query.OrderByDescending(drp => drp.PaymentRequest.ApplicantName);
                    break;
                case "PaymentRequestStatus":
                    query = isDescending == false ?
                        query.OrderBy(drp => drp.PaymentRequest.PaymentRequestStatusId) :
                        query.OrderByDescending(drp => drp.PaymentRequest.PaymentRequestStatusId);
                    break;
                case "ObligationStatus":
                    query = isDescending == false ?
                        query.OrderBy(drp => drp.PaymentRequest.ObligationStatusId) :
                        query.OrderByDescending(drp => drp.PaymentRequest.ObligationStatusId);
                    break;
            }

            return query;
        }
    }
}
