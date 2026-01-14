using EPayments.Common.Data;
using EPayments.Common.Helpers;
using EPayments.Common.Linq;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Data.ViewObjects.Admin;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;

namespace EPayments.Data.Repositories.Implementations
{
    public class EquationControlsRepository : IEquationControlsRepository
    {
        private readonly IUnitOfWork UnitOfWork;

        public EquationControlsRepository(IUnitOfWork unitOfWork)
        {
            this.UnitOfWork = unitOfWork ?? throw new ArgumentNullException("unitOfWork is null");
        }

        public async Task<int> CountUndistributetPayments(string filterPaymentIdentifier,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            ObligationStatusEnum? obligationStatus)
        {
            var predicate = CreatePaymentRequestPredicate(
                filterPaymentIdentifier,
                filterDateFrom,
                filterDateTo,
                filterAmountFrom,
                filterAmountTo,
                filterServiceProvider,
                filterPaymentReason,
                obligationStatus);

            return await this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .Where(predicate)
                .CountAsync();
        }

        public async Task<List<UndistributedPaymentRequestVO>> GetUndistributetPayments(string filterPaymentIdentifier,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            ObligationStatusEnum? obligationStatus,
            string sortBy,
            bool sortDescending,
            int page,
            int resultsPerPage)
        {
            var predicate = CreatePaymentRequestPredicate(filterPaymentIdentifier,
                filterDateFrom,
                filterDateTo,
                filterAmountFrom,
                filterAmountTo,
                filterServiceProvider,
                filterPaymentReason,
                obligationStatus);

            return await this.SortPaymentRequests(this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .Where(predicate), sortBy, sortDescending)
                .Skip((page - 1) * resultsPerPage)
                .Take(resultsPerPage)
                .Select(UndistributedPaymentRequestVO.Map)
                .ToListAsync();
        }

        public async Task<List<PaymentRequestVO>> GetAllUndistributetPayments(string filterPaymentIdentifier,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            ObligationStatusEnum? obligationStatus,
            string sortBy,
            bool sortDescending,
            int takeCount = int.MaxValue)
        {
            var predicate = CreatePaymentRequestPredicate(filterPaymentIdentifier,
                filterDateFrom,
                filterDateTo,
                filterAmountFrom,
                filterAmountTo,
                filterServiceProvider,
                filterPaymentReason,
                obligationStatus);

            return await this.SortPaymentRequests(this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .Where(predicate), sortBy, sortDescending)
                .Select(PaymentRequestVO.Map)
                .Take(takeCount)
                .ToListAsync();
        }

        private Expression<Func<PaymentRequest, bool>> CreatePaymentRequestPredicate(string filterPaymentIdentifier,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            ObligationStatusEnum? obligationStatus)
        {
            var predicate = PredicateBuilder.True<PaymentRequest>();

            if (!String.IsNullOrWhiteSpace(filterPaymentIdentifier))
            {
                predicate = predicate.And(e => e.PaymentRequestIdentifier == filterPaymentIdentifier);
            }

            if (filterDateFrom.HasValue)
            {
                predicate = predicate.And(e => e.PaymentRequestStatusChangeTime >= filterDateFrom.Value);
            }

            if (filterDateTo.HasValue)
            {
                predicate = predicate.And(e => e.PaymentRequestStatusChangeTime <= filterDateTo.Value);
            }

            if (filterAmountFrom.HasValue)
            {
                predicate = predicate.And(e =>
                      (CurrencyHelper.IsEuroTimePeriod ?
                        e.CreateDate < CurrencyHelper.EuroAcceptanceDate ? e.PaymentAmount / CurrencyHelper.BgnToEuroRate : e.PaymentAmount :
                        e.CreateDate > CurrencyHelper.EuroAcceptanceDate ? e.PaymentAmount * CurrencyHelper.BgnToEuroRate : e.PaymentAmount) >= filterAmountFrom.Value);
            }

            if (filterAmountTo.HasValue)
            {
                predicate = predicate.And(e =>
                    (CurrencyHelper.IsEuroTimePeriod ?
                        e.CreateDate < CurrencyHelper.EuroAcceptanceDate ? e.PaymentAmount / CurrencyHelper.BgnToEuroRate : e.PaymentAmount :
                        e.CreateDate > CurrencyHelper.EuroAcceptanceDate ? e.PaymentAmount * CurrencyHelper.BgnToEuroRate : e.PaymentAmount) <= filterAmountTo.Value);
            }

            if (!String.IsNullOrWhiteSpace(filterServiceProvider))
            {
                predicate = predicate.AndStringContains(e => e.ServiceProviderName, filterServiceProvider);
            }

            if (!String.IsNullOrWhiteSpace(filterPaymentReason))
            {
                predicate = predicate.AndStringContains(e => e.PaymentReason, filterPaymentReason);
            }

            if (obligationStatus.HasValue)
            {
                predicate = predicate.And(e => e.ObligationStatusId == obligationStatus.Value);
            }

            return predicate;
        }

        private IQueryable<PaymentRequest> SortPaymentRequests(
            IQueryable<PaymentRequest> query,
            string sortBy,
            bool sortAscending)
        {
            switch (sortBy)
            {
                case nameof(PaymentRequest.CreateDate):
                    query = sortAscending == true ?
                        query.OrderBy(pr => pr.CreateDate) :
                        query.OrderByDescending(pr => pr.CreateDate);
                    break;
                case nameof(PaymentRequest.PaymentRequestIdentifier):
                    query = sortAscending == true ?
                        query.OrderBy(pr => pr.PaymentRequestIdentifier) :
                        query.OrderByDescending(pr => pr.PaymentRequestIdentifier);
                    break;
                case nameof(PaymentRequest.ServiceProviderName):
                    query = sortAscending == true ?
                        query.OrderBy(pr => pr.ServiceProviderName) :
                        query.OrderByDescending(pr => pr.ServiceProviderName);
                    break;
                case nameof(PaymentRequest.PaymentReason):
                    query = sortAscending == true ?
                        query.OrderBy(pr => pr.PaymentReason) :
                        query.OrderByDescending(pr => pr.PaymentReason);
                    break;
                case nameof(PaymentRequest.PaymentAmount):
                    query = sortAscending == true ?
                        query.OrderBy(pr => pr.CreateDate < CurrencyHelper.EuroAcceptanceDate ? pr.PaymentAmount / CurrencyHelper.BgnToEuroRate : pr.PaymentAmount) :
                        query.OrderByDescending(pr => pr.CreateDate < CurrencyHelper.EuroAcceptanceDate ? pr.PaymentAmount / CurrencyHelper.BgnToEuroRate : pr.PaymentAmount);
                    break;
                case nameof(PaymentRequest.ObligationStatusId):
                    query = sortAscending == true ?
                        query.OrderBy(pr => pr.ObligationStatusId) :
                        query.OrderByDescending(pr => pr.ObligationStatusId);
                    break;
            }

            return query;        }



        public async Task<int> CountOldPayments(DateTime? filterDateFrom,
            DateTime? filterDateTo,
            ObligationStatusEnum? obligationStatus)
        {
            var predicate = CreateOldPaymentRequestPredicate(filterDateFrom,
                filterDateTo,
                obligationStatus);

            return await this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .Where(predicate)
                .CountAsync();
        }

        public async Task<List<PaymentRequestVO>> GetOldPayments(DateTime? filterDateFrom,
            DateTime? filterDateTo,
            ObligationStatusEnum? obligationStatus,
            string sortBy,
            bool sortDescending,
            int page,
            int resultsPerPage,
            int takeCount = int.MaxValue)
        {
            var baseQuery = this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .AsNoTracking()
                .Where(CreateOldPaymentRequestPredicate(filterDateFrom, filterDateTo, obligationStatus));

            baseQuery = SortPaymentRequests(baseQuery, sortBy, sortDescending);

            var paymentRequestIds = await baseQuery
                .AsNoTracking()
                .Skip((page - 1) * resultsPerPage)
                .Take(resultsPerPage)
                .Select(pr => pr.PaymentRequestId)
                .ToListAsync();

            if (!paymentRequestIds.Any())
                return new List<PaymentRequestVO>();

            var paymentRequests = await this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .AsNoTracking()
                .Include(pr => pr.ObligationType)
                .Where(pr => paymentRequestIds.Contains(pr.PaymentRequestId))
                .ToListAsync();

            var latestTransactions = await this.UnitOfWork.DbContext.Set<BoricaTransaction>()
                .AsNoTracking()
                .Where(bt => bt.PaymentRequests.Any(btpr =>
                    paymentRequestIds.Contains(btpr.PaymentRequestId)))
                .Select(bt => new
                {
                    bt.BoricaTransactionId,
                    bt.TransactionDate,
                    bt.Rc,
                    bt.StatusMessage,
                    PaymentRequestId = bt.PaymentRequests
                        .Select(btpr => btpr.PaymentRequestId)
                        .FirstOrDefault()
                })
                .ToListAsync();

            var latestTransactionsByRequestId = latestTransactions
                .GroupBy(t => t.PaymentRequestId)
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderByDescending(t => t.TransactionDate).FirstOrDefault()
                );

            return paymentRequests
                .OrderBy(vo => paymentRequestIds.IndexOf(vo.PaymentRequestId))
                .Select(pr => new PaymentRequestVO
                {
                    Gid = pr.Gid,
                    CreateDate = pr.CreateDate,
                    ApplicantName = pr.ApplicantName,
                    PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                    PaymentReferenceNumber = pr.PaymentReferenceNumber,
                    ExpirationDate = pr.ExpirationDate,
                    ServiceProviderName = pr.ServiceProviderName,
                    PaymentReason = pr.PaymentReason,
                    PaymentAmount = pr.PaymentAmount,
                    PaymentRequestStatusId = pr.PaymentRequestStatusId,
                    ObligationStatusId = pr.ObligationStatusId,
                })
                .Take(takeCount)
                .ToList();
        }

        public async Task<List<PaymentRequestVO>> GetAllOldPayments(DateTime? filterDateFrom,
            DateTime? filterDateTo,
            ObligationStatusEnum? obligationStatus,
            string sortBy,
            bool sortDescending,
            int takeCount = int.MaxValue)
        {
            var predicate = CreateOldPaymentRequestPredicate(filterDateFrom,
                filterDateTo,
                obligationStatus);

            return await this.SortPaymentRequests(this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .Where(predicate), sortBy, sortDescending)
                .AsNoTracking()
                .Select(PaymentRequestVO.Map)
                .Take(takeCount)
                .ToListAsync();
        }

        public async Task<List<BoricaTransactionVO>> GetBoricaTransactions(DateTime? dateFrom, 
            DateTime? dateTo,  
            int? transactionStatus, 
            string sortBy, 
            bool sortDescending, 
            int page, 
            int resultsPerPage)
        {
            var predicate = this.CreateBoricaTransactionPredicate(transactionStatus, dateFrom, dateTo);

            var transactions = await this.SortBoricaTransactionRequests(
                this.UnitOfWork.DbContext.Set<BoricaTransaction>().Where(predicate), sortBy, sortDescending)
                .AsNoTracking()
                .Select(BoricaTransactionVO.MapFrom)
                .Skip((page - 1) * resultsPerPage)
                .Take(resultsPerPage)
                .ToListAsync();

            return transactions;
        }

        public async Task<TotalsBoricaTransactionVO> CountBoricaTransactions(DateTime? dateFrom,
                                                        DateTime? dateTo,
                                                        int? transactionStatus)
        {
            var predicate = this.CreateBoricaTransactionPredicate(transactionStatus, dateFrom, dateTo);

            var totals = await this.UnitOfWork.DbContext.Set<BoricaTransaction>()
                .Where(predicate)
                .GroupBy(t => 1)
                .AsNoTracking()
                .Select(g => new TotalsBoricaTransactionVO
                {
                    TotalPages = g.Count(),
                    CommissionInEuro = g.Sum(t => t.Commission ?? 0),
                    TotalFeeInEuro = g.Sum(t => t.Fee ?? 0),
                    TotalAmountInEuro = g.Sum(t => t.Amount)
                })
                .FirstOrDefaultAsync() ?? new TotalsBoricaTransactionVO();

            return totals;
        }

        public async Task<List<BoricaTransactionVO>> GetBoricaTransactions(DateTime? dateFrom, 
            DateTime? dateTo, 
            int? transactionStatus, 
            string sortBy, 
            bool sortDescending,
            int takeCount = int.MaxValue)
        {
            var predicate = this.CreateBoricaTransactionPredicate(transactionStatus, dateFrom, dateTo);

            var transactions = await this.SortBoricaTransactionRequests(this.UnitOfWork.DbContext.Set<BoricaTransaction>()
                .Where(predicate), sortBy, sortDescending)
                .AsNoTracking()
                .Select(BoricaTransactionVO.MapFrom)
                .Take(takeCount)
                .ToListAsync();

            return transactions;
        }

        public async Task<BoricaTransactionVO> GetBoricaTransaction(int transactionId)
        {
            BoricaTransaction boricaTransaction = await this.UnitOfWork.DbContext.Set<BoricaTransaction>()
                .AsNoTracking()
                .FirstOrDefaultAsync(bt => bt.BoricaTransactionId == transactionId);

            if (boricaTransaction != null)
            {
                return BoricaTransactionVO.MapFrom.Compile()(boricaTransaction);
            }

            return null;
        }

        public async Task<List<PaymentRequestVO>> GetTransactionPaymentRequests(int transactionId)
        {
            return await this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .Where(pr => pr.BoricaTransactions.Any(bt => bt.BoricaTransactionId == transactionId))
                .AsNoTracking()
                .Select(PaymentRequestVO.Map)
                .ToListAsync();
        }

        private Expression<Func<PaymentRequest, bool>> CreateOldPaymentRequestPredicate(DateTime? filterDateFrom,
            DateTime? filterDateTo,
            ObligationStatusEnum? obligationStatus)
        {
            var predicate = PredicateBuilder.True<PaymentRequest>();

            if (filterDateFrom.HasValue)
            {
                predicate = predicate.And(e => e.PaymentRequestStatusChangeTime >= filterDateFrom.Value);
            }

            if (filterDateTo.HasValue)
            {
                predicate = predicate.And(e => e.PaymentRequestStatusChangeTime <= filterDateTo.Value);
            }           

            if (obligationStatus.HasValue)
            {
                predicate = predicate.And(e => e.ObligationStatusId == obligationStatus.Value);
            }

            return predicate;
        }

        private Expression<Func<BoricaTransaction, bool>> CreateBoricaTransactionPredicate(int? transactionStatus, DateTime? dateFrom, DateTime? dateTo)
        {
            var predicate = PredicateBuilder.True<BoricaTransaction>();

            if (transactionStatus.HasValue)
            {
                predicate = predicate.And(e => e.TransactionStatusId == transactionStatus.Value);
            }
            if (dateFrom.HasValue)
            {
                predicate = predicate.And(t => t.TransactionDate >= dateFrom);
            }
            if (dateTo.HasValue)
            {
                predicate = predicate.And(t => t.TransactionDate <= dateTo);
            }

            if (dateFrom.HasValue)
            {
                predicate = predicate.And(e => e.SettlementDate == null ? e.TransactionDate >= dateFrom : e.SettlementDate >= dateFrom);
            }

            if (dateTo.HasValue)
            {
                predicate = predicate.And(e => e.SettlementDate == null ? e.TransactionDate <= dateTo : e.SettlementDate <= dateTo);
            }

            return predicate;
        }

        private IQueryable<BoricaTransaction> SortBoricaTransactionRequests(
            IQueryable<BoricaTransaction> query,
            string sortBy,
            bool sortDescending)
        {
            switch (sortBy)
            {
                case nameof(BoricaTransaction.Order):
                    query = sortDescending == true ? query.OrderByDescending(q => q.Order) : query.OrderBy(q => q.Order);
                    break;
                case nameof(BoricaTransaction.Amount):
                    query = sortDescending == true ? query.OrderByDescending(q => q.TransactionDate < CurrencyHelper.EuroAcceptanceDate ? q.Amount / CurrencyHelper.BgnToEuroRate : q.Amount) :
                        query.OrderBy(q => q.TransactionDate < CurrencyHelper.EuroAcceptanceDate ? q.Amount / CurrencyHelper.BgnToEuroRate : q.Amount);
                    break;
                case nameof(BoricaTransaction.Fee):
                    query = sortDescending == true ? query.OrderByDescending(q => q.Fee) : query.OrderBy(q => q.Fee);
                    break;
                case nameof(BoricaTransaction.Commission):
                    query = sortDescending == true ? query.OrderByDescending(q => q.Commission) : query.OrderBy(q => q.Commission);
                    break;
                case nameof(BoricaTransaction.TransactionDate):
                    query = sortDescending == true ? query.OrderByDescending(q => q.TransactionDate) : query.OrderBy(q => q.TransactionDate);
                    break;
                case nameof(BoricaTransaction.Card):
                    query = sortDescending == true ? query.OrderByDescending(q => q.Card) : query.OrderBy(q => q.Card);
                    break;
                case nameof(BoricaTransaction.SettlementDate):
                    query = sortDescending == true ? query.OrderByDescending(q => q.SettlementDate) : query.OrderBy(q => q.SettlementDate);
                    break;
                case nameof(BoricaTransaction.StatusMessage):
                    query = sortDescending == true ? query.OrderByDescending(q => q.StatusMessage) : query.OrderBy(q => q.StatusMessage);
                    break;
            }

            return query;
        }
    }
}
