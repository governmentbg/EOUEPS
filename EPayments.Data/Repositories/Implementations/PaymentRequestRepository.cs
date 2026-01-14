using EPayments.Common.Data;
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
using EPayments.Common.Helpers;

namespace EPayments.Data.Repositories.Implementations
{
    public class PaymentRequestRepository : IPaymentRequestRepository
    {
        private readonly IUnitOfWork UnitOfWork;
        private ISystemRepository systemRepository;

        public PaymentRequestRepository(IUnitOfWork unitOfWork, ISystemRepository systemRepository)
        {
            this.UnitOfWork = unitOfWork ?? throw new ArgumentNullException("unitOfWork is null");
            this.systemRepository = systemRepository;
        }

        public async Task ChangePaymentRequestsStatus(string filterPaymentIdentifier,
            string filterPaymentReferenceNumber,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            PaymentRequestStatus? filterRequestStatus,
            ObligationStatusEnum? obligationStatus,
            string applicantName,
            string applicantUin,
            PaymentRequestStatus newRequestStatus)
        {
            var predicate = CreatePaymentRequestPredicate(filterPaymentIdentifier,
                filterPaymentReferenceNumber,
                filterDateFrom,
                filterDateTo,
                filterAmountFrom,
                filterAmountTo,
                filterServiceProvider,
                filterPaymentReason,
                filterRequestStatus,
                obligationStatus,
                applicantName,
                applicantUin);

            var paymentRequestsList = await this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .Include(p => p.ObligationType)
                .Include(p => p.BoricaTransactions)
                .Where(predicate)
                .ToListAsync();

            foreach (var paymentRequest in paymentRequestsList)
            {
                paymentRequest.PaymentRequestStatusId = newRequestStatus;
                paymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;

                if (!String.IsNullOrWhiteSpace(paymentRequest.AdministrativeServiceNotificationURL))
                {
                    var shouldSendNotification = true;

                    if ((paymentRequest.PayOrder.HasValue || (paymentRequest.ObligationType != null && paymentRequest.ObligationType.AlgorithmId == 2))
                        && paymentRequest.PaymentRequestStatusId != PaymentRequestStatus.Paid)
                    {
                        shouldSendNotification = false;
                    }

                    if (shouldSendNotification)
                    {
                        EserviceNotification statusNotification = new EserviceNotification(paymentRequest, (int?)null);

                        this.UnitOfWork.DbContext.Set<EserviceNotification>().Add(statusNotification);
                    }
                }

                if (!paymentRequest.BoricaTransactions.Any() && newRequestStatus == PaymentRequestStatus.Paid)
                {
                    int counter;
                    using (var transaction = this.UnitOfWork.BeginTransaction())
                    {
                        counter = this.systemRepository.GetPaymentRequestTotal();
                        transaction.Commit();
                    }

                    if (paymentRequest.ObligationStatusId != ObligationStatusEnum.ForDistribution
                                                || paymentRequest.ObligationStatusId != ObligationStatusEnum.CheckedAccount)
                    {
                        paymentRequest.ObligationStatusId = ObligationStatusEnum.Paid;
                    }

                    BoricaTransaction boricaTransaction = new BoricaTransaction()
                    {
                        Amount = paymentRequest.PaymentAmount,
                        Order = String.Format("{0}{1}", Formatter.DateToShortFormat(DateTime.Now), counter),
                        PaymentRequests = new List<PaymentRequest>() { paymentRequest },
                        TransactionDate = DateTime.Now.ToUniversalTime(),
                        Gid = Guid.NewGuid(),
                        TransactionStatusId = (int)BoricaTransactionStatusEnum.TaxReceived,
                        Rc = "00",
                        StatusMessage = "Платено през админ"
                    };
                    this.UnitOfWork.DbContext.Set<BoricaTransaction>().Add(boricaTransaction);
                }
            }
            this.UnitOfWork.Save();
        }

        public async Task<int> CountPaymentRequests(string filterPaymentIdentifier,
            string filterPaymentReferenceNumber,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            PaymentRequestStatus? paymentRequestStatus,
            ObligationStatusEnum? obligationStatus,
            string applicantName,
            string applicantUin)
        {
            var predicate = CreatePaymentRequestPredicate(
                filterPaymentIdentifier,
                filterPaymentReferenceNumber,
                filterDateFrom,
                filterDateTo,
                filterAmountFrom,
                filterAmountTo,
                filterServiceProvider,
                filterPaymentReason,
                paymentRequestStatus,
                obligationStatus,
                applicantName,
                applicantUin);

            return await this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .AsNoTracking()
                .Where(predicate)
                .CountAsync();
        }

        public async Task<List<PaymentRequestVO>> GetPaymentRequests(string filterPaymentIdentifier,
            string filterPaymentReferenceNumber,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            PaymentRequestStatus? filterRequestStatus,
            ObligationStatusEnum? obligationStatus,
            string applicantName,
            string applicantUin,
            string sortBy,
            bool sortDescending,
            int page,
            int resultsPerPage)
        {
            var predicate = CreatePaymentRequestPredicate(filterPaymentIdentifier,
                filterPaymentReferenceNumber,
                filterDateFrom,
                filterDateTo,
                filterAmountFrom,
                filterAmountTo,
                filterServiceProvider,
                filterPaymentReason,
                filterRequestStatus,
                obligationStatus,
                applicantName,
                applicantUin);

            return await this.SortPaymentRequests(this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .AsNoTracking()
                .Where(predicate), sortBy, sortDescending)
                .Skip((page - 1) * resultsPerPage)
                .Take(resultsPerPage)
                .Select(PaymentRequestVO.Map)
                .ToListAsync();
        }

        public async Task<List<PaymentRequestVO>> GetAllRequests(string filterPaymentIdentifier, 
            string filterPaymentReferenceNumber, 
            DateTime? filterDateFrom, 
            DateTime? filterDateTo, 
            decimal? filterAmountFrom, 
            decimal? filterAmountTo, 
            string filterServiceProvider, 
            string filterPaymentReason, 
            PaymentRequestStatus? filterRequestStatus, 
            ObligationStatusEnum? obligationStatus, 
            string applicantName, 
            string applicantUin, 
            string sortBy, 
            bool sortDescending,
            int takeCount = int.MaxValue)
        {
            var predicate = CreatePaymentRequestPredicate(filterPaymentIdentifier,
                filterPaymentReferenceNumber,
                filterDateFrom,
                filterDateTo,
                filterAmountFrom,
                filterAmountTo,
                filterServiceProvider,
                filterPaymentReason,
                filterRequestStatus,
                obligationStatus,
                applicantName,
                applicantUin);

            return await this.SortPaymentRequests(this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .AsNoTracking()
                .Where(predicate), sortBy, sortDescending)
                .Select(PaymentRequestVO.Map)
                .Take(takeCount)
                .ToListAsync();
        }

        public List<string> GetAllPaymentReferenceNumbers()
        {
            return this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .AsNoTracking()
                .Select(x => x.PaymentReferenceNumber.Trim())
                .ToList();
        }

        public bool IsPaymentReferenceNumberExists(string referenceNumber)
        {
            referenceNumber = referenceNumber.Trim();

            return this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .AsNoTracking() // Avoids tracking to improve performance
                .Where(x => x.PaymentReferenceNumber == referenceNumber) // No Trim() on column for index usage
                .Select(x => x.PaymentReferenceNumber) // Optimized for minimal data retrieval
                .Any();
        }

        public List<string> GetPaymentReferenceNumbersForAlgorithm(int algorithmNumber)
        {
            return this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .AsNoTracking()
                .Where(x => x.ObligationType.AlgorithmId == algorithmNumber)
                .Select(x => x.PaymentReferenceNumber.Trim())
                .ToList();
        }

        public bool IsPaymentReferenceNumberExistsForAlgorithm(int algorithmNumber, string referenceNumber)
        {
            referenceNumber = referenceNumber.Trim();

            return this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .AsNoTracking() // Prevents EF from tracking entities, improving performance
                .Where(x => x.PaymentReferenceNumber == referenceNumber) // First, filter by reference number
                .Where(x => x.ObligationType.AlgorithmId == algorithmNumber) // Then, filter by algorithm ID
                .Select(x => x.PaymentReferenceNumber)
                .Any();
        }

        public List<PaymentRequest> GetPendingPaymentRequestByPaymentReferenceNumberAndAisPaymentId(string paymentReferenceNumber, string aisPaymentId)
        {
            List<PaymentRequest> result = null;

            if (aisPaymentId == null)
            {
                result = this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .AsNoTracking()
                .Where(x => x.PaymentReferenceNumber == paymentReferenceNumber).ToList();
                //if(result != null && result..AisPaymentId != null)
                //{
                //    result = null;
                //}
            }
            else
            {
                result = this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .Where(x => x.PaymentReferenceNumber == paymentReferenceNumber && x.AisPaymentId == aisPaymentId).ToList();
            }
            return result;
        }

        public bool IsAisPaymentIdAlreadyUsed(string aisPaymentId)
        {
            PaymentRequest payment = this.UnitOfWork.DbContext.Set<PaymentRequest>()
                .AsNoTracking()
                .Where(x => x.AisPaymentId == aisPaymentId)
                .FirstOrDefault();

            if (payment != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void DeletePaymentRequest(PaymentRequest paymentRequest)
        {
            this.UnitOfWork.DbContext.Set<PaymentRequest>().Remove(paymentRequest);
        }

        private Expression<Func<PaymentRequest, bool>> CreatePaymentRequestPredicate(string filterPaymentIdentifier,
            string filterPaymentReferenceNumber,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            PaymentRequestStatus? paymentRequestStatus,
            ObligationStatusEnum? obligationStatus,
            string applicantName,
            string applicantUin)
        {
            var predicate = PredicateBuilder.True<PaymentRequest>();

            if (!String.IsNullOrWhiteSpace(filterPaymentIdentifier))
            {
                predicate = predicate.And(e => e.PaymentRequestIdentifier == filterPaymentIdentifier);
            }

            if (!String.IsNullOrWhiteSpace(filterPaymentReferenceNumber))
            {
                predicate = predicate.And(e => e.PaymentReferenceNumber == filterPaymentReferenceNumber);
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

            if (paymentRequestStatus.HasValue)
            {
                predicate = predicate.And(e => e.PaymentRequestStatusId == paymentRequestStatus.Value);
            }

            if (obligationStatus.HasValue)
            {
                predicate = predicate.And(e => e.ObligationStatusId == obligationStatus.Value);
            }

            if (!String.IsNullOrWhiteSpace(applicantUin))
            {
                predicate = predicate.And(e => e.ApplicantUin.Contains(applicantUin));
            }

            if (!String.IsNullOrWhiteSpace(applicantName))
            {
                predicate = predicate.AndStringContains(e => e.ApplicantName, applicantName);
            }

            return predicate;
        }

        private IQueryable<PaymentRequest> SortPaymentRequests(
            IQueryable<PaymentRequest> query,
            string sortBy,
            bool sortDescending)
        {
            switch (sortBy)
            {
                case nameof(PaymentRequest.CreateDate):
                    query = sortDescending == false ?
                        query.OrderBy(pr => pr.CreateDate) :
                        query.OrderByDescending(pr => pr.CreateDate);
                    break;
                case nameof(PaymentRequest.ApplicantName):
                    query = sortDescending == false ?
                        query.OrderBy(pr => pr.ApplicantName) :
                        query.OrderByDescending(pr => pr.ApplicantName);
                    break;
                case nameof(PaymentRequest.PaymentRequestIdentifier):
                    query = sortDescending == false ?
                        query.OrderBy(pr => pr.PaymentRequestIdentifier) :
                        query.OrderByDescending(pr => pr.PaymentRequestIdentifier);
                    break;
                case nameof(PaymentRequest.ExpirationDate):
                    query = sortDescending == false ?
                        query.OrderBy(pr => pr.ExpirationDate) :
                        query.OrderByDescending(pr => pr.ExpirationDate);
                    break;
                case nameof(PaymentRequest.ServiceProviderName):
                    query = sortDescending == false ?
                        query.OrderBy(pr => pr.ServiceProviderName) :
                        query.OrderByDescending(pr => pr.ServiceProviderName);
                    break;
                case nameof(PaymentRequest.PaymentReason):
                    query = sortDescending == false ?
                        query.OrderBy(pr => pr.PaymentReason) :
                        query.OrderByDescending(pr => pr.PaymentReason);
                    break;
                case nameof(PaymentRequest.PaymentAmount):
                    query = sortDescending == false ?
                        query.OrderBy(pr => pr.CreateDate < CurrencyHelper.EuroAcceptanceDate ? pr.PaymentAmount / CurrencyHelper.BgnToEuroRate : pr.PaymentAmount) :
                        query.OrderByDescending(pr => pr.CreateDate < CurrencyHelper.EuroAcceptanceDate ? pr.PaymentAmount / CurrencyHelper.BgnToEuroRate : pr.PaymentAmount);
                    break;
                case nameof(PaymentRequest.PaymentRequestStatusId):
                    query = sortDescending == false ?
                        query.OrderBy(pr => pr.PaymentRequestStatusId) :
                        query.OrderByDescending(pr => pr.PaymentRequestStatusId);
                    break;
                case nameof(PaymentRequest.PaymentReferenceNumber):
                    query = sortDescending == false ?
                        query.OrderBy(pr => pr.PaymentReferenceNumber) :
                        query.OrderByDescending(pr => pr.PaymentReferenceNumber);
                    break;
                case nameof(PaymentRequest.ObligationStatusId):
                    query = sortDescending == false ?
                        query.OrderBy(pr => pr.ObligationStatusId) :
                        query.OrderByDescending(pr => pr.ObligationStatusId);
                    break;
            }

            return query;
        }
    }
}
