using EPayments.Common.Data;
using EPayments.Data.Repositories.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using EPayments.Model.Models;
using EPayments.Model.Enums;
using EPayments.Data.ViewObjects.Web;
using System.Linq.Expressions;
using EPayments.Common.Linq;
using EPayments.Common.Helpers;
using System.Data.Entity;
using EPayments.Data.ViewObjects.Admin;

namespace EPayments.Data.Repositories.Implementations
{
    internal class WebRepository : BaseRepository, IWebRepository
    {
        public WebRepository(IUnitOfWork unitOfWork)
            : base(unitOfWork)
        {
        }

        public int CountPendingRequestsByUin(string uin)
        {
            return (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                    where pr.ApplicantUin == uin && pr.PaymentRequestStatusId == PaymentRequestStatus.Pending && !pr.IsTempRequest
                    select pr)
                    .Count();
        }

        public int CountProcessedRequestsByUin(
            string uin,
            string filterPaymentIdentifier,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            PaymentRequestStatus? filterRequestStatus)
        {
            var predicate = CreateSearchProcessedRequestPredicate(
                uin,
                filterPaymentIdentifier,
                filterDateFrom,
                filterDateTo,
                filterAmountFrom,
                filterAmountTo,
                filterServiceProvider,
                filterPaymentReason,
                filterRequestStatus);

            return
                this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(p => p.PaymentRequestStatusId != PaymentRequestStatus.Expired)
                .Where(predicate)
                .Count();
        }

        public int CountEserviceAdminRequests(
            int eserviceClientId,
            string filterPaymentIdentifier,
            string filterReferencefNumber,
            string filterObligationType,
            string filterRefid,
            string filterPaymentType,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            PaymentRequestStatus? filterRequestStatus,
            ObligationStatusEnum? filterObligationStatus,
            string applicantName,
            string applicantUin)
        {
            var predicate = CreateSearchEserviceAdminRequestPredicate(
                eserviceClientId,
                filterPaymentIdentifier,
                filterReferencefNumber,
                filterRefid,
                filterPaymentType,
                filterDateFrom,
                filterDateTo,
                filterAmountFrom,
                filterAmountTo,
                filterServiceProvider,
                filterPaymentReason,
                filterRequestStatus,
                filterObligationStatus,
                applicantName,
                applicantUin);

            var obligationTypes = this.unitOfWork.DbContext.Set<ObligationType>();

            if (!string.IsNullOrEmpty(filterObligationType))
            {
                obligationTypes.Where(o => o.PaymentTypeCode == filterObligationType);
            }

            return
                this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(predicate)
                .Where(pr => obligationTypes.Select(o => o.ObligationTypeId).Contains(pr.ObligationTypeId))
                .Count();
        }

        public Tuple<int, decimal> CountAndGetTotalAmountForTransactionRecords(
            int eserviceBankAccountId,
            DateTime? transactionAccountingDateFrom,
            DateTime? transactionAccountingDateTo,
            decimal? transactionAmountFrom,
            decimal? transactionAmountTo,
            string infoDocumentNumber,
            DateTime? infoDocumentDateFrom,
            DateTime? infoDocumentDateTo,
            string infoSenderIban,
            string infoSenderName,
            string infoDebtorName,
            string infoDebtorBulstatEgnLnch,
            string infoPaymentReason,
            string infoAC1AuthorizationCode,
            int? transactionRecordPaymentMethod,
            int? transactionRecordRefStatus)
        {
            var predicate = CreateSearchTransactionRecordsPredicate(
                eserviceBankAccountId,
                transactionAccountingDateFrom,
                transactionAccountingDateTo,
                transactionAmountFrom,
                transactionAmountTo,
                infoDocumentNumber,
                infoDocumentDateFrom,
                infoDocumentDateTo,
                infoSenderIban,
                infoSenderName,
                infoDebtorName,
                infoDebtorBulstatEgnLnch,
                infoPaymentReason,
                infoAC1AuthorizationCode,
                transactionRecordPaymentMethod,
                transactionRecordRefStatus);

            List<decimal?> transactionAmounts = this.unitOfWork.DbContext.Set<TransactionRecord>()
                .Where(predicate)
                .Select(e => e.TransactionAmount)
                .ToList();

            int count = transactionAmounts.Count;
            decimal totalAmount = transactionAmounts.Sum(e => e ?? 0);

            return new Tuple<int, decimal>(count, totalAmount);
        }

        public int CountEserviceRecords(
            string aisName,
            string departmentName,
            int? vposClientId,
            int? isActiveId)
        {
            var predicate = CreateEserviceRecordsPredicate(
                aisName,
                departmentName,
                vposClientId,
                isActiveId);

            return this.unitOfWork.DbContext.Set<EserviceClient>()
                .Where(predicate)
                .Count();
        }

        public int CountObligationTypeRecords(
        string oblName,
        int? isActiveId)
        {
            var predicate = CreateObligationTypeRecordsPredicate(
                oblName,
                isActiveId);

            return this.unitOfWork.DbContext.Set<ObligationType>()
                .Where(predicate)
                .Count();
        }

        public int CountDepartmentRecords(
            string departmentName,
            string departmentUniqueIdentificationNumber,
            string departmentUnifiedBudgetClassifier)
        {
            var predicate = CreateDepartmentRecordsPredicate(
                departmentName,
                departmentUniqueIdentificationNumber,
                departmentUnifiedBudgetClassifier);

            return this.unitOfWork.DbContext.Set<Department>()
                .Where(predicate)
                .Where(d => d.IsActive)
                .Count();
        }

        public EserviceClient GetEserviceClient(int eserviceClientId)
        {
            return this.unitOfWork.DbContext.Set<EserviceClient>()
                .Include(e => e.Department)
                .Where(e => e.EserviceClientId == eserviceClientId)
                .SingleOrDefault();
        }

        public Department GetDepartment(int departmentId)
        {
            return this.unitOfWork.DbContext.Set<Department>()
                .Where(d => d.DepartmentId == departmentId)
                .SingleOrDefault();
        }

        public ObligationType GetObligationType(int obligationTypeId)
        {
            return this.unitOfWork.DbContext.Set<ObligationType>()
                .Where(d => d.ObligationTypeId == obligationTypeId)
                .SingleOrDefault();
        }

        public EserviceAdminUser GetEserviceAdminUserByReferringEserviceClientId(int referringEserviceClientId)
        {
            return this.unitOfWork.DbContext.Set<EserviceAdminUser>()
                .Where(e => e.ReferringEserviceClientId == referringEserviceClientId)
                .FirstOrDefault();
        }

        public bool EserviceAdminUserExists(string username)
        {
            return this.unitOfWork.DbContext.Set<EserviceAdminUser>()
                .Where(e => e.Username == username)
                .Any();
        }

        public bool DepartmentHasReferencedClients(int departmentId, int excludeEserviceClientId)
        {
            return this.unitOfWork.DbContext.Set<EserviceClient>()
                .Where(e => e.EserviceClientId != excludeEserviceClientId && e.DepartmentId == departmentId)
                .Any();
        }

        public IList<EserviceRecordVO> GetEserviceRecords(
            string aisName,
            string departmentName,
            int? vposClientId,
            int? isActiveId,
            EserviceListColumn sortBy,
            bool sortDescending,
            int? page = null,
            int? resultsPerPage = null)
        {
            var predicate = CreateEserviceRecordsPredicate(
                aisName,
                departmentName,
                vposClientId,
                isActiveId);

            IQueryable<EserviceRecordVO> query =
                this.unitOfWork.DbContext.Set<EserviceClient>()
                .Where(predicate)
                .Select(e =>
                    new EserviceRecordVO()
                    {
                        EserviceClientId = e.EserviceClientId,
                        AisName = e.AisName,
                        DepartmentId = e.DepartmentId,
                        DepartmentName = e.Department.Name,
                        AccountBank = e.AccountBank,
                        VposClientId = e.VposClientId,
                        IsActive = e.IsActive
                    });

            query = GetSortableEserviceRecordsQuery(query, sortBy, sortDescending);

            if (page.HasValue)
            {
                return query
                    .Skip((page.Value - 1) * resultsPerPage.Value)
                    .Take(resultsPerPage.Value)
                    .ToList();
            }
            else
            {
                return query
                    .ToList();
            }
        }

        public IList<ObligationTypeVO> GetObligationTypeRecords(
            string oblName,
            int? isActiveId,
            ObligationTypeListColumn sortBy,
            bool sortDescending,
            int? page = null,
            int? resultsPerPage = null)

        {
            var predicate = CreateObligationTypeRecordsPredicate(
                oblName,
                isActiveId);

            IQueryable<ObligationTypeVO> query =
                this.unitOfWork.DbContext.Set<ObligationType>()
                .Where(predicate)
                .Select(e =>
                    new ObligationTypeVO()
                    {
                        ObligationTypeId = e.ObligationTypeId,
                        Name = e.Name,
                        IsActive = e.IsActive
                    });

            query = GetSortableObligationTypeRecordsQuery(query, sortBy, sortDescending);

            if (page.HasValue)
            {
                return query
                    .Skip((page.Value - 1) * resultsPerPage.Value)
                    .Take(resultsPerPage.Value)
                    .ToList();
            }
            else
            {
                return query
                    .ToList();
            }
        }

        public IList<DepartmentVO> GetDepartmentRecords(
            string departmentName,
            string departmentUniqueIdentificationNumber,
            string departmentUnifiedBudgetClassifier,
            DepartmentListColumn sortBy,
            bool sortDescending,
            int? page = null,
            int? resultsPerPage = null)
        {
            var predicate = CreateDepartmentRecordsPredicate(
                departmentName,
                departmentUniqueIdentificationNumber,
                departmentUnifiedBudgetClassifier);

            IQueryable<DepartmentVO> query =
                this.unitOfWork.DbContext.Set<Department>()
                .Where(predicate)
                .Where(d => d.IsActive)
                .Select(d =>
                    new DepartmentVO()
                    {
                        DepartmentId = d.DepartmentId,
                        Name = d.Name,
                        UniqueIdentificationNumber = d.UniqueIdentificationNumber,
                        UnifiedBudgetClassifier = d.UnifiedBudgetClassifier
                    });

            query = GetSortableDepartmentRecordsQuery(query, sortBy, sortDescending);

            if (page.HasValue)
            {
                return query
                    .Skip((page.Value - 1) * resultsPerPage.Value)
                    .Take(resultsPerPage.Value)
                    .ToList();
            }
            else
            {
                return query
                    .ToList();
            }
        }

        public List<EserviceBankAccount> GetEserviceBankAccountsByAdminUserId(int eserviceAdminUserId)
        {
            return
                (from ba in this.unitOfWork.DbContext.Set<EserviceBankAccount>()
                 join auba in this.unitOfWork.DbContext.Set<EserviceAdminUserBankAccount>() on ba.EserviceBankAccountId equals auba.EserviceBankAccountId
                 join au in this.unitOfWork.DbContext.Set<EserviceAdminUser>() on auba.EserviceAdminUserId equals au.EserviceAdminUserId
                 where ba.IsActive && au.EserviceAdminUserId == eserviceAdminUserId && au.IsActive && auba.IsActive
                 select ba)
                 .Include(e => e.Bank)
                 .ToList();
        }

        public int CountRequestAccessByUin(string uin)
        {
            IQueryable<RequestAccessVO> query = GetRequestAccessByUinQuery(uin);

            return query.Count();
        }

        public IList<PendingRequestVO> GetPendingRequestsByUin(
            string uin,
            PendingPaymentColumn sortBy,
            bool sortDescending)
        {
            var query = this.GetPendingRequestsByUinCustomer(uin, sortBy, sortDescending);
            return query.ToList();
        }


        public IList<PendingRequestVO> GetPendingRequestsByUin(
           string uin,
           PendingPaymentColumn sortBy,
           bool sortDescending,
           int page,
           int resultsPerPage)
        {
            var query = this.GetPendingRequestsByUinCustomer(uin, sortBy, sortDescending);
            return query
                    .Skip((page - 1) * resultsPerPage)
                    .Take(resultsPerPage)
                    .ToList();
        }

        public IList<ProcessedRequestVO> GetProcessedRequestsByUin(
            string uin,
            string filterPaymentIdentifier,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            PaymentRequestStatus? filterRequestStatus,
            ProcessedPaymentColumn sortBy,
            bool sortDescending,
            int page,
            int resultsPerPage)
        {
            var predicate = CreateSearchProcessedRequestPredicate(
                uin,
                filterPaymentIdentifier,
                filterDateFrom,
                filterDateTo,
                filterAmountFrom,
                filterAmountTo,
                filterServiceProvider,
                filterPaymentReason,
                filterRequestStatus);

            IQueryable<ProcessedRequestVO> query =
                this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Include(pr => pr.InitiatorEserviceClient)
                .Where(p => p.PaymentRequestStatusId != PaymentRequestStatus.Expired)
                .Where(predicate)
                .Select(pr =>
                    new ProcessedRequestVO()
                    {
                        Gid = pr.Gid,
                        ApplicantName = pr.ApplicantName,
                        ApplicantUin = pr.ApplicantUin,
                        PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                        TransactionDate = pr.PaymentRequestStatusChangeTime,
                        ServiceProviderName = pr.ServiceProviderName,
                        PaymentReason = pr.PaymentReason,
                        PaymentAmountRequest = pr.PaymentAmount,
                        PaymentRequestGid = pr.Gid,
                        PaymentRequestStatusId = pr.PaymentRequestStatusId,
                        ExpirationDate = pr.ExpirationDate,
                        CreateDate = pr.CreateDate,
                        PaymentReferenceType = pr.PaymentReferenceType,
                        ObligationStatusId = pr.ObligationStatusId,
                        ObligationType = pr.ObligationType.Name,
                        InitiatorName = pr.InitiatorEserviceClient.ServiceName,
                        AdditionalInformation = pr.AdditionalInformation
                    });

            query = GetSortableProcessedRequestQuery(query, sortBy, sortDescending);

            return query
                .Skip((page - 1) * resultsPerPage)
                .Take(resultsPerPage)
                .ToList();
        }

        public IList<ProcessedRequestVO> GetEserviceAdminRequests(
            int eserviceClientId,
            string filterPaymentIdentifier,
            string filterReferencefNumber,
            string filterObligationType,
            string filterRefid,
            string filterPeymentType,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            PaymentRequestStatus? filterRequestStatus,
            ObligationStatusEnum? filterObligationStatus,
            string applicantName,
            string applicantUin,
            EserviceAdminRequestColumn sortBy,
            bool sortDescending,
            int? page = null,
            int? resultsPerPage = null,
            int takeCount = int.MaxValue)
        {
            var predicate = CreateSearchEserviceAdminRequestPredicate(
                eserviceClientId,
                filterPaymentIdentifier,
                filterReferencefNumber,
                filterRefid,
                filterPeymentType,
                filterDateFrom,
                filterDateTo,
                filterAmountFrom,
                filterAmountTo,
                filterServiceProvider,
                filterPaymentReason,
                filterRequestStatus,
                filterObligationStatus,
                applicantName,
                applicantUin);

            var obligationTypes = this.unitOfWork.DbContext.Set<ObligationType>();

            if (!string.IsNullOrEmpty(filterObligationType))
            {
                obligationTypes.Where(o => o.PaymentTypeCode == filterObligationType);
            }

            IQueryable<ProcessedRequestVO> query =
                this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Include(pr => pr.PaymentRequestObligationLogs)
                .Where(pr => obligationTypes.Select(o => o.ObligationTypeId).Contains(pr.ObligationTypeId))
                .Where(predicate)
                .Select(pr =>
                    new ProcessedRequestVO()
                    {
                        Gid = pr.Gid,
                        ApplicantName = pr.ApplicantName,
                        ApplicantUin = pr.ApplicantUin,
                        PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                        TransactionDate = pr.PaymentRequestStatusChangeTime,
                        ServiceProviderName = pr.ServiceProviderName,
                        PaymentReason = pr.PaymentReason,
                        PaymentAmountRequest = pr.PaymentAmount,
                        CreateDate = pr.CreateDate,
                        PaymentRequestGid = pr.Gid,
                        PaymentRequestStatusId = pr.PaymentRequestStatusId,
                        ExpirationDate = pr.ExpirationDate,
                        ObligationStatusId = pr.ObligationStatusId,
                        PaymentReferenceNumber = pr.PaymentReferenceNumber,
                        AdditionalInformation = pr.AdditionalInformation,
                        ObligationType = pr.ObligationType != null ? pr.ObligationType.PaymentTypeCode : string.Empty,
                        PaymentReferenceType = pr.PaymentReferenceType,
                        Refid = pr.DistributionRevenuePayment.DistributionRevenueId.ToString(),
                        PaymentRequestObligationLogs = pr.PaymentRequestObligationLogs.Select(prol => new ViewObjects.Web.PaymentRequestObligationLog()
                        {
                            PaymentRequestId = prol.PaymentRequestId,
                            ObligationStatusId = prol.ObligationStatusId,
                            ChangeDate = prol.ChangeDate
                        })
                    });

            query = GetSortableEserviceAdminRequestQuery(query, sortBy, sortDescending);

            if (page.HasValue && resultsPerPage.HasValue)
            {
                return query
                   .Skip((page.Value - 1) * resultsPerPage.Value)
                   .Take(resultsPerPage.Value)
                   .ToList();
            }
            else
            {
                return query
                    .Take(takeCount)
                    .ToList();
            }
        }

        public IList<RequestAccessVO> GetRequestAccessByUin(
            string uin,
            RequestAccessListColumn sortBy,
            bool sortDescending,
            int page,
            int resultsPerPage)
        {
            IQueryable<RequestAccessVO> query = GetRequestAccessByUinQuery(uin);

            query = GetSortableRequestAccessByUinQuery(query, sortBy, sortDescending);

            return query
                .Skip((page - 1) * resultsPerPage)
                .Take(resultsPerPage)
                .ToList();
        }

        public IList<TransactionRecordVO> GetTransactionRecords(
            int eserviceBankAccountId,
            DateTime? transactionAccountingDateFrom,
            DateTime? transactionAccountingDateTo,
            decimal? transactionAmountFrom,
            decimal? transactionAmountTo,
            string infoDocumentNumber,
            DateTime? infoDocumentDateFrom,
            DateTime? infoDocumentDateTo,
            string infoSenderIban,
            string infoSenderName,
            string infoDebtorName,
            string infoDebtorBulstatEgnLnch,
            string infoPaymentReason,
            string infoAC1AuthorizationCode,
            int? transactionRecordPaymentMethod,
            int? transactionRecordRefStatus,
            TransactionListColumn sortBy,
            bool sortDescending,
            int? page = null,
            int? resultsPerPage = null)
        {
            var predicate = CreateSearchTransactionRecordsPredicate(
                eserviceBankAccountId,
                transactionAccountingDateFrom,
                transactionAccountingDateTo,
                transactionAmountFrom,
                transactionAmountTo,
                infoDocumentNumber,
                infoDocumentDateFrom,
                infoDocumentDateTo,
                infoSenderIban,
                infoSenderName,
                infoDebtorName,
                infoDebtorBulstatEgnLnch,
                infoPaymentReason,
                infoAC1AuthorizationCode,
                transactionRecordPaymentMethod,
                transactionRecordRefStatus);

            IQueryable<TransactionRecordVO> query =
                this.unitOfWork.DbContext.Set<TransactionRecord>()
                .Include(e => e.PaymentRequest)
                .Where(predicate)
                .Select(tr =>
                    new TransactionRecordVO()
                    {
                        TransactionRecordId = tr.TransactionRecordId,
                        TransactionAccountingDate = tr.TransactionAccountingDate,
                        TransactionAmount = tr.TransactionAmount,
                        InfoDocumentNumber = tr.InfoDocumentNumber,
                        InfoDocumentDate = tr.InfoDocumentDate,
                        InfoDocumentNumberDate = tr.InfoDocumentNumberDate,
                        InfoSenderIban = tr.InfoSenderIban,
                        InfoSenderName = tr.InfoSenderName,
                        InfoSenderIbanName = tr.InfoSenderIbanName,
                        InfoDebtorName = tr.InfoDebtorName,
                        InfoDebtorBulstatEgnLnch = tr.InfoDebtorBulstatEgnLnch,
                        InfoDebtorBulstatEgnLnchName = tr.InfoDebtorBulstatEgnLnchName,
                        InfoPaymentReason = tr.InfoPaymentReason,
                        TransactionRecordPaymentMethodId = tr.TransactionRecordPaymentMethodId,
                        TransactionDate = tr.TransactionDate,
                        TransactionRecordRefStatusId = tr.TransactionRecordRefStatusId,
                        PaymentRequestGid = tr.PaymentRequest.Gid
                    });

            query = GetSortableTransactionRecordsQuery(query, sortBy, sortDescending);

            if (page.HasValue)
            {
                return query
                    .Skip((page.Value - 1) * resultsPerPage.Value)
                    .Take(resultsPerPage.Value)
                    .ToList();
            }
            else
            {
                return query
                    .ToList();
            }
        }

        public int GetRequestAccessDetailsCount(int paymentRequestId, string uin)
        {
            return
                GetRequestAccessDetailsQuery(paymentRequestId, uin)
                .Count();
        }

        public List<RequestAccessDetailsVO> GetRequestAccessDetails(int paymentRequestId, string uin, int limit)
        {
            return
                GetRequestAccessDetailsQuery(paymentRequestId, uin)
                .OrderByDescending(e => e.AccessDate)
                .Take(limit)
                .ToList();
        }

        public IList<PendingRequestVO> GetPendingRequestsByAccessCode(string accessCode)
        {
            return
                (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>().Include(p => p.ObligationType)
                 join ac in this.unitOfWork.DbContext.Set<EserviceClient>() on pr.EserviceClientId equals ac.EserviceClientId
                 join vpc in this.unitOfWork.DbContext.Set<VposClient>() on ac.VposClientId equals vpc.VposClientId into g1
                 from vpc in g1.DefaultIfEmpty()
                 where pr.PaymentRequestAccessCode == accessCode && pr.PaymentRequestStatusId == PaymentRequestStatus.Pending && !pr.IsTempRequest
                 select new PendingRequestVO()
                 {
                     Gid = pr.Gid,
                     ApplicantName = pr.ApplicantName,
                     PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                     CreateDate = pr.CreateDate,
                     ExpirationDate = pr.ExpirationDate,
                     ServiceProviderName = pr.ServiceProviderName,
                     InitiatorName = pr.InitiatorEserviceClient.ServiceName,
                     PaymentReason = pr.PaymentReason,
                     PaymentAmount = pr.PaymentAmount,
                     IsEpayVposEnabled = ac.IsEpayVposEnabled,
                     ObligationType = pr.ObligationType.Name,
                     IsCvposEnabled = ac.IsBoricaVposEnabled,
                     Vpos = vpc != null && vpc.IsActive ? (Vpos)vpc.VposClientId : (Vpos?)null,
                     IsInProgress = pr.PaymentRequestStatusId == PaymentRequestStatus.InProcess,
                     PayOrder = pr.PayOrder,
                     EserviceClientId = pr.EserviceClientId,
                     AdditionalInfo = pr.AdditionalInformation,
                     AlgorithmId = pr.ObligationType.AlgorithmId,
                 })
                 .ToList();
        }

        public List<ProcessedRequestVO> GetProcessedRequestsByAccessCode(string accessCode)
        {
            return
                (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                 where pr.PaymentRequestAccessCode == accessCode &&
                       pr.PaymentRequestStatusId != PaymentRequestStatus.Pending && pr.PaymentRequestStatusId != PaymentRequestStatus.Expired &&
                       (!pr.IsTempRequest || pr.PaymentRequestStatusId == PaymentRequestStatus.Authorized || pr.PaymentRequestStatusId == PaymentRequestStatus.Paid)
                 select new ProcessedRequestVO()
                 {
                     Gid = pr.Gid,
                     ApplicantName = pr.ApplicantName,
                     ApplicantUin = pr.ApplicantUin,
                     PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                     TransactionDate = pr.PaymentRequestStatusChangeTime,
                     ServiceProviderName = pr.ServiceProviderName,
                     PaymentReason = pr.PaymentReason,
                     CreateDate = pr.CreateDate,
                     PaymentAmountRequest = pr.PaymentAmount,
                     PaymentRequestGid = pr.Gid,
                     PaymentRequestStatusId = pr.PaymentRequestStatusId,
                     ExpirationDate = pr.ExpirationDate,
                     ObligationType = pr.ObligationType.Name,
                     ObligationStatusId = pr.ObligationStatusId
                 })
                 .ToList();
        }

        public PaymentRequest GetPaymentRequestById(int paymentRequestId)
        {
            return this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(e => e.PaymentRequestId == paymentRequestId)
                .Include(p => p.ObligationType)
                .SingleOrDefault();
        }

        public PaymentRequest GetPaymentRequestByIdentifier(string paymentRequestIdentifier)
        {
            return this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(e => e.PaymentRequestIdentifier == paymentRequestIdentifier)
                .Include(p => p.ObligationType)
                .SingleOrDefault();
        }

        public List<PaymentRequest> GetPendngPaymentRequestByAisClient(int eServiceClientId)
        {
            return this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(e => e.EserviceClientId == eServiceClientId)
                .Where(e => e.PaymentRequestStatusId == PaymentRequestStatus.Pending)
                .Where(e => e.PayOrder.HasValue)
                .Where(e => !string.IsNullOrEmpty(e.AdditionalInformation))
                .ToList();
        }

        public List<PaymentRequest> GetPendngAndInProcessPaymentRequestByAisClient(int eServiceClientId, string applicantUin)
        {
            return this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(e => e.EserviceClientId == eServiceClientId)
                .Where(e => e.PaymentRequestStatusId == PaymentRequestStatus.Pending || e.PaymentRequestStatusId == PaymentRequestStatus.InProcess)
                .Where(e => e.PayOrder.HasValue)
                .Where(e => !string.IsNullOrEmpty(e.AdditionalInformation))
                .Where(e => e.ApplicantUin == applicantUin)
                .ToList();
        }

        public PaymentRequest GetPaymentRequestByGidAndUin(Guid gid, string uin)
        {
            return this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(e => e.Gid == gid && e.ApplicantUin == uin)
                .SingleOrDefault();
        }

        public PaymentOrderVO GetPaymentOrderByGid(Guid gid)
        {
            return GetPaymentOrderInternal(gid, null, null);
        }

        public PaymentOrderVO GetPaymentOrderByGidAndUin(Guid gid, string uin)
        {
            return GetPaymentOrderInternal(gid, uin, null);
        }

        public PaymentOrderVO GetPaymentOrderByAccessCode(string accessCode)
        {
            return GetPaymentOrderInternal(null, null, accessCode);
        }

        public TransactionRecord GetTransactionRecordByTransactionRecordId(int transactionRecordId)
        {
            return this.unitOfWork.DbContext.Set<TransactionRecord>()
                .Include(e => e.TransactionFile)
                .Where(e => e.TransactionRecordId == transactionRecordId)
                .SingleOrDefault();
        }

        public PaymentRequest GetPaymentRequestByAccessCode(string accessCode)
        {
            return this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(e => e.PaymentRequestAccessCode == accessCode)
                .SingleOrDefault();
        }

        public VposRequestDataVO GetVposRequestDataByGidAndUin(Guid gid, string uin)
        {
            return GetVposRequestDataInternal(gid, uin, null);
        }

        public VposRequestDataVO GetVposRequestDataByAccessCode(string accessCode)
        {
            return GetVposRequestDataInternal(null, null, accessCode);
        }

        public string GenerateVposBoricaRequestIdentifier()
        {
            CodeGenerator codeGenerator = new CodeGenerator();
            codeGenerator.Minimum = 15;
            codeGenerator.Maximum = 15;
            codeGenerator.ConsecutiveCharacters = true;
            codeGenerator.RepeatCharacters = true;

            while (true)
            {
                string code = codeGenerator.Generate();

                if (!this.unitOfWork.DbContext.Set<VposBoricaRequest>().Any(e => e.RequestIdentifier == code))
                {
                    return code;
                }
            }
        }

        public VposBoricaRequest GetVposBoricaRequestByRequestIdentifier(string requestIdentifier)
        {
            return this.unitOfWork.DbContext.Set<VposBoricaRequest>().SingleOrDefault(e => e.RequestIdentifier == requestIdentifier);
        }

        public VposBoricaRequest GetVposBoricaRequestByRequestIdentifier(string requestIdentifier, DateTime startDate, DateTime endDate)
        {
            return this.unitOfWork.DbContext.Set<VposBoricaRequest>()
                .SingleOrDefault(e => e.RequestIdentifier == requestIdentifier && startDate <= e.CreateDate && e.CreateDate <= endDate);
        }

        public bool VposEpayInvoiceNumberExists(string invoiceNo)
        {
            return this.unitOfWork.DbContext.Set<VposEpayRequest>().Any(e => e.TransactionInvoiceNo == invoiceNo);
        }

        public VposEpayRequest GetVposEpayRequestByInvoiceNo(string invoiceNo)
        {
            return
                this.unitOfWork.DbContext.Set<VposEpayRequest>()
                .Include(e => e.PaymentRequest)
                .Include(e => e.PaymentRequest.ObligationType)
                .SingleOrDefault(e => e.TransactionInvoiceNo == invoiceNo);
        }

        public VposEpayRequest GetVposEpayRequestById(int vposEpayRequestId)
        {
            return
                this.unitOfWork.DbContext.Set<VposEpayRequest>()
                .Include(e => e.PaymentRequest)
                .SingleOrDefault(e => e.VposEpayRequestId == vposEpayRequestId);
        }

        public List<EserviceClient> GetBoricaEserviceClientWarnings()
        {
            var warningThreshold = DateTime.Now.AddDays(30);

            return
                this.unitOfWork.DbContext.Set<EserviceClient>()
                .Where(e =>
                    e.VposClientId == (int)Vpos.Borica &&
                    e.BoricaVposRequestSignCertValidTo != null &&
                    warningThreshold >= e.BoricaVposRequestSignCertValidTo.Value &&
                    (e.BoricaVposRequestSignCertExpHideAdminMsg == null || e.BoricaVposRequestSignCertExpHideAdminMsg == false))
                .ToList();
        }

        private IQueryable<PendingRequestVO> GetPendingRequestsByUinCustomer(string uin,
            PendingPaymentColumn sortBy,
            bool sortDescending)
        {
            IQueryable<PendingRequestVO> query =
                                (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>().Include(p => p.ObligationType)
                                 join acInitiator in this.unitOfWork.DbContext.Set<EserviceClient>() on pr.InitiatorId equals acInitiator.EserviceClientId
                                 join ac in this.unitOfWork.DbContext.Set<EserviceClient>() on pr.EserviceClientId equals ac.EserviceClientId
                                 join vpc in this.unitOfWork.DbContext.Set<VposClient>() on ac.VposClientId equals vpc.VposClientId into g1
                                 from vpc in g1.DefaultIfEmpty()
                                 where pr.ApplicantUin == uin && pr.PaymentRequestStatusId == PaymentRequestStatus.Pending && !pr.IsTempRequest
                                 select new PendingRequestVO()
                                 {
                                     Gid = pr.Gid,
                                     ApplicantName = pr.ApplicantName,
                                     PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                                     CreateDate = pr.CreateDate,
                                     ExpirationDate = pr.ExpirationDate,
                                     ServiceProviderName = pr.ServiceProviderName,
                                     PaymentReason = pr.PaymentReason,
                                     PaymentAmount = pr.PaymentAmount,
                                     IsEpayVposEnabled = ac.IsEpayVposEnabled,
                                     IsCvposEnabled = ac.IsBoricaVposEnabled,
                                     ObligationType = pr.ObligationType.Name,
                                     Vpos = vpc != null && vpc.IsActive ? (Vpos)vpc.VposClientId : (Vpos?)null,
                                     InitiatorName = acInitiator.ServiceName,
                                     IsInProgress = pr.PaymentRequestStatusId == PaymentRequestStatus.InProcess,
                                     PayOrder = pr.PayOrder,
                                     ObligationTypeId = pr.ObligationTypeId,
                                     EserviceClientId = pr.EserviceClientId,
                                     AdditionalInfo = pr.AdditionalInformation,
                                     AlgorithmId = pr.ObligationType.AlgorithmId
                                 });

            query = GetSortablePendingRequestQuery(query);

            return query;
        }

        private Expression<Func<PaymentRequest, bool>> CreateSearchProcessedRequestPredicate(
            string uin,
            string filterPaymentIdentifier,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            PaymentRequestStatus? filterRequestStatus)
        {
            var predicate = PredicateBuilder.True<PaymentRequest>();

            predicate = predicate.And(e =>
                e.ApplicantUin == uin &&
                e.PaymentRequestStatusId != PaymentRequestStatus.Pending && e.PaymentRequestStatusId != PaymentRequestStatus.Expired &&
                (!e.IsTempRequest || e.PaymentRequestStatusId == PaymentRequestStatus.Authorized || e.PaymentRequestStatusId == PaymentRequestStatus.Paid));

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

            if (filterRequestStatus.HasValue && filterRequestStatus.Value != PaymentRequestStatus.Pending)
            {
                predicate = predicate.And(e => e.PaymentRequestStatusId == filterRequestStatus.Value);
            }

            return predicate;
        }

        private Expression<Func<PaymentRequest, bool>> CreateSearchEserviceAdminRequestPredicate(
            int eserviceClientId,
            string filterPaymentIdentifier,
            string filterReferencefNumber,
            string filterRefid,
            string filterPeymentType,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            PaymentRequestStatus? filterRequestStatus,
            ObligationStatusEnum? filterObligationStatus,
            string applicantName,
            string applicantUin)
        {
            var predicate = PredicateBuilder.True<PaymentRequest>();

            predicate = predicate.And(e => e.EserviceClientId == eserviceClientId);

            if (!String.IsNullOrWhiteSpace(filterPaymentIdentifier))
            {
                predicate = predicate.And(e => e.PaymentRequestIdentifier == filterPaymentIdentifier);
            }
            if (!String.IsNullOrWhiteSpace(filterReferencefNumber))
            {
                predicate = predicate.And(e => e.PaymentReferenceNumber == filterReferencefNumber);
            }
            if (!String.IsNullOrWhiteSpace(filterRefid))
            {
                predicate = predicate.And(e => e.DistributionRevenuePayment.DistributionRevenueId.ToString() == filterRefid);
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

            if (filterRequestStatus.HasValue)
            {
                predicate = predicate.And(e => e.PaymentRequestStatusId == filterRequestStatus.Value);
            }

            if (filterObligationStatus.HasValue)
            {
                predicate = predicate.And(e => e.ObligationStatusId == filterObligationStatus.Value);
            }

            if (!String.IsNullOrWhiteSpace(filterPeymentType))
            {
                predicate = predicate.And(e => e.PaymentReferenceType == filterPeymentType);
            }

            if (!String.IsNullOrWhiteSpace(applicantUin))
            {
                predicate = predicate.And(e => e.ApplicantUin == applicantUin);
            }

            if (!String.IsNullOrWhiteSpace(applicantName))
            {
                predicate = predicate.AndStringContains(e => e.ApplicantName, applicantName);
            }

            if (!String.IsNullOrWhiteSpace(applicantUin))
            {
                predicate = predicate.And(e => e.ApplicantUin == applicantUin);
            }

            return predicate;
        }

        private Expression<Func<EserviceClient, bool>> CreateEserviceRecordsPredicate(
            string aisName,
            string departmentName,
            int? vposClientId,
            int? isActiveId)
        {
            var predicate = PredicateBuilder.True<EserviceClient>();

            if (!String.IsNullOrWhiteSpace(aisName))
            {
                predicate = predicate.AndStringContains(e => e.AisName, aisName);
            }

            if (!String.IsNullOrWhiteSpace(departmentName))
            {
                predicate = predicate.AndStringContains(e => e.Department.Name, departmentName);
            }

            if (vposClientId.HasValue)
            {
                predicate = predicate.And(e => e.VposClientId == vposClientId.Value);
            }

            if (isActiveId.HasValue)
            {
                predicate = predicate.And(e => e.IsActive == (isActiveId.Value == 1 ? true : false));
            }

            return predicate;
        }


        private Expression<Func<ObligationType, bool>> CreateObligationTypeRecordsPredicate(
           string oblName,
           int? isActiveId)
        {
            var predicate = PredicateBuilder.True<ObligationType>();

            if (!String.IsNullOrWhiteSpace(oblName))
            {
                predicate = predicate.AndStringContains(e => e.Name, oblName);
            }

            if (isActiveId.HasValue)
            {
                predicate = predicate.And(e => e.IsActive == (isActiveId.Value == 1 ? true : false));
            }

            return predicate;
        }

        private Expression<Func<Department, bool>> CreateDepartmentRecordsPredicate(
            string departmentName,
            string departmentUniqueIdentificationNumber,
            string departmentUnifiedBudgetClassifier)
        {
            var predicate = PredicateBuilder.True<Department>();

            if (!String.IsNullOrWhiteSpace(departmentName))
            {
                predicate = predicate.AndStringContains(d => d.Name, departmentName);
            }

            if (!String.IsNullOrWhiteSpace(departmentUniqueIdentificationNumber))
            {
                predicate = predicate.AndStringContains(d => d.UniqueIdentificationNumber, departmentUniqueIdentificationNumber);
            }

            if (!String.IsNullOrWhiteSpace(departmentUnifiedBudgetClassifier))
            {
                predicate = predicate.AndStringContains(d => d.UnifiedBudgetClassifier, departmentUnifiedBudgetClassifier);
            }

            return predicate;
        }

        private Expression<Func<TransactionRecord, bool>> CreateSearchTransactionRecordsPredicate(
            int eserviceBankAccountId,
            DateTime? transactionAccountingDateFrom,
            DateTime? transactionAccountingDateTo,
            decimal? transactionAmountFrom,
            decimal? transactionAmountTo,
            string infoDocumentNumber,
            DateTime? infoDocumentDateFrom,
            DateTime? infoDocumentDateTo,
            string infoSenderIban,
            string infoSenderName,
            string infoDebtorName,
            string infoDebtorBulstatEgnLnch,
            string infoPaymentReason,
            string infoAC1AuthorizationCode,
            int? transactionRecordPaymentMethod,
            int? transactionRecordRefStatus)
        {
            var predicate = PredicateBuilder.True<TransactionRecord>();

            predicate = predicate.And(e => e.TransactionFile.EserviceBankAccountId == eserviceBankAccountId);

            if (transactionAccountingDateFrom.HasValue)
            {
                DateTime fromDate = transactionAccountingDateFrom.Value.Date;
                predicate = predicate.And(e => e.TransactionAccountingDate >= fromDate);
            }

            if (transactionAccountingDateTo.HasValue)
            {
                DateTime toDate = transactionAccountingDateTo.Value.Date.AddDays(1).AddMilliseconds(-1);
                predicate = predicate.And(e => e.TransactionAccountingDate <= toDate);
            }

            if (transactionAmountFrom.HasValue)
            {
                predicate = predicate.And(e => (e.TransactionAmount.HasValue ? e.TransactionAmount.Value : 0) >= transactionAmountFrom.Value);
            }

            if (transactionAmountTo.HasValue)
            {
                predicate = predicate.And(e => (e.TransactionAmount.HasValue ? e.TransactionAmount.Value : 0) <= transactionAmountTo.Value);
            }

            if (!String.IsNullOrWhiteSpace(infoDocumentNumber))
            {
                predicate = predicate.And(e => e.InfoDocumentNumber == infoDocumentNumber);
            }

            if (infoDocumentDateFrom.HasValue)
            {
                DateTime fromDate = infoDocumentDateFrom.Value.Date;
                predicate = predicate.And(e => e.InfoDocumentDate >= fromDate);
            }

            if (infoDocumentDateTo.HasValue)
            {
                DateTime toDate = infoDocumentDateTo.Value.Date.AddDays(1).AddMilliseconds(-1);
                predicate = predicate.And(e => e.InfoDocumentDate <= toDate);
            }

            if (!String.IsNullOrWhiteSpace(infoSenderIban))
            {
                predicate = predicate.And(e => e.InfoSenderIban == infoSenderIban);
            }

            if (!String.IsNullOrWhiteSpace(infoSenderName))
            {
                predicate = predicate.AndStringContains(e => e.InfoSenderName, infoSenderName);
            }

            if (!String.IsNullOrWhiteSpace(infoDebtorName))
            {
                predicate = predicate.AndStringContains(e => e.InfoDebtorName, infoDebtorName);
            }

            if (!String.IsNullOrWhiteSpace(infoDebtorBulstatEgnLnch))
            {
                predicate = predicate.And(e =>
                    e.InfoDebtorBulstat == infoDebtorBulstatEgnLnch ||
                    e.InfoDebtorEgn == infoDebtorBulstatEgnLnch ||
                    e.InfoDebtorLnch == infoDebtorBulstatEgnLnch);
            }

            if (!String.IsNullOrWhiteSpace(infoPaymentReason))
            {
                predicate = predicate.AndStringContains(e => e.InfoPaymentReason, infoPaymentReason);
            }

            if (!String.IsNullOrWhiteSpace(infoAC1AuthorizationCode))
            {
                predicate = predicate.And(e => e.InfoAC1AuthorizationCode == infoAC1AuthorizationCode);
            }

            if (transactionRecordPaymentMethod.HasValue)
            {
                predicate = predicate.And(e => e.TransactionRecordPaymentMethodId == (TransactionRecordPaymentMethod)transactionRecordPaymentMethod.Value);
            }

            if (transactionRecordRefStatus.HasValue)
            {
                if (transactionRecordRefStatus.Value == -1)
                {
                    predicate = predicate.And(e =>
                        e.TransactionRecordRefStatusId == TransactionRecordRefStatus.ReferencedSuccessfully ||
                        e.TransactionRecordRefStatusId == TransactionRecordRefStatus.ReferencedOverpaidAmount ||
                        e.TransactionRecordRefStatusId == TransactionRecordRefStatus.ReferencedInsufficientAmount);
                }
                else
                {
                    predicate = predicate.And(e => e.TransactionRecordRefStatusId == (TransactionRecordRefStatus)transactionRecordRefStatus.Value);
                }
            }

            return predicate;
        }

        private IQueryable<PendingRequestVO> GetSortablePendingRequestQuery(IQueryable<PendingRequestVO> sourceQuery, PendingPaymentColumn sortBy, bool sortDescending)
        {
            IQueryable<PendingRequestVO> sortalbleQuery = null;

            if (sortBy == PendingPaymentColumn.PaymentId)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentRequestIdentifier);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentRequestIdentifier);
                }
            }
            else if (sortBy == PendingPaymentColumn.CreateDate)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.CreateDate);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.CreateDate);
                }
            }
            else if (sortBy == PendingPaymentColumn.ExpirationDate)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.ExpirationDate);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.ExpirationDate);
                }
            }
            else if (sortBy == PendingPaymentColumn.ServiceProvider)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.ServiceProviderName);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.ServiceProviderName);
                }
            }
            else if (sortBy == PendingPaymentColumn.PaymentReason)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentReason);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentReason);
                }
            }
            else if (sortBy == PendingPaymentColumn.Amount)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentAmount);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentAmount);
                }
            }
            else //if sortBy has invalid value
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentRequestIdentifier);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentRequestIdentifier);
                }
            }

            return sortalbleQuery;
        }

        private IQueryable<PendingRequestVO> GetSortablePendingRequestQuery(IQueryable<PendingRequestVO> sourceQuery)
        {
            IQueryable<PendingRequestVO> sortalbleQuery = null;

            sortalbleQuery = sourceQuery
                    .OrderBy(e => e.ObligationTypeId)
                    .ThenBy(e => e.CreateDate);

            return sortalbleQuery;
        }

        private IQueryable<ProcessedRequestVO> GetSortableProcessedRequestQuery(IQueryable<ProcessedRequestVO> sourceQuery, ProcessedPaymentColumn sortBy, bool sortDescending)
        {
            IQueryable<ProcessedRequestVO> sortalbleQuery = null;

            if (sortBy == ProcessedPaymentColumn.PaymentId)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentRequestIdentifier);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentRequestIdentifier);
                }
            }
            else if (sortBy == ProcessedPaymentColumn.Date)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.TransactionDate);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.TransactionDate);
                }
            }
            else if (sortBy == ProcessedPaymentColumn.ServiceProvider)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.ServiceProviderName);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.ServiceProviderName);
                }
            }
            else if (sortBy == ProcessedPaymentColumn.PaymentReason)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentReason);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentReason);
                }
            }
            else if (sortBy == ProcessedPaymentColumn.Amount)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.CreateDate < CurrencyHelper.EuroAcceptanceDate ? e.PaymentAmountRequest / CurrencyHelper.BgnToEuroRate : e.PaymentAmountRequest);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.CreateDate < CurrencyHelper.EuroAcceptanceDate ? e.PaymentAmountRequest / CurrencyHelper.BgnToEuroRate : e.PaymentAmountRequest);
                }
            }
            else if (sortBy == ProcessedPaymentColumn.Status)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentRequestStatusId);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentRequestStatusId);
                }
            }
            else //if sortBy has invalid value
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentRequestIdentifier);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentRequestIdentifier);
                }
            }

            return sortalbleQuery;
        }

        private IQueryable<ProcessedRequestVO> GetSortableEserviceAdminRequestQuery(IQueryable<ProcessedRequestVO> sourceQuery, EserviceAdminRequestColumn sortBy, bool sortDescending)
        {
            IQueryable<ProcessedRequestVO> sortalbleQuery = null;

            if (sortBy == EserviceAdminRequestColumn.PaymentId)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentRequestIdentifier);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentRequestIdentifier);
                }
            }
            else if (sortBy == EserviceAdminRequestColumn.Date)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.TransactionDate);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.TransactionDate);
                }
            }
            else if (sortBy == EserviceAdminRequestColumn.ServiceProvider)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.ServiceProviderName);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.ServiceProviderName);
                }
            }
            else if (sortBy == EserviceAdminRequestColumn.PaymentReason)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentReason);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentReason);
                }
            }
            else if (sortBy == EserviceAdminRequestColumn.Amount)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.CreateDate < CurrencyHelper.EuroAcceptanceDate ? e.PaymentAmountRequest / CurrencyHelper.BgnToEuroRate : e.PaymentAmountRequest);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.CreateDate < CurrencyHelper.EuroAcceptanceDate ? e.PaymentAmountRequest / CurrencyHelper.BgnToEuroRate : e.PaymentAmountRequest);
                }
            }
            else if (sortBy == EserviceAdminRequestColumn.Status)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentRequestStatusId);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentRequestStatusId);
                }
            }
            else if (sortBy == EserviceAdminRequestColumn.ApplicantName)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.ApplicantName);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.ApplicantName);
                }
            }
            else if (sortBy == EserviceAdminRequestColumn.ApplicantUin)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.ApplicantUin);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.ApplicantUin);
                }
            }
            else if (sortBy == EserviceAdminRequestColumn.ObligationType)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.ObligationType);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.ObligationType);
                }
            }
            else if (sortBy == EserviceAdminRequestColumn.RefIf)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.Refid);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.Refid);
                }
            }
            else if (sortBy == EserviceAdminRequestColumn.PaymentReferenceNumber)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentReferenceNumber);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentReferenceNumber);
                }
            }
            else if (sortBy == EserviceAdminRequestColumn.ObligationStatus)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.ObligationStatusId);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.ObligationStatusId);
                }
            }
            else if (sortBy == EserviceAdminRequestColumn.ExpirationDate)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.ExpirationDate);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.ExpirationDate);
                }
            }
            else //if sortBy has invalid value
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentRequestIdentifier);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentRequestIdentifier);
                }
            }

            return sortalbleQuery;
        }

        private IQueryable<RequestAccessVO> GetSortableRequestAccessByUinQuery(IQueryable<RequestAccessVO> sourceQuery, RequestAccessListColumn sortBy, bool sortDescending)
        {
            IQueryable<RequestAccessVO> sortalbleQuery = null;

            if (sortBy == RequestAccessListColumn.PaymentRequestIdentifier)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentRequestIdentifier);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentRequestIdentifier);
                }
            }
            else if (sortBy == RequestAccessListColumn.AccessDate)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.AccessDate);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.AccessDate);
                }
            }
            else if (sortBy == RequestAccessListColumn.AccessCount)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.AccessCount);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.AccessCount);
                }
            }
            else if (sortBy == RequestAccessListColumn.ServiceProvider)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.ServiceProviderName);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.ServiceProviderName);
                }
            }
            else if (sortBy == RequestAccessListColumn.PaymentReason)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.PaymentReason);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.PaymentReason);
                }
            }
            else //if sortBy has invalid value
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.AccessDate);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.AccessDate);
                }
            }

            return sortalbleQuery;
        }

        private IQueryable<RequestAccessVO> GetRequestAccessByUinQuery(string uin)
        {
            return
                from al in this.unitOfWork.DbContext.Set<EbankingAccessLog>()
                join pr in this.unitOfWork.DbContext.Set<PaymentRequest>() on al.PaymentRequestId equals pr.PaymentRequestId
                where pr.ApplicantUin == uin
                group al by new { al.PaymentRequestId, pr.Gid, pr.ServiceProviderName, pr.PaymentReason, pr.PaymentRequestIdentifier, pr.CreateDate, pr.PaymentRequestStatusId } into g
                select new RequestAccessVO
                {
                    PaymentRequestGid = g.Key.Gid,
                    PaymentRequestIdentifier = g.Key.PaymentRequestIdentifier,
                    PaymentReqestCreateDate = g.Key.CreateDate,
                    PaymentRequestStatusId = g.Key.PaymentRequestStatusId,
                    AccessDate = g.Max(e => e.AccessDate),
                    ServiceProviderName = g.Key.ServiceProviderName,
                    PaymentReason = g.Key.PaymentReason,
                    AccessCount = g.Count()
                };
        }

        private IQueryable<TransactionRecordVO> GetSortableTransactionRecordsQuery(IQueryable<TransactionRecordVO> sourceQuery, TransactionListColumn sortBy, bool sortDescending)
        {
            IQueryable<TransactionRecordVO> sortalbleQuery = null;

            if (sortBy == TransactionListColumn.TransactionAccountingDate)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.TransactionAccountingDate);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.TransactionAccountingDate);
                }
            }
            else if (sortBy == TransactionListColumn.TransactionAmount)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.TransactionAmount);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.TransactionAmount);
                }
            }
            else if (sortBy == TransactionListColumn.InfoDocumentNumberDate)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.InfoDocumentNumberDate);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.InfoDocumentNumberDate);
                }
            }
            else if (sortBy == TransactionListColumn.InfoSenderIbanName)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.InfoSenderIbanName);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.InfoSenderIbanName);
                }
            }
            else if (sortBy == TransactionListColumn.InfoDebtorBulstatEgnLnchName)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.InfoDebtorBulstatEgnLnchName);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.InfoDebtorBulstatEgnLnchName);
                }
            }
            else if (sortBy == TransactionListColumn.TransactionRecordPaymentMethod)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.TransactionRecordPaymentMethodId);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.TransactionRecordPaymentMethodId);
                }
            }
            else if (sortBy == TransactionListColumn.TransactionRecordRefStatus)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.TransactionRecordRefStatusId);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.TransactionRecordRefStatusId);
                }
            }
            else //if sortBy has invalid value
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.TransactionAccountingDate);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.TransactionAccountingDate);
                }
            }

            return sortalbleQuery;
        }

        private IQueryable<EserviceRecordVO> GetSortableEserviceRecordsQuery(IQueryable<EserviceRecordVO> sourceQuery, EserviceListColumn sortBy, bool sortDescending)
        {
            IQueryable<EserviceRecordVO> sortalbleQuery = null;

            if (sortBy == EserviceListColumn.EserviceClientId)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.EserviceClientId);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.EserviceClientId);
                }
            }
            else if (sortBy == EserviceListColumn.AisName)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.AisName);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.AisName);
                }
            }
            else if (sortBy == EserviceListColumn.DepartmentName)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.DepartmentName);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.DepartmentName);
                }
            }
            else if (sortBy == EserviceListColumn.AccountBank)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.AccountBank);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.AccountBank);
                }
            }
            else if (sortBy == EserviceListColumn.VposClientName)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.VposClientId);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.VposClientId);
                }
            }
            else if (sortBy == EserviceListColumn.IsActive)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.IsActive);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.IsActive);
                }
            }

            return sortalbleQuery;
        }

        private IQueryable<ObligationTypeVO> GetSortableObligationTypeRecordsQuery(IQueryable<ObligationTypeVO> sourceQuery, ObligationTypeListColumn sortBy, bool sortDescending)
        {
            IQueryable<ObligationTypeVO> sortalbleQuery = null;

            if (sortBy == ObligationTypeListColumn.ObligationTypeId)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.ObligationTypeId);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.ObligationTypeId);
                }
            }
            else if (sortBy == ObligationTypeListColumn.Name)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.Name);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.Name);
                }
            }
            else if (sortBy == ObligationTypeListColumn.IsActive)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(e => e.IsActive);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(e => e.IsActive);
                }
            }

            return sortalbleQuery;
        }


        private IQueryable<DepartmentVO> GetSortableDepartmentRecordsQuery(IQueryable<DepartmentVO> sourceQuery, DepartmentListColumn sortBy, bool sortDescending)
        {
            IQueryable<DepartmentVO> sortalbleQuery = null;

            if (sortBy == DepartmentListColumn.DepartmentId)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(d => d.DepartmentId);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(d => d.DepartmentId);
                }
            }
            else if (sortBy == DepartmentListColumn.DepartmentName)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(d => d.Name);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(d => d.Name);
                }
            }
            else if (sortBy == DepartmentListColumn.DepartmentUniqueIdentificationNumber)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(d => d.UniqueIdentificationNumber);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(d => d.UniqueIdentificationNumber);
                }
            }
            else if (sortBy == DepartmentListColumn.DepartmentUnifiedBudgetClassifier)
            {
                if (sortDescending)
                {
                    sortalbleQuery = sourceQuery.OrderByDescending(d => d.UnifiedBudgetClassifier);
                }
                else
                {
                    sortalbleQuery = sourceQuery.OrderBy(d => d.UnifiedBudgetClassifier);
                }
            }

            return sortalbleQuery;
        }

        private IQueryable<RequestAccessDetailsVO> GetRequestAccessDetailsQuery(int paymentRequestId, string uin)
        {
            return
                from al in this.unitOfWork.DbContext.Set<EbankingAccessLog>()
                join pr in this.unitOfWork.DbContext.Set<PaymentRequest>() on al.PaymentRequestId equals pr.PaymentRequestId
                join bc in this.unitOfWork.DbContext.Set<EbankingClient>() on al.EbankingClientId equals bc.EbankingClientId
                where al.PaymentRequestId == paymentRequestId && pr.ApplicantUin == uin
                select new RequestAccessDetailsVO()
                {
                    AccessDate = al.AccessDate,
                    IpAddress = al.IpAddress,
                    EbankingClientName = bc.Name
                };
        }

        private PaymentOrderVO GetPaymentOrderInternal(Guid? gid, string uin, string accessCode)
        {
            IQueryable<PaymentRequest> query;
            if (String.IsNullOrWhiteSpace(accessCode))
            {
                if (uin == null)
                {
                    query =
                        from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                        join cl in this.unitOfWork.DbContext.Set<EserviceClient>()
                        on pr.EserviceClientId equals cl.EserviceClientId
                        join ot in this.unitOfWork.DbContext.Set<ObligationType>()
                        on pr.ObligationTypeId equals ot.ObligationTypeId
                        where pr.Gid == gid.Value
                        select pr;
                }
                else
                {
                    query =
                        from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                        join cl in this.unitOfWork.DbContext.Set<EserviceClient>()
                        on pr.EserviceClientId equals cl.EserviceClientId
                        join ot in this.unitOfWork.DbContext.Set<ObligationType>()
                        on pr.ObligationTypeId equals ot.ObligationTypeId
                        where pr.Gid == gid.Value && pr.ApplicantUin == uin
                        select pr;
                }
            }
            else
            {
                query =
                    from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                    join cl in this.unitOfWork.DbContext.Set<EserviceClient>()
                    on pr.EserviceClientId equals cl.EserviceClientId
                    join ot in this.unitOfWork.DbContext.Set<ObligationType>()
                    on pr.ObligationTypeId equals ot.ObligationTypeId
                    where pr.PaymentRequestAccessCode == accessCode
                    select pr;
            }

            return (from pr in query
                    select new PaymentOrderVO()
                    {
                        Gid = pr.Gid,
                        ServiceProviderName = pr.ServiceProviderName,
                        BankName = pr.ServiceProviderBank,
                        IBAN = pr.ServiceProviderIBAN,
                        BIC = pr.ServiceProviderBIC,
                        PaymentAmount = pr.PaymentAmount,
                        PaymentReason = pr.PaymentReason,
                        DocumentType = pr.PaymentReferenceType,
                        DocumentNumber = pr.PaymentReferenceNumber,
                        DocumentDate = pr.PaymentReferenceDate,
                        CreateDate = pr.CreateDate,
                        ApplicantName = pr.ApplicantName,
                        ApplicantUinTypeId = pr.ApplicantUinTypeId,
                        ApplicantUin = pr.ApplicantUin,
                        IsVposAuthorized = pr.IsVposAuthorized,
                        VposAuthorizationId = pr.VposAuthorizationId,
                        PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                        PaymentRequestStatusChangeTime = pr.PaymentRequestStatusChangeTime,
                        PaymentRequestStatusId = pr.PaymentRequestStatusId,
                        ObligationType = pr.ObligationType.Name,
                        ObligationTypeCode = pr.ObligationType.PaymentTypeCode,
                        ObligationTypeAlgorithmId = pr.ObligationType.AlgorithmId,
                        AdditionalInformation = pr.AdditionalInformation
                    })
                   .SingleOrDefault();
        }

        private VposRequestDataVO GetVposRequestDataInternal(Guid? gid, string uin, string accessCode)
        {
            //TODO: must be refactor to use IQueryable
            if (String.IsNullOrWhiteSpace(accessCode))
            {
                return
                    (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                     join ec in this.unitOfWork.DbContext.Set<EserviceClient>() on pr.EserviceClientId equals ec.EserviceClientId
                     join vpc in this.unitOfWork.DbContext.Set<VposClient>() on ec.VposClientId equals vpc.VposClientId into g1
                     from vpc in g1.DefaultIfEmpty()
                     where
                         pr.Gid == gid.Value &&
                         pr.ApplicantUin == uin &&
                         pr.PaymentRequestStatusId == PaymentRequestStatus.Pending &&
                         (ec.VposClientId == null || vpc.IsActive)
                     select new VposRequestDataVO()
                     {
                         PaymentRequestId = pr.PaymentRequestId,
                         PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                         PaymentRequestGid = pr.Gid,
                         PaymentRequestStatusId = pr.PaymentRequestStatusId,
                         PaymentAmount = pr.PaymentAmount,
                         PaymentReason = pr.PaymentReason,
                         VposPaymentRequestUri = vpc.PaymentRequestUrl,
                         VposClientId = vpc.VposClientId,
                         ЕserviceClientGid = ec.Gid,
                         DskVposMerchantId = ec.DskVposMerchantId,
                         DskVposMerchantPassword = ec.DskVposMerchantPassword,
                         BoricaVposTerminalId = ec.BoricaVposTerminalId,
                         BoricaVposBOReqSignCertFileName = ec.BoricaVposRequestSignCertFileName,
                         BoricaVposBOReqSignCertPassword = ec.BoricaVposRequestSignCertPassword,
                         FiBankVposAccessKeystoreFilename = ec.FiBankVposAccessKeystoreFilename,
                         FiBankVposAccessKeystorePassword = ec.FiBankVposAccessKeystorePassword
                     })
                     .SingleOrDefault();
            }
            else
            {
                return
                    (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                     join ec in this.unitOfWork.DbContext.Set<EserviceClient>() on pr.EserviceClientId equals ec.EserviceClientId
                     join vpc in this.unitOfWork.DbContext.Set<VposClient>() on ec.VposClientId equals vpc.VposClientId into g1
                     from vpc in g1.DefaultIfEmpty()
                     where
                         pr.PaymentRequestAccessCode == accessCode &&
                         pr.PaymentRequestStatusId == PaymentRequestStatus.Pending &&
                         (ec.VposClientId == null || vpc.IsActive)
                     select new VposRequestDataVO()
                     {
                         PaymentRequestId = pr.PaymentRequestId,
                         PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                         PaymentRequestGid = pr.Gid,
                         PaymentRequestStatusId = pr.PaymentRequestStatusId,
                         PaymentAmount = pr.PaymentAmount,
                         PaymentReason = pr.PaymentReason,
                         VposPaymentRequestUri = vpc.PaymentRequestUrl,
                         VposClientId = vpc.VposClientId,
                         ЕserviceClientGid = ec.Gid,
                         DskVposMerchantId = ec.DskVposMerchantId,
                         DskVposMerchantPassword = ec.DskVposMerchantPassword,
                         BoricaVposTerminalId = ec.BoricaVposTerminalId,
                         BoricaVposBOReqSignCertFileName = ec.BoricaVposRequestSignCertFileName,
                         BoricaVposBOReqSignCertPassword = ec.BoricaVposRequestSignCertPassword,
                         FiBankVposAccessKeystoreFilename = ec.FiBankVposAccessKeystoreFilename,
                         FiBankVposAccessKeystorePassword = ec.FiBankVposAccessKeystorePassword
                     })
                     .SingleOrDefault();
            }
        }

        public List<PaymentRequest> GetPendingPaymentRequestByUid(Guid[] guids)
        {
            return this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(pr => guids.Contains(pr.Gid) && pr.PaymentRequestStatusId == PaymentRequestStatus.Pending)
                .ToList();
        }

        public IList<PendingRequestVO> GetAllPendingRequestsByUin(string uin,
            PendingPaymentColumn sortBy,
            bool sortDescending)
        {
            IQueryable<PendingRequestVO> query =
                from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                join ac in this.unitOfWork.DbContext.Set<EserviceClient>() on pr.EserviceClientId equals ac.EserviceClientId
                join vpc in this.unitOfWork.DbContext.Set<VposClient>() on ac.VposClientId equals vpc.VposClientId into g1
                from vpc in g1.DefaultIfEmpty()
                where pr.ApplicantUin == uin && pr.PaymentRequestStatusId == PaymentRequestStatus.Pending && !pr.IsTempRequest
                select new PendingRequestVO()
                {
                    Gid = pr.Gid,
                    ApplicantName = pr.ApplicantName,
                    PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                    CreateDate = pr.CreateDate,
                    ExpirationDate = pr.ExpirationDate,
                    ServiceProviderName = pr.ServiceProviderName,
                    PaymentReason = pr.PaymentReason,
                    PaymentAmount = pr.PaymentAmount,
                    IsEpayVposEnabled = ac.IsEpayVposEnabled,
                    ObligationType = pr.ObligationType.Name,
                    ObligationTypeId = pr.ObligationTypeId,
                    IsCvposEnabled = ac.IsBoricaVposEnabled,
                    Vpos = vpc != null && vpc.IsActive ? (Vpos)vpc.VposClientId : (Vpos?)null
                };

            query = GetSortablePendingRequestQuery(query);

            return query.ToList();
        }

        public IList<ProcessedRequestVO> GetAllProcessedRequestsByUin(string uin,
            string filterPaymentIdentifier,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            PaymentRequestStatus? filterRequestStatus,
            ProcessedPaymentColumn sortBy,
            bool sortDescending)
        {
            var predicate = CreateSearchProcessedRequestPredicate(
            uin,
            filterPaymentIdentifier,
            filterDateFrom,
            filterDateTo,
            filterAmountFrom,
            filterAmountTo,
            filterServiceProvider,
            filterPaymentReason,
            filterRequestStatus);

            IQueryable<ProcessedRequestVO> query =
                this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(predicate)
                .Select(pr =>
                    new ProcessedRequestVO()
                    {
                        Gid = pr.Gid,
                        ApplicantName = pr.ApplicantName,
                        ApplicantUin = pr.ApplicantUin,
                        PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                        TransactionDate = pr.PaymentRequestStatusChangeTime,
                        ServiceProviderName = pr.ServiceProviderName,
                        PaymentReason = pr.PaymentReason,
                        PaymentAmountRequest = pr.PaymentAmount,
                        PaymentRequestGid = pr.Gid,
                        PaymentRequestStatusId = pr.PaymentRequestStatusId,
                        ExpirationDate = pr.ExpirationDate,
                        ObligationType = pr.ObligationType.Name,
                        ObligationStatusId = pr.ObligationStatusId,
                        PaymentReferenceType = pr.PaymentReferenceType,
                        Refid = pr.DistributionRevenuePayment.DistributionRevenueId.ToString()
                    });

            query = GetSortableProcessedRequestQuery(query, sortBy, sortDescending);

            return query.ToList();
        }
    }
}