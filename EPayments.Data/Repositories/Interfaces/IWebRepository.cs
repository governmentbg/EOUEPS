using EPayments.Data.ViewObjects.Admin;
using EPayments.Data.ViewObjects.Web;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;

namespace EPayments.Data.Repositories.Interfaces
{
    public interface IWebRepository : IBaseRepository
    {
        int CountPendingRequestsByUin(string uin);

        int CountProcessedRequestsByUin(
            string uin,
            string filterPaymentIdentifier,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            PaymentRequestStatus? filterRequestStatus);

        int CountEserviceAdminRequests(
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
            string applicantUin);

        Tuple<int, decimal> CountAndGetTotalAmountForTransactionRecords(
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
            int? transactionRecordRefStatus);

        int CountEserviceRecords(
            string aisName,
            string departmentName,
            int? vposClientId,
            int? isActiveId);

        int CountDepartmentRecords(
            string departmentName,
            string departmentUniqueIdentificationNumber,
            string departmentUnifiedBudgetClassifier);
        int CountObligationTypeRecords(
            string name,
            int? isActiveId);

        EserviceClient GetEserviceClient(int eserviceClientId);

        Department GetDepartment(int departmentId);

        ObligationType GetObligationType(int obligationTypeId);

        EserviceAdminUser GetEserviceAdminUserByReferringEserviceClientId(int referringEserviceClientId);

        bool EserviceAdminUserExists(string username);

        bool DepartmentHasReferencedClients(int departmentId, int excludeEserviceClientId);

        List<EserviceBankAccount> GetEserviceBankAccountsByAdminUserId(int eserviceAdminUserId);

        int CountRequestAccessByUin(string uin);

        PaymentRequest GetPaymentRequestById(int paymentRequestId);

        PaymentRequest GetPaymentRequestByGidAndUin(Guid gid, string uin);

        PaymentRequest GetPaymentRequestByIdentifier(string paymentRequestIdentifier);

        PaymentRequest GetPaymentRequestByAccessCode(string accessCode);

        IList<PendingRequestVO> GetPendingRequestsByUin(
            string uin,
            PendingPaymentColumn sortBy,
            bool sortDescending,
            int page,
            int resultsPerPage);

        IList<PendingRequestVO> GetPendingRequestsByUin(
          string uin,
          PendingPaymentColumn sortBy,
          bool sortDescending);

        IList<PendingRequestVO> GetAllPendingRequestsByUin(
            string uin,
            PendingPaymentColumn sortBy,
            bool sortDescending);

        IList<ProcessedRequestVO> GetProcessedRequestsByUin(
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
            int resultsPerPage);

        IList<ProcessedRequestVO> GetAllProcessedRequestsByUin(
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
            bool sortDescending);

        IList<ProcessedRequestVO> GetEserviceAdminRequests(
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
            int takeCount = int.MaxValue);

        IList<RequestAccessVO> GetRequestAccessByUin(
           string uin,
           RequestAccessListColumn sortBy,
           bool sortDescending,
           int page,
           int resultsPerPage);

        IList<TransactionRecordVO> GetTransactionRecords(
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
            int? resultsPerPage = null);

        IList<EserviceRecordVO> GetEserviceRecords(
            string aisName,
            string departmentName,
            int? vposClientId,
            int? isActiveId,
            EserviceListColumn sortBy,
            bool sortDescending,
            int? page = null,
            int? resultsPerPage = null);

        IList<ObligationTypeVO> GetObligationTypeRecords(
            string oblName,
            int? isActiveId,
            ObligationTypeListColumn sortBy,
            bool sortDescending,
            int? page = null,
            int? resultsPerPage = null);


        IList<DepartmentVO> GetDepartmentRecords(
            string departmentName,
            string departmentUniqueIdentificationNumber,
            string departmentUnifiedBudgetClassifier,
            DepartmentListColumn sortBy,
            bool sortDescending,
            int? page = null,
            int? resultsPerPage = null);

        int GetRequestAccessDetailsCount(int paymentRequestId, string uin);

        List<RequestAccessDetailsVO> GetRequestAccessDetails(int paymentRequestId, string uin, int limit);

        IList<PendingRequestVO> GetPendingRequestsByAccessCode(string accessCode);

        List<ProcessedRequestVO> GetProcessedRequestsByAccessCode(string accessCode);

        PaymentOrderVO GetPaymentOrderByGid(Guid gid);

        PaymentOrderVO GetPaymentOrderByGidAndUin(Guid gid, string uin);

        PaymentOrderVO GetPaymentOrderByAccessCode(string accessCode);

        TransactionRecord GetTransactionRecordByTransactionRecordId(int transactionRecordId);

        VposRequestDataVO GetVposRequestDataByGidAndUin(Guid gid, string uin);

        VposRequestDataVO GetVposRequestDataByAccessCode(string accessCode);

        string GenerateVposBoricaRequestIdentifier();

        VposBoricaRequest GetVposBoricaRequestByRequestIdentifier(string requestIdentifier);

        VposBoricaRequest GetVposBoricaRequestByRequestIdentifier(string requestIdentifier, DateTime startDate, DateTime endDate);

        bool VposEpayInvoiceNumberExists(string invoiceNo);

        VposEpayRequest GetVposEpayRequestByInvoiceNo(string invoiceNo);

        VposEpayRequest GetVposEpayRequestById(int vposEpayRequestId);

        List<EserviceClient> GetBoricaEserviceClientWarnings();

        List<PaymentRequest> GetPendingPaymentRequestByUid(Guid[] guids);

        List<PaymentRequest> GetPendngPaymentRequestByAisClient(int eServiceClientId);

        List<PaymentRequest> GetPendngAndInProcessPaymentRequestByAisClient(int eServiceClientId, string applicantUin);
    }
}
