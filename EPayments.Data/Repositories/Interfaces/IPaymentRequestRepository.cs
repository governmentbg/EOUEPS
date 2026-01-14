using EPayments.Data.ViewObjects.Admin;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace EPayments.Data.Repositories.Interfaces
{
    public interface IPaymentRequestRepository
    {
        Task ChangePaymentRequestsStatus(string filterPaymentIdentifier,
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
            PaymentRequestStatus newRequestStatus);

        Task<int> CountPaymentRequests(string filterPaymentIdentifier,
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
            string applicantUin);

        Task<List<PaymentRequestVO>> GetPaymentRequests(string filterPaymentIdentifier,
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
            int resultsPerPage);

        Task<List<PaymentRequestVO>> GetAllRequests(string filterPaymentIdentifier,
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
            int takeCount = int.MaxValue);

        List<string> GetAllPaymentReferenceNumbers();

        bool IsPaymentReferenceNumberExists(string referenceNumber);

        List<string> GetPaymentReferenceNumbersForAlgorithm(int algorithmNumber);

        bool IsPaymentReferenceNumberExistsForAlgorithm(int algorithmNumber, string referenceNumber);

        List<PaymentRequest> GetPendingPaymentRequestByPaymentReferenceNumberAndAisPaymentId(string paymentReferenceNumber, string aisPaymentId);

        bool IsAisPaymentIdAlreadyUsed(string aisPaymentId);

        void DeletePaymentRequest(PaymentRequest paymentRequest);
    }
}
