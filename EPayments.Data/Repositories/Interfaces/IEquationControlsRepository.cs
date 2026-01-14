using EPayments.Data.ViewObjects.Admin;
using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace EPayments.Data.Repositories.Interfaces
{
    public interface IEquationControlsRepository
    {
        Task<TotalsBoricaTransactionVO> CountBoricaTransactions(DateTime? dateFrom,
                                                        DateTime? dateTo,
                                                        int? transactionStatus);

        Task<int> CountUndistributetPayments(string filterPaymentIdentifier,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            ObligationStatusEnum? obligationStatus);

        Task<List<UndistributedPaymentRequestVO>> GetUndistributetPayments(string filterPaymentIdentifier,
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
            int resultsPerPage);

        Task<List<PaymentRequestVO>> GetAllUndistributetPayments(string filterPaymentIdentifier,
            DateTime? filterDateFrom,
            DateTime? filterDateTo,
            decimal? filterAmountFrom,
            decimal? filterAmountTo,
            string filterServiceProvider,
            string filterPaymentReason,
            ObligationStatusEnum? obligationStatus,
            string sortBy,
            bool sortDescending,
            int takeCount = int.MaxValue);

        Task<int> CountOldPayments(DateTime? filterDateFrom,
            DateTime? filterDateTo,
            ObligationStatusEnum? obligationStatus);

        Task<List<PaymentRequestVO>> GetOldPayments(DateTime? filterDateFrom,
            DateTime? filterDateTo,
            ObligationStatusEnum? obligationStatus,
            string sortBy,
            bool sortDescending,
            int page,
            int resultsPerPage,
            int takeCount = int.MaxValue);

        Task<List<PaymentRequestVO>> GetAllOldPayments(DateTime? filterDateFrom,
            DateTime? filterDateTo,
            ObligationStatusEnum? obligationStatus,
            string sortBy,
            bool sortDescending,
            int takeCount = int.MaxValue);

        Task<List<BoricaTransactionVO>> GetBoricaTransactions(DateTime? dateFrom,
            DateTime? dateTo,
            int? transactionStatus,
            string sortBy,
            bool sortDescending,
            int page,
            int resultsPerPage);

        Task<List<BoricaTransactionVO>> GetBoricaTransactions(DateTime? dateFrom,
            DateTime? dateTo,
            int? transactionStatus,
            string sortBy,
            bool sortDescending,
            int takeCount = int.MaxValue);

        Task<BoricaTransactionVO> GetBoricaTransaction(int transactionId);

        Task<List<PaymentRequestVO>> GetTransactionPaymentRequests(int transactionId);
    }
}
