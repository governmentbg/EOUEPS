using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;

namespace EPayments.Web.Models.EserviceAdmin
{
    public class TransactionDetailsVM
    {
        public string BankStatementIban { get; set; }
        public DateTime? BankStatementDate { get; set; }
        public string BankStatementNumber { get; set; }

        public DateTime? TransactionDate { get; set; }
        public DateTime? TransactionAccountingDate { get; set; }
        public decimal? TransactionAmount { get; set; }
        public string TransactionReferenceId { get; set; }
        public string InfoSystemTransactionType { get; set; }
        public string InfoSystemTransactionDesc { get; set; }
        public string InfoPaymentType { get; set; }
        public string InfoDocumentType { get; set; }
        public string InfoDocumentNumber { get; set; }
        public DateTime? InfoDocumentDate { get; set; }
        public DateTime? InfoPaymentPeriodBegining { get; set; }
        public DateTime? InfoPaymentPeriodEnd { get; set; }
        public string InfoDebtorBulstat { get; set; }
        public string InfoDebtorEgn { get; set; }
        public string InfoDebtorLnch { get; set; }
        public string InfoDebtorName { get; set; }
        public string InfoAC1AuthorizationCode { get; set; }
        public string InfoAC1BankCardInfo { get; set; }
        public string InfoPaymentReason { get; set; }
        public string InfoSenderBic { get; set; }
        public string InfoSenderIban { get; set; }
        public string InfoSenderName { get; set; }
        public TransactionRecordRefStatus? TransactionRecordRefStatusId { get; set; }

        public void SetFields(TransactionRecord transactionRecord)
        {
            this.BankStatementIban = transactionRecord.TransactionFile.BankStatementIban;
            this.BankStatementDate = transactionRecord.TransactionFile.BankStatementDate;
            this.BankStatementNumber = transactionRecord.TransactionFile.BankStatementNumber;

            this.TransactionDate = transactionRecord.TransactionDate;
            this.TransactionAccountingDate = transactionRecord.TransactionAccountingDate;
            this.TransactionAmount = transactionRecord.TransactionAmount;
            this.TransactionReferenceId = transactionRecord.TransactionReferenceId;
            this.InfoSystemTransactionType = transactionRecord.InfoSystemTransactionType;
            this.InfoSystemTransactionDesc = transactionRecord.InfoSystemTransactionDesc;
            this.InfoPaymentType = transactionRecord.InfoPaymentType;
            this.InfoDocumentType = transactionRecord.InfoDocumentType;
            this.InfoDocumentNumber = transactionRecord.InfoDocumentNumber;
            this.InfoDocumentDate = transactionRecord.InfoDocumentDate;
            this.InfoPaymentPeriodBegining = transactionRecord.InfoPaymentPeriodBegining;
            this.InfoPaymentPeriodEnd = transactionRecord.InfoPaymentPeriodEnd;
            this.InfoDebtorBulstat = transactionRecord.InfoDebtorBulstat;
            this.InfoDebtorEgn = transactionRecord.InfoDebtorEgn;
            this.InfoDebtorLnch = transactionRecord.InfoDebtorLnch;
            this.InfoDebtorName = transactionRecord.InfoDebtorName;
            this.InfoAC1AuthorizationCode = transactionRecord.InfoAC1AuthorizationCode;
            this.InfoAC1BankCardInfo = transactionRecord.InfoAC1BankCardInfo;
            this.InfoPaymentReason = transactionRecord.InfoPaymentReason;
            this.InfoSenderBic = transactionRecord.InfoSenderBic;
            this.InfoSenderIban = transactionRecord.InfoSenderIban;
            this.InfoSenderName = transactionRecord.InfoSenderName;
            this.TransactionRecordRefStatusId = transactionRecord.TransactionRecordRefStatusId;
        }
    }
}