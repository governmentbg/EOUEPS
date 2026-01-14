using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.ViewObjects.Web
{
    [Serializable()]
    public class TransactionRecordVO
    {
        public int TransactionRecordId { get; set; }
        public DateTime? TransactionAccountingDate { get; set; }
        public decimal? TransactionAmount { get; set; }
        public string InfoDocumentNumber { get; set; }
        public DateTime? InfoDocumentDate { get; set; }
        public string InfoDocumentNumberDate { get; set; }
        public string InfoSenderIban { get; set; }
        public string InfoSenderName { get; set; }
        public string InfoSenderIbanName { get; set; }
        public string InfoDebtorName { get; set; }
        public string InfoDebtorBulstatEgnLnch { get; set; }
        public string InfoDebtorBulstatEgnLnchName { get; set; }
        public string InfoPaymentReason { get; set; }
        public TransactionRecordPaymentMethod TransactionRecordPaymentMethodId { get; set; }
        public TransactionRecordRefStatus TransactionRecordRefStatusId { get; set; }
        public Guid? PaymentRequestGid { get; set; }
        public DateTime? TransactionDate { get; set; }
    }
}
