using EPayments.Common.Helpers;
using EPayments.Model.Enums;
using System;
using System.Collections.Generic;

namespace EPayments.Data.ViewObjects.Web
{
    public class PendingRequestVO
    {
        public Guid Gid { get; set; }
        public DateTime CreateDate { get; set; }
        public string ApplicantName { get; set; }
        public string PaymentRequestIdentifier { get; set; }
        public DateTime ExpirationDate { get; set; }
        public string ServiceProviderName { get; set; }
        public string PaymentReason { get; set; }
        public decimal PaymentAmount { get; set; }
        public decimal PaymentAmountInBgn => CurrencyHelper.GetBgnValue(PaymentAmount, CreateDate);
        public decimal PaymentAmountInEuro => CurrencyHelper.GetEuroValue(PaymentAmount, CreateDate);
        public bool IsEpayVposEnabled { get; set; }
        public bool IsCvposEnabled { get; set; }
        public Vpos? Vpos { get; set; }
        public int? ObligationStatusId { get; set; }
        public bool SelectedForPayment { get; set; }

        public string ObligationType { get; set; }

        public int ObligationTypeId { get; set; }

        public string InitiatorName { get; set; }

        public bool IsInProgress { get; set; }

        public bool IsDisabled { get; set; }

        public int? PayOrder { get; set; }

        public int EserviceClientId { get; set; }
        public int? MaxPayOrder { get; set; }

        public string AdditionalInfo { get; set; }

        public int AlgorithmId { get; set; }
    }
}
