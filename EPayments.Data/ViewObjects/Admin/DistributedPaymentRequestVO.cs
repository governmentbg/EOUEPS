using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Linq.Expressions;

namespace EPayments.Data.ViewObjects.Admin
{
    public class DistributedPaymentRequestVO
    {
        public string PaymentRequestIdentifier { get; set; }
        
        public string PaymentReason { get; set; }
        
        public decimal PaymentAmount { get; set; } 

        public DateTime CreateDate { get; set; }

        public string EServiceClientName { get; set; }

        public int TargetEServiceClientId { get; set; }

        public string TargetEServiceClientName { get; set; }

        public string ApplicantName { get; set; }

        public PaymentRequestStatus PaymentRequestStatus { get; set; }

        public ObligationStatusEnum? ObligationStatus { get; set; }

        public static Expression<Func<DistributionRevenuePayment, DistributedPaymentRequestVO>> Map { get; } = (drp) =>
        new DistributedPaymentRequestVO()
        {
            PaymentRequestIdentifier = drp.PaymentRequest.PaymentRequestIdentifier,
            PaymentReason = drp.PaymentRequest.PaymentReason,
            PaymentAmount = drp.PaymentRequest.PaymentAmount,
            CreateDate = drp.PaymentRequest.CreateDate,
            EServiceClientName = drp.PaymentRequest.EserviceClient.AisName,
            TargetEServiceClientId = drp.EserviceClient.EserviceClientId,
            TargetEServiceClientName = drp.EserviceClient.AisName,
            ApplicantName = drp.PaymentRequest.ApplicantName,
            PaymentRequestStatus = drp.PaymentRequest.PaymentRequestStatusId,
            ObligationStatus = drp.PaymentRequest.ObligationStatusId
        };
    }
}
