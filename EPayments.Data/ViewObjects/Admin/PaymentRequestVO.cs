using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Linq.Expressions;

namespace EPayments.Data.ViewObjects.Admin
{
    public class PaymentRequestVO
    {
        public Guid Gid { get; set; }

        public DateTime CreateDate { get; set; }

        public string ApplicantName { get; set; }

        public string PaymentRequestIdentifier { get; set; }

        public string PaymentReferenceNumber { get; set; } 

        public DateTime ExpirationDate { get; set; }

        public string ServiceProviderName { get; set; }

        public string PaymentReason { get; set; }

        public decimal PaymentAmount { get; set; }

        public PaymentRequestStatus PaymentRequestStatusId { get; set; }

        public ObligationStatusEnum? ObligationStatusId { get; set; }

        public static Expression<Func<PaymentRequest, PaymentRequestVO>> Map = (pr) =>
            new PaymentRequestVO()
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
                ObligationStatusId = pr.ObligationStatusId
            };
    }
}
