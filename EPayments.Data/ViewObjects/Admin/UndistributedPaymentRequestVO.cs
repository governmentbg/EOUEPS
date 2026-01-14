using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Linq.Expressions;

namespace EPayments.Data.ViewObjects.Admin
{
    public class UndistributedPaymentRequestVO
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

        public int? DistributionRevenuePaymentId { get; set; } 

        public static Expression<Func<PaymentRequest, UndistributedPaymentRequestVO>> Map = (pr) =>
            new UndistributedPaymentRequestVO()
            {
                Gid = pr.Gid,
                CreateDate = pr.CreateDate,
                ApplicantName = pr.ApplicantName,
                PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                PaymentReferenceNumber = pr.PaymentReferenceNumber,
                ExpirationDate = pr.ExpirationDate,
                ServiceProviderName = pr.ServiceProviderName,
                DistributionRevenuePaymentId = pr.DistributionRevenuePayment != null ? pr.DistributionRevenuePayment.DistributionRevenueId : default(int?),
                PaymentAmount = pr.PaymentAmount,
                PaymentRequestStatusId = pr.PaymentRequestStatusId,
                ObligationStatusId = pr.ObligationStatusId
            };
    }
}
