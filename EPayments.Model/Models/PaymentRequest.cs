using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class PaymentRequest
    {
        public int PaymentRequestId { get; set; }
        public Guid Gid { get; set; }
        public int PaymentRequestXmlId { get; set; }
        public int EserviceClientId { get; set; }

        public bool IsPaymentMultiple { get; set; }

        public string ServiceProviderName { get; set; }
        public string ServiceProviderBank { get; set; }
        public string ServiceProviderBIC { get; set; }
        public string ServiceProviderIBAN { get; set; }

        public string Currency { get; set; }
        public string PaymentTypeCode { get; set; }
        public decimal PaymentAmount { get; set; }
        public string PaymentReason { get; set; }

        public Enums.UinType ApplicantUinTypeId { get; set; }
        public string ApplicantUin { get; set; }
        public string ApplicantName { get; set; }

        public string PaymentReferenceType { get; set; }
        public string PaymentReferenceNumber { get; set; }
        public DateTime PaymentReferenceDate { get; set; }

        public DateTime ExpirationDate { get; set; }
        public string AdditionalInformation { get; set; }
        public string AdministrativeServiceUri { get; set; }
        public string AdministrativeServiceSupplierUri { get; set; }
        public string AdministrativeServiceNotificationURL { get; set; }
        public string AisPaymentId { get; set; }
        public string PaymentRequestIdentifier { get; set; }
        public string PaymentRequestAccessCode { get; set; }
        public bool IsTempRequest { get; set; }
        public bool IsVposAuthorized { get; set; }
        public int? VposResultId { get; set; }
        public string VposAuthorizationId { get; set; }
        public PaymentRequestStatus PaymentRequestStatusId { get; set; }
        public ObligationStatusEnum? ObligationStatusId { get; set; }
        public PaidStatusPaymentMethod? PaidStatusPaymentMethodId { get; set; }
        public string PaidStatusPaymentMethodDescription { get; set; }
        public DateTime PaymentRequestStatusChangeTime { get; set; }
        public DateTime CreateDate { get; set; }
        public int InitiatorId { get; set; }
        public int ObligationTypeId { get; set; }
        public virtual ObligationType ObligationType { get; set; }

        [ForeignKey("InitiatorId")]
        public virtual EserviceClient InitiatorEserviceClient { get; set; }
        public virtual PaymentRequestXml PaymentRequestXml { get; set; }
        public virtual VposResult VposResult { get; set; }

        [ForeignKey("EserviceClientId")]
        public virtual EserviceClient EserviceClient { get; set; }

        public virtual ICollection<PaymentRequestObligationLog> PaymentRequestObligationLogs { get; set; } = new HashSet<PaymentRequestObligationLog>();

        public virtual ICollection<BoricaTransaction> BoricaTransactions { get; set; } = new HashSet<BoricaTransaction>();

        public virtual DistributionRevenuePayment DistributionRevenuePayment { get; set; }

        public string RedirectUrl { get; set; }

        public int? PayOrder { get; set; }
    }

    public class PaymentRequestMap : EntityTypeConfiguration<PaymentRequest>
    {
        public PaymentRequestMap()
        {
            // Primary Key
            this.HasKey(t => t.PaymentRequestId);

            this.Property(t => t.PaymentRequestId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.ServiceProviderName)
                .IsRequired();

            this.Property(t => t.ServiceProviderBank)
                .IsRequired()
                .HasMaxLength(100);

            this.Property(t => t.ServiceProviderBIC)
                .IsRequired()
                .HasMaxLength(50);

            this.Property(t => t.ServiceProviderIBAN)
                .IsRequired()
                .HasMaxLength(50);

            this.Property(t => t.Currency)
                .IsRequired()
                .HasMaxLength(10);

            this.Property(t => t.PaymentTypeCode)
                .HasMaxLength(10);

            this.Property(t => t.PaymentReason)
                .IsRequired()
                .HasMaxLength(200);

            this.Property(t => t.ApplicantUin)
                .IsRequired()
                .HasMaxLength(50);

            this.Property(t => t.ApplicantName)
                .IsRequired();

            this.Property(t => t.PaymentReferenceType)
                .HasMaxLength(50);

            this.Property(t => t.PaymentReferenceNumber)
                .IsRequired()
                .HasMaxLength(50);

            this.Property(t => t.AdministrativeServiceUri)
                .HasMaxLength(100);

            this.Property(t => t.AdministrativeServiceSupplierUri)
                .HasMaxLength(100);

            this.Property(t => t.PaymentRequestIdentifier)
                .IsRequired()
                .HasMaxLength(50);

            this.Property(t => t.PaymentRequestAccessCode)
                .HasMaxLength(10);

            // Table & Column Mappings
            this.ToTable("PaymentRequests");
            this.Property(t => t.PaymentRequestId).HasColumnName("PaymentRequestId");
            this.Property(t => t.Gid).HasColumnName("Gid");
            this.Property(t => t.PaymentRequestXmlId).HasColumnName("PaymentRequestXmlId");
            this.Property(t => t.EserviceClientId).HasColumnName("EserviceClientId");
            this.Property(t => t.InitiatorId).HasColumnName("InitiatorId");
            this.Property(t => t.IsPaymentMultiple).HasColumnName("IsPaymentMultiple");
            this.Property(t => t.Currency).HasColumnName("Currency");
            this.Property(t => t.PaymentTypeCode).HasColumnName("PaymentTypeCode");
            this.Property(t => t.PaymentAmount).HasColumnName("PaymentAmount");
            this.Property(t => t.PaymentReason).HasColumnName("PaymentReason");
            this.Property(t => t.ServiceProviderName).HasColumnName("ServiceProviderName");
            this.Property(t => t.ServiceProviderBank).HasColumnName("ServiceProviderBank");
            this.Property(t => t.ServiceProviderBIC).HasColumnName("ServiceProviderBIC");
            this.Property(t => t.ServiceProviderIBAN).HasColumnName("ServiceProviderIBAN");
            this.Property(t => t.ApplicantName).HasColumnName("ApplicantName");
            this.Property(t => t.ApplicantUinTypeId).HasColumnName("ApplicantUinTypeId");
            this.Property(t => t.ApplicantUin).HasColumnName("ApplicantUin");
            this.Property(t => t.PaymentReferenceType).HasColumnName("PaymentReferenceType");
            this.Property(t => t.PaymentReferenceNumber).HasColumnName("PaymentReferenceNumber");
            this.Property(t => t.PaymentReferenceDate).HasColumnName("PaymentReferenceDate");
            this.Property(t => t.ExpirationDate).HasColumnName("ExpirationDate");
            this.Property(t => t.AdditionalInformation).HasColumnName("AdditionalInformation");
            this.Property(t => t.AdministrativeServiceUri).HasColumnName("AdministrativeServiceUri");
            this.Property(t => t.AdministrativeServiceSupplierUri).HasColumnName("AdministrativeServiceSupplierUri");
            this.Property(t => t.AdministrativeServiceNotificationURL).HasColumnName("AdministrativeServiceNotificationURL");
            this.Property(t => t.AisPaymentId).HasColumnName("AisPaymentId");
            this.Property(t => t.PaymentRequestIdentifier).HasColumnName("PaymentRequestIdentifier");
            this.Property(t => t.PaymentRequestAccessCode).HasColumnName("PaymentRequestAccessCode");
            this.Property(t => t.IsTempRequest).HasColumnName("IsTempRequest");
            this.Property(t => t.IsVposAuthorized).HasColumnName("IsVposAuthorized");
            this.Property(t => t.VposResultId).HasColumnName("VposResultId");
            this.Property(t => t.VposAuthorizationId).HasColumnName("VposAuthorizationId");
            this.Property(t => t.PaymentRequestStatusId).HasColumnName("PaymentRequestStatusId");
            this.Property(t => t.PaidStatusPaymentMethodId).HasColumnName("PaidStatusPaymentMethodId");
            this.Property(t => t.PaidStatusPaymentMethodDescription).HasColumnName("PaidStatusPaymentMethodDescription");
            this.Property(t => t.PaymentRequestStatusChangeTime).HasColumnName("PaymentRequestStatusChangeTime");
            this.Property(t => t.ObligationTypeId).HasColumnName("ObligationTypeId");
            this.Property(t => t.CreateDate).HasColumnName("CreateDate");
            this.Property(t => t.ObligationStatusId).HasColumnName("ObligationStatusId");
        }
    }
}
