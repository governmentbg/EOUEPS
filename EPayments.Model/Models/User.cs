using EPayments.Common.Helpers;
using EPayments.Model.DataObjects.EmailTemplateContext;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class User
    {
        public int UserId { get; set; }
        public string Egn { get; set; }
        public string Email { get; set; }
        public bool RequestNotifications { get; set; }
        public bool StatusNotifications { get; set; }
        public bool StatusObligationNotifications { get; set; }
        public bool AccessCodeNotifications { get; set; }
        public int? LastCertificateId { get; set; }
        public int? LastDisclaimerTemplateId { get; set; }

        public Email CreateStatusObligationNotificationEmail(PaymentRequest paymentRequest)
        {
            if (paymentRequest == null || StatusObligationNotifications == false || string.IsNullOrEmpty(Email))
            {
                return null;
            }
            
            StatusChangedObligationContextDO contextOblDO = new StatusChangedObligationContextDO(
                paymentRequest.PaymentRequestIdentifier,
                paymentRequest.ServiceProviderName,
                paymentRequest.PaymentReason,
                paymentRequest.PaymentAmount,
                paymentRequest.ObligationStatusId.GetDescription(),
                paymentRequest.CreateDate);

            return new Email(contextOblDO, this.Email);
        }

        public Email CreateStatusNotificationEmail(PaymentRequest paymentRequest)
        {
            if (paymentRequest == null || StatusNotifications == false || string.IsNullOrEmpty(Email))
            {
                return null;
            }

            StatusChangedPaymentRequestContextDO contextOblDO = new StatusChangedPaymentRequestContextDO(
                paymentRequest.PaymentRequestIdentifier,
                paymentRequest.ServiceProviderName,
                paymentRequest.PaymentReason,
                paymentRequest.PaymentAmount,
                paymentRequest.PaymentRequestStatusId.GetDescription(),
                paymentRequest.CreateDate);

            return new Email(contextOblDO, this.Email);
        }
    }

    public class UserMap : EntityTypeConfiguration<User>
    {
        public UserMap()
        {
            // Primary Key
            this.HasKey(t => t.UserId);

            this.Property(t => t.UserId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.Egn)
                .IsRequired()
                .HasMaxLength(14);

            this.Property(t => t.Email)
                .HasMaxLength(100);

            // Table & Column Mappings
            this.ToTable("Users");
            this.Property(t => t.UserId).HasColumnName("UserId");
            this.Property(t => t.Egn).HasColumnName("Egn");
            this.Property(t => t.Email).HasColumnName("Email");
            this.Property(t => t.RequestNotifications).HasColumnName("RequestNotifications");
            this.Property(t => t.StatusNotifications).HasColumnName("StatusNotifications");
            this.Property(t => t.StatusObligationNotifications).HasColumnName("StatusObligationNotifications");
            this.Property(t => t.AccessCodeNotifications).HasColumnName("AccessCodeNotifications");
            this.Property(t => t.LastCertificateId).HasColumnName("LastCertificateId");
            this.Property(t => t.LastDisclaimerTemplateId).HasColumnName("LastDisclaimerTemplateId");
        }
    }
}
