using EPayments.Model.Enums;
using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class VposDskEcommRequest
    {
        public int VposFiBankRequestId { get; set; }
        public int PaymentRequestId { get; set; }
        public int? VposRedirectId { get; set; }
        public string ClientIpAddress { get; set; }
        public string TransactionId { get; set; }
        public bool IsVposPostCallbackReceived { get; set; }
        public DateTime? VposPostCallbackReceiveDate { get; set; }
        public string TransactionResult { get; set; }
	    public DateTime? TransactionResultReceiveDate { get; set; }
        public bool? IsPaymentSuccessful { get; set; }
        public VposRequestResultStatus ResultStatus { get; set; }
        public int JobCheckResultFailedAttempts { get; set; }
        public string JobCheckResultFailedAttemptsErrors { get; set; }
        public DateTime? JobLastCheckResultDate { get; set; }
        public string JobLastCheckResultTransactionInfo { get; set; }
        public DateTime CreateDate { get; set; }

        public virtual PaymentRequest PaymentRequest { get; set; }
    }

    public class VposDskEcommRequestMap : EntityTypeConfiguration<VposDskEcommRequest>
    {
        public VposDskEcommRequestMap()
        {
            // Primary Key
            this.HasKey(t => t.VposFiBankRequestId);

            this.Property(t => t.VposFiBankRequestId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.TransactionId)
                .IsRequired();

            // Table & Column Mappings
            this.ToTable("VposDskEcommRequests");

            this.Property(t => t.VposFiBankRequestId).HasColumnName("VposDskEcommRequestId");
            this.Property(t => t.PaymentRequestId).HasColumnName("PaymentRequestId");
            this.Property(t => t.TransactionId).HasColumnName("TransactionId");
            this.Property(t => t.ClientIpAddress).HasColumnName("ClientIpAddress");
            this.Property(t => t.VposRedirectId).HasColumnName("VposRedirectId");
            this.Property(t => t.IsVposPostCallbackReceived).HasColumnName("IsVposPostCallbackReceived");
            this.Property(t => t.VposPostCallbackReceiveDate).HasColumnName("VposPostCallbackReceiveDate");
            this.Property(t => t.TransactionResult).HasColumnName("TransactionResult");
            this.Property(t => t.TransactionResultReceiveDate).HasColumnName("TransactionResultReceiveDate");
            this.Property(t => t.IsPaymentSuccessful).HasColumnName("IsPaymentSuccessful");
            this.Property(t => t.ResultStatus).HasColumnName("ResultStatus");
            this.Property(t => t.JobCheckResultFailedAttempts).HasColumnName("JobCheckResultFailedAttempts");
            this.Property(t => t.JobCheckResultFailedAttemptsErrors).HasColumnName("JobCheckResultFailedAttemptsErrors");
            this.Property(t => t.JobLastCheckResultDate).HasColumnName("JobLastCheckResultDate");
            this.Property(t => t.JobLastCheckResultTransactionInfo).HasColumnName("JobLastCheckResultTransactionInfo");
            this.Property(t => t.CreateDate).HasColumnName("CreateDate");
        }
    }
}
