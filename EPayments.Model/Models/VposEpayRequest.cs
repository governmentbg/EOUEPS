using EPayments.Model.Enums;
using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class VposEpayRequest
    {
        public int VposEpayRequestId { get; set; }
        public int PaymentRequestId { get; set; }
        public int? VposRedirectId { get; set; }
        public string ClientIpAddress { get; set; }
        public string TransactionInvoiceNo { get; set; }
        public bool IsVposPostCallbackReceived { get; set; }
        public DateTime? VposPostCallbackReceiveDate { get; set; }
        public bool? VposPostCallbackSuccessConfirmation { get; set; }
        public string VposPostCallbackAisRedirectedStatus { get; set; }
        public string TransactionResult { get; set; }
	    public DateTime? TransactionResultReceiveDate { get; set; }
        public bool? IsPaymentSuccessful { get; set; }
        public DateTime CreateDate { get; set; }

        public virtual PaymentRequest PaymentRequest { get; set; }
    }

    public class VposEpayRequestMap : EntityTypeConfiguration<VposEpayRequest>
    {
        public VposEpayRequestMap()
        {
            // Primary Key
            this.HasKey(t => t.VposEpayRequestId);

            this.Property(t => t.VposEpayRequestId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.TransactionInvoiceNo)
                .IsRequired();

            // Table & Column Mappings
            this.ToTable("VposEpayRequests");

            this.Property(t => t.VposEpayRequestId).HasColumnName("VposEpayRequestId");
            this.Property(t => t.PaymentRequestId).HasColumnName("PaymentRequestId");
            this.Property(t => t.TransactionInvoiceNo).HasColumnName("TransactionInvoiceNo");
            this.Property(t => t.ClientIpAddress).HasColumnName("ClientIpAddress");
            this.Property(t => t.VposRedirectId).HasColumnName("VposRedirectId");
            this.Property(t => t.IsVposPostCallbackReceived).HasColumnName("IsVposPostCallbackReceived");
            this.Property(t => t.VposPostCallbackReceiveDate).HasColumnName("VposPostCallbackReceiveDate");
            this.Property(t => t.VposPostCallbackSuccessConfirmation).HasColumnName("VposPostCallbackSuccessConfirmation");
            this.Property(t => t.VposPostCallbackAisRedirectedStatus).HasColumnName("VposPostCallbackAisRedirectedStatus");
            this.Property(t => t.TransactionResult).HasColumnName("TransactionResult");
            this.Property(t => t.TransactionResultReceiveDate).HasColumnName("TransactionResultReceiveDate");
            this.Property(t => t.IsPaymentSuccessful).HasColumnName("IsPaymentSuccessful");
            this.Property(t => t.CreateDate).HasColumnName("CreateDate");
        }
    }
}
