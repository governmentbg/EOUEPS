using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class TransactionRecord
    {
        public int TransactionRecordId { get; set; }
        public int TransactionFileId { get; set; }
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
        public string InfoDocumentNumberDate { get; set; }
        public DateTime? InfoPaymentPeriodBegining { get; set; }
        public DateTime? InfoPaymentPeriodEnd { get; set; }
        public string InfoDebtorBulstat { get; set; }
        public string InfoDebtorEgn { get; set; }
        public string InfoDebtorLnch { get; set; }
        public string InfoDebtorName { get; set; }
        public string InfoDebtorBulstatEgnLnch { get; set; }
        public string InfoDebtorBulstatEgnLnchName { get; set; }
        public string InfoAC1AuthorizationCode { get; set; }
        public string InfoAC1BankCardInfo { get; set; }
        public string InfoPaymentDetailsRaw { get; set; }
        public string InfoPaymentReason { get; set; }
        public string InfoSenderBic { get; set; }
        public string InfoSenderIban { get; set; }
        public string InfoSenderName { get; set; }
        public string InfoSenderIbanName { get; set; }
        public TransactionRecordPaymentMethod TransactionRecordPaymentMethodId { get; set; }
        public TransactionRecordRefStatus TransactionRecordRefStatusId { get; set; }
        public int? PaymentRequestId { get; set; }

        public virtual TransactionFile TransactionFile { get; set; }
        public virtual PaymentRequest PaymentRequest { get; set; }
    }

    public class TransactionRecordMap : EntityTypeConfiguration<TransactionRecord>
    {
        public TransactionRecordMap()
        {
            // Primary Key
            this.HasKey(t => t.TransactionRecordId);

            // Properties
            this.Property(t => t.TransactionReferenceId)
                .HasMaxLength(100);

            this.Property(t => t.InfoSystemTransactionType)
                .HasMaxLength(100);

            this.Property(t => t.InfoPaymentType)
                .HasMaxLength(100);

            this.Property(t => t.InfoDocumentType)
                .HasMaxLength(100);

            this.Property(t => t.InfoDocumentNumber)
                .HasMaxLength(100);

            this.Property(t => t.InfoDocumentNumberDate)
                .HasMaxLength(150);

            this.Property(t => t.InfoDebtorBulstat)
                .HasMaxLength(100);

            this.Property(t => t.InfoDebtorEgn)
                .HasMaxLength(100);

            this.Property(t => t.InfoDebtorLnch)
                .HasMaxLength(100);

            this.Property(t => t.InfoDebtorBulstatEgnLnch)
                .HasMaxLength(300);

            this.Property(t => t.InfoDebtorBulstatEgnLnchName)
                .HasMaxLength(600);

            this.Property(t => t.InfoDebtorName)
                .HasMaxLength(300);

            this.Property(t => t.InfoAC1AuthorizationCode)
                .HasMaxLength(100);

            this.Property(t => t.InfoAC1BankCardInfo)
                .HasMaxLength(100);

            this.Property(t => t.InfoSenderBic)
                .HasMaxLength(100);

            this.Property(t => t.InfoSenderIban)
                .HasMaxLength(100);

            // Table & Column Mappings
            this.ToTable("TransactionRecords");
            this.Property(t => t.TransactionRecordId).HasColumnName("TransactionRecordId");
            this.Property(t => t.TransactionFileId).HasColumnName("TransactionFileId");
            this.Property(t => t.TransactionDate).HasColumnName("TransactionDate");
            this.Property(t => t.TransactionAccountingDate).HasColumnName("TransactionAccountingDate");
            this.Property(t => t.TransactionAmount).HasColumnName("TransactionAmount");
            this.Property(t => t.TransactionReferenceId).HasColumnName("TransactionReferenceId");
            this.Property(t => t.InfoSystemTransactionType).HasColumnName("InfoSystemTransactionType");
            this.Property(t => t.InfoSystemTransactionDesc).HasColumnName("InfoSystemTransactionDesc");
            this.Property(t => t.InfoPaymentType).HasColumnName("InfoPaymentType");
            this.Property(t => t.InfoDocumentType).HasColumnName("InfoDocumentType");
            this.Property(t => t.InfoDocumentNumber).HasColumnName("InfoDocumentNumber");
            this.Property(t => t.InfoDocumentDate).HasColumnName("InfoDocumentDate");
            this.Property(t => t.InfoDocumentNumberDate).HasColumnName("InfoDocumentNumberDate");
            this.Property(t => t.InfoPaymentPeriodBegining).HasColumnName("InfoPaymentPeriodBegining");
            this.Property(t => t.InfoPaymentPeriodEnd).HasColumnName("InfoPaymentPeriodEnd");
            this.Property(t => t.InfoDebtorBulstat).HasColumnName("InfoDebtorBulstat");
            this.Property(t => t.InfoDebtorEgn).HasColumnName("InfoDebtorEgn");
            this.Property(t => t.InfoDebtorLnch).HasColumnName("InfoDebtorLnch");
            this.Property(t => t.InfoDebtorName).HasColumnName("InfoDebtorName");
            this.Property(t => t.InfoDebtorBulstatEgnLnch).HasColumnName("InfoDebtorBulstatEgnLnch");
            this.Property(t => t.InfoDebtorBulstatEgnLnchName).HasColumnName("InfoDebtorBulstatEgnLnchName");
            this.Property(t => t.InfoAC1AuthorizationCode).HasColumnName("InfoAC1AuthorizationCode");
            this.Property(t => t.InfoAC1BankCardInfo).HasColumnName("InfoAC1BankCardInfo");
            this.Property(t => t.InfoPaymentDetailsRaw).HasColumnName("InfoPaymentDetailsRaw");
            this.Property(t => t.InfoPaymentReason).HasColumnName("InfoPaymentReason");
            this.Property(t => t.InfoSenderBic).HasColumnName("InfoSenderBic");
            this.Property(t => t.InfoSenderIban).HasColumnName("InfoSenderIban");
            this.Property(t => t.InfoSenderName).HasColumnName("InfoSenderName");
            this.Property(t => t.InfoSenderIbanName).HasColumnName("InfoSenderIbanName");
            this.Property(t => t.TransactionRecordPaymentMethodId).HasColumnName("TransactionRecordPaymentMethodId");
            this.Property(t => t.TransactionRecordRefStatusId).HasColumnName("TransactionRecordRefStatusId");
            this.Property(t => t.PaymentRequestId).HasColumnName("PaymentRequestId");
        }
    }
}
