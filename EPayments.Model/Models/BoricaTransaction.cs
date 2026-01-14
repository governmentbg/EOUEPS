using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;
using EPayments.Model.DataObjects;
using Newtonsoft.Json;

namespace EPayments.Model.Models
{
    public class BoricaTransaction
    {
        public int BoricaTransactionId { get; set; }

        public Guid Gid { get; set; }

        public decimal Amount { get; set; }

        public string Order { get; set; }

        public string Description { get; set; }

        public decimal? Fee { get; set; }

        public decimal? Commission { get; set; }

        public DateTime TransactionDate { get; set; }// date return from borika

        public string Terminal { get; set; }

        public string Action { get; set; }

        public string Rc { get; set; }

        public string Approval { get; set; }

        public string Rrn { get; set; }

        public string IntRef { get; set; }

        public string StatusMessage { get; set; }

        public string Card { get; set; }

        public string Eci { get; set; }

        public string ParesStatus { get; set; }

        public DateTime? SettlementDate { get; set; }

        public string AuthorizationCode { get; set; } 

        public DateTime? DistributionDate { get; set; } 

        public string ProductCategory { get; set; }

        public string AreaOfIssue { get; set; } 

        public int TransactionStatusId { get; set; }

        public int? JobCheckResultFailedAttempts { get; set; }

        public string JobCheckResultFailedAttemptsErrors { get; set; }

        public DateTime? JobLastCheckResultDate { get; set; }

        public string JobLastCheckResultTransactionInfo { get; set; }

        public virtual BoricaTransactionStatus BoricaTransactionStatus { get; set; }

        public virtual ICollection<PaymentRequest> PaymentRequests { get; set; } = new HashSet<PaymentRequest>();

        public virtual ICollection<DistributionRevenuePayment> DistributionRevenuePayments { get; set; } = new HashSet<DistributionRevenuePayment>(); 

        public string BoricaTransactionSettlement { get; set; }
    }

    public class BoricaTransactionMap : EntityTypeConfiguration<BoricaTransaction>
    {
        public BoricaTransactionMap()
        {
            this.HasKey(t => t.BoricaTransactionId);

            this.Property(t => t.BoricaTransactionId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.Order)
                .IsRequired()
                .HasMaxLength(16);

            this.Property(t => t.Description)
                .HasMaxLength(50);

            this.Property(t => t.Terminal)
                .HasMaxLength(8);

            this.Property(t => t.Action)
                .HasMaxLength(4);

            this.Property(t => t.Rc)
                .HasMaxLength(4);

            this.Property(t => t.Approval)
                .HasMaxLength(6);

            this.Property(t => t.Rrn)
                .HasMaxLength(12);

            this.Property(t => t.IntRef)
                .HasMaxLength(32);

            this.Property(t => t.StatusMessage)
                .HasMaxLength(255);

            this.Property(t => t.Card)
                .HasMaxLength(20);

            this.Property(t => t.Eci)
                .HasMaxLength(2);

            this.Property(t => t.ParesStatus)
                .HasMaxLength(20);

            this.Property(t => t.AuthorizationCode)
                .HasMaxLength(6);

            this.Property(t => t.ProductCategory)
                .HasMaxLength(100);

            this.Property(t => t.AreaOfIssue)
                .HasMaxLength(100);

            this.Property(t => t.JobCheckResultFailedAttemptsErrors)
                .HasMaxLength(100);

            this.Property(t => t.JobLastCheckResultTransactionInfo)
                .HasMaxLength(100);

            this.Property(t => t.BoricaTransactionSettlement);

            this.HasMany(t => t.PaymentRequests)
                .WithMany(pr => pr.BoricaTransactions)
                .Map(conf => conf.ToTable("BoricaTransactionPaymentRequest").MapLeftKey("BoricaTransactionId").MapRightKey("PaymentRequestId"));

            this.HasRequired(t => t.BoricaTransactionStatus)
                .WithMany(bts => bts.BoricaTransactions)
                .HasForeignKey(t => t.TransactionStatusId);

            // Table & Column Mappings
            this.ToTable("BoricaTransactions");
            this.Property(t => t.BoricaTransactionId).HasColumnName("BoricaTransactionId");
            this.Property(t => t.Gid).HasColumnName("Gid");
            this.Property(t => t.Amount).HasColumnName("Amount");
            this.Property(t => t.Order).HasColumnName("Order");
            this.Property(t => t.Description).HasColumnName("Description");
            this.Property(t => t.Terminal).HasColumnName("Terminal");
            this.Property(t => t.Action).HasColumnName("Action");
            this.Property(t => t.Rc).HasColumnName("Rc");
            this.Property(t => t.Approval).HasColumnName("Approval");
            this.Property(t => t.Rrn).HasColumnName("Rrn");
            this.Property(t => t.IntRef).HasColumnName("IntRef");
            this.Property(t => t.StatusMessage).HasColumnName("StatusMessage");
            this.Property(t => t.Card).HasColumnName("Card");
            this.Property(t => t.Eci).HasColumnName("Eci");
            this.Property(t => t.ParesStatus).HasColumnName("ParesStatus");
            this.Property(t => t.Fee).HasColumnName("Fee").HasPrecision(18, 4);
            this.Property(t => t.Commission).HasColumnName("Commission").HasPrecision(18, 4);
            this.Property(t => t.TransactionDate).HasColumnName("TransactionDate");
            this.Property(t => t.TransactionStatusId).HasColumnName("TransactionStatusId");
            this.Property(t => t.SettlementDate).HasColumnName("SettlementDate");
            this.Property(t => t.AuthorizationCode).HasColumnName("AuthorizationCode");
            this.Property(t => t.ProductCategory).HasColumnName("ProductCategory");
            this.Property(t => t.AreaOfIssue).HasColumnName("AreaOfIssue");
            this.Property(t => t.DistributionDate).HasColumnName("DistributionDate");
            this.Property(t => t.JobLastCheckResultDate).HasColumnName("JobLastCheckResultDate");
            this.Property(t => t.JobLastCheckResultTransactionInfo).HasColumnName("JobLastCheckResultTransactionInfo");
            this.Property(t => t.JobCheckResultFailedAttemptsErrors).HasColumnName("JobCheckResultFailedAttemptsErrors");
            this.Property(t => t.JobCheckResultFailedAttempts).HasColumnName("JobCheckResultFailedAttempts");
            this.Property(t => t.BoricaTransactionSettlement).HasColumnName("BoricaTransactionSettlement");
        }
    }
}
