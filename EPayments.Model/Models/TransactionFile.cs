using System;
using System.Collections.Generic;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class TransactionFile
    {
        public int TransactionFileId { get; set; }
        public int EserviceBankAccountId { get; set; }
        public string FileName { get; set; }
        public string BankStatementIban { get; set; }
        public DateTime? BankStatementDate { get; set; }
        public string BankStatementNumber { get; set; }
        public DateTime CreateDate { get; set; }

        public virtual ICollection<TransactionRecord> TransactionRecords { get; set; }

        public TransactionFile()
        {
            this.TransactionRecords = new List<TransactionRecord>();
        }
    }

    public class TransactionFileMap : EntityTypeConfiguration<TransactionFile>
    {
        public TransactionFileMap()
        {
            // Primary Key
            this.HasKey(t => t.TransactionFileId);

            // Properties
            this.Property(t => t.BankStatementIban)
                .IsRequired()
                .HasMaxLength(100);

            this.Property(t => t.BankStatementNumber)
                .HasMaxLength(100);

            // Table & Column Mappings
            this.ToTable("TransactionFiles");
            this.Property(t => t.TransactionFileId).HasColumnName("TransactionFileId");
            this.Property(t => t.EserviceBankAccountId).HasColumnName("EserviceBankAccountId");
            this.Property(t => t.FileName).HasColumnName("FileName");
            this.Property(t => t.BankStatementIban).HasColumnName("BankStatementIban");
            this.Property(t => t.BankStatementDate).HasColumnName("BankStatementDate");
            this.Property(t => t.BankStatementNumber).HasColumnName("BankStatementNumber");
            this.Property(t => t.CreateDate).HasColumnName("CreateDate");
        }
    }
}
