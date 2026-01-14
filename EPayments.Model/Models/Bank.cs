using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class Bank
    {
        public int BankId { get; set; }
        public string Name { get; set; }
        public string BIC { get; set; }
        public int? CertificateId { get; set; }
    }

    public class BankMap : EntityTypeConfiguration<Bank>
    {
        public BankMap()
        {
            // Primary Key
            this.HasKey(t => t.BankId);

            this.Property(t => t.BankId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.Name)
                .IsRequired()
                .HasMaxLength(200);

            this.Property(t => t.BIC)
                .IsRequired()
                .HasMaxLength(50);

            // Table & Column Mappings
            this.ToTable("Banks");
            this.Property(t => t.BankId).HasColumnName("BankId");
            this.Property(t => t.Name).HasColumnName("Name");
            this.Property(t => t.BIC).HasColumnName("BIC");
            this.Property(t => t.CertificateId).HasColumnName("CertificateId");
        }
    }
}
