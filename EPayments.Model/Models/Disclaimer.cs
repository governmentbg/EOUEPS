using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class Disclaimer
    {
        public int DisclaimerId { get; set; }
        public int UserId { get; set; }
        public int CertificateId { get; set; }
        public int DisclaimerTemplateId { get; set; }
        public string SignedXml { get; set; }
        public DateTime SignDate { get; set; }
    }

    public class DisclaimerMap : EntityTypeConfiguration<Disclaimer>
    {
        public DisclaimerMap()
        {
            // Primary Key
            this.HasKey(t => t.DisclaimerId);

            this.Property(t => t.DisclaimerId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.SignedXml)
                .IsRequired();

            // Table & Column Mappings
            this.ToTable("Disclaimers");
            this.Property(t => t.DisclaimerId).HasColumnName("DisclaimerId");
            this.Property(t => t.UserId).HasColumnName("UserId");
            this.Property(t => t.CertificateId).HasColumnName("CertificateId");
            this.Property(t => t.DisclaimerTemplateId).HasColumnName("DisclaimerTemplateId");
            this.Property(t => t.SignedXml).HasColumnName("SignedXml");
            this.Property(t => t.SignDate).HasColumnName("SignDate");
        }
    }
}
