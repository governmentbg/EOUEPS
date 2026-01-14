using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class Certificate
    {
        public int CertificateId { get; set; }
        public byte[] CertificateFile { get; set; }
        public string CertificateSubject { get; set; }
        public string CertificateIssuer { get; set; }
        public string CertificateThumbprint { get; set; }
    }

    public class CertificateMap : EntityTypeConfiguration<Certificate>
    {
        public CertificateMap()
        {
            // Primary Key
            this.HasKey(t => t.CertificateId);

            this.Property(t => t.CertificateId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.CertificateThumbprint)
                .HasMaxLength(200);

            // Table & Column Mappings
            this.ToTable("Certificates");
            this.Property(t => t.CertificateId).HasColumnName("CertificateId");
            this.Property(t => t.CertificateFile).HasColumnName("CertificateFile");
            this.Property(t => t.CertificateSubject).HasColumnName("CertificateSubject");
            this.Property(t => t.CertificateIssuer).HasColumnName("CertificateIssuer");
            this.Property(t => t.CertificateThumbprint).HasColumnName("CertificateThumbprint");
        }
    }
}
