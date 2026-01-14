using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class LoginAttemptLog
    {
        public int LoginAttemptLogId { get; set; }
        public DateTime LogDate { get; set; }
        public string IP { get; set; }
        public byte[] CertificateFile { get; set; }
        public string CertificateIssuer { get; set; }
        public string CertificatePolicies { get; set; }
        public string CertificateSubject { get; set; }
        public string AlternativeSubject { get; set; }
        public string CertificateThumbprint { get; set; }
        public string Egn { get; set; }
        public string NameLatin { get; set; }
        public string ErrorCode { get; set; }
        public bool IsIisErrorOccurred { get; set; }
        public bool IsUesParsed { get; set; }
        public bool? IsPersonal { get; set; }
        public bool? IsEgnParsed { get; set; }
        public bool? IsNameParsed { get; set; }
        public bool IsLoginSuccessful { get; set; }
    }

    public class LoginAttemptLogMap : EntityTypeConfiguration<LoginAttemptLog>
    {
        public LoginAttemptLogMap()
        {
            // Primary Key
            this.HasKey(t => t.LoginAttemptLogId);

            this.Property(t => t.LoginAttemptLogId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.IP)
                .IsRequired()
                .HasMaxLength(50);

            this.Property(t => t.CertificateThumbprint)
                .HasMaxLength(200);

            this.Property(t => t.Egn)
                .HasMaxLength(50);

            this.Property(t => t.NameLatin)
                .HasMaxLength(250);

            this.Property(t => t.ErrorCode)
                .HasMaxLength(30);

            // Table & Column Mappings
            this.ToTable("LoginAttemptLogs");
            this.Property(t => t.LoginAttemptLogId).HasColumnName("LoginAttemptLogId");
            this.Property(t => t.LogDate).HasColumnName("LogDate");
            this.Property(t => t.IP).HasColumnName("IP");
            this.Property(t => t.CertificateFile).HasColumnName("CertificateFile");
            this.Property(t => t.CertificateIssuer).HasColumnName("CertificateIssuer");
            this.Property(t => t.CertificatePolicies).HasColumnName("CertificatePolicies");
            this.Property(t => t.CertificateSubject).HasColumnName("CertificateSubject");
            this.Property(t => t.AlternativeSubject).HasColumnName("AlternativeSubject");
            this.Property(t => t.CertificateThumbprint).HasColumnName("CertificateThumbprint");
            this.Property(t => t.Egn).HasColumnName("Egn");
            this.Property(t => t.NameLatin).HasColumnName("NameLatin");
            this.Property(t => t.ErrorCode).HasColumnName("ErrorCode");
            this.Property(t => t.IsIisErrorOccurred).HasColumnName("IsIisErrorOccurred");
            this.Property(t => t.IsUesParsed).HasColumnName("IsUesParsed");
            this.Property(t => t.IsPersonal).HasColumnName("IsPersonal");
            this.Property(t => t.IsEgnParsed).HasColumnName("IsEgnParsed");
            this.Property(t => t.IsNameParsed).HasColumnName("IsNameParsed");
            this.Property(t => t.IsLoginSuccessful).HasColumnName("IsLoginSuccessful");
        }
    }
}
