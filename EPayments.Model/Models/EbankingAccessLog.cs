using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class EbankingAccessLog
    {
        public int EbankingAccessLogId { get; set; }
        public int EbankingClientId { get; set; }
        public int PaymentRequestId { get; set; }
        public string IpAddress { get; set; }
        public DateTime AccessDate { get; set; }
    }

    public class EbankingAccessLogMap : EntityTypeConfiguration<EbankingAccessLog>
    {
        public EbankingAccessLogMap()
        {
            // Primary Key
            this.HasKey(t => t.EbankingAccessLogId);

            this.Property(t => t.EbankingAccessLogId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.IpAddress)
                .IsRequired()
                .HasMaxLength(50);

            // Table & Column Mappings
            this.ToTable("EbankingAccessLogs");
            this.Property(t => t.EbankingAccessLogId).HasColumnName("EbankingAccessLogId");
            this.Property(t => t.EbankingClientId).HasColumnName("EbankingClientId");
            this.Property(t => t.IpAddress).HasColumnName("IpAddress");
            this.Property(t => t.AccessDate).HasColumnName("AccessDate");
        }
    }
}
