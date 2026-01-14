using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class VposRedirect
    {
        public int VposRedirectId { get; set; }
        public Guid Gid { get; set; }
        public string PaymentRequestIdentifier { get; set; }
        public string OkUrl { get; set; }
        public string CancelUrl { get; set; }
        public string IP { get; set; }
        public DateTime LogDate { get; set; }
    }

    public class VposRedirectMap : EntityTypeConfiguration<VposRedirect>
    {
        public VposRedirectMap()
        {
            // Primary Key
            this.HasKey(t => t.VposRedirectId);

            this.Property(t => t.VposRedirectId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.PaymentRequestIdentifier)
                .HasMaxLength(50);

            this.Property(t => t.IP)
                .HasMaxLength(50);

            // Table & Column Mappings
            this.ToTable("VposRedirects");
            this.Property(t => t.VposRedirectId).HasColumnName("VposRedirectId");
            this.Property(t => t.Gid).HasColumnName("Gid");
            this.Property(t => t.PaymentRequestIdentifier).HasColumnName("PaymentRequestIdentifier");
            this.Property(t => t.OkUrl).HasColumnName("OkUrl");
            this.Property(t => t.CancelUrl).HasColumnName("CancelUrl");
            this.Property(t => t.IP).HasColumnName("IP");
            this.Property(t => t.LogDate).HasColumnName("LogDate");
        }
    }
}
