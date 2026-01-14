using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class VposClient
    {
        public int VposClientId { get; set; }
        public string Name { get; set; }
        public string Alias { get; set; }
        public string PaymentRequestUrl { get; set; }
        public bool IsActive { get; set; }
    }

    public class VposClientMap : EntityTypeConfiguration<VposClient>
    {
        public VposClientMap()
        {
            // Primary Key
            this.HasKey(t => t.VposClientId);

            this.Property(t => t.VposClientId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.Name)
                .IsRequired()
                .HasMaxLength(200);

            this.Property(t => t.Alias)
                .IsRequired()
                .HasMaxLength(100);

            // Table & Column Mappings
            this.ToTable("VposClients");
            this.Property(t => t.VposClientId).HasColumnName("VposClientId");
            this.Property(t => t.Name).HasColumnName("Name");
            this.Property(t => t.Alias).HasColumnName("Alias");
            this.Property(t => t.PaymentRequestUrl).HasColumnName("PaymentRequestUrl");
            this.Property(t => t.IsActive).HasColumnName("IsActive");
        }
    }
}
