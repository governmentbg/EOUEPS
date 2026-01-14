using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class EbankingClient
    {
        public int EbankingClientId { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string ClientId { get; set; }
        public string SecretKey { get; set; }
        public bool IsActive { get; set; }
    }

    public class EbankingClientMap : EntityTypeConfiguration<EbankingClient>
    {
        public EbankingClientMap()
        {
            // Primary Key
            this.HasKey(t => t.EbankingClientId);

            this.Property(t => t.EbankingClientId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.Name)
                .IsRequired()
                .HasMaxLength(200);

            this.Property(t => t.ClientId)
                .IsRequired()
                .HasMaxLength(100);

            this.Property(t => t.SecretKey)
                .IsRequired()
                .HasMaxLength(200);

            // Table & Column Mappings
            this.ToTable("EbankingClients");
            this.Property(t => t.EbankingClientId).HasColumnName("EbankingClientId");
            this.Property(t => t.Name).HasColumnName("Name");
            this.Property(t => t.Description).HasColumnName("Description");
            this.Property(t => t.ClientId).HasColumnName("ClientId");
            this.Property(t => t.SecretKey).HasColumnName("SecretKey");
            this.Property(t => t.IsActive).HasColumnName("IsActive");
        }
    }
}
