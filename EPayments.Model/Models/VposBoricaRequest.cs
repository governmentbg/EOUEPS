using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class VposBoricaRequest
    {
        public int VposBoricaRequestId { get; set; }
        public string RequestIdentifier { get; set; }
        public int PaymentRequestId { get; set; }
        public int? VposRedirectId { get; set; }
        public string RedirectUrl { get; set; }
        public DateTime CreateDate { get; set; }
    }

    public class VposBoricaRequestMap : EntityTypeConfiguration<VposBoricaRequest>
    {
        public VposBoricaRequestMap()
        {
            // Primary Key
            this.HasKey(t => t.VposBoricaRequestId);

            this.Property(t => t.VposBoricaRequestId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.RequestIdentifier)
                .IsRequired()
                .HasMaxLength(15);

            this.Property(t => t.RedirectUrl)
                .IsRequired();

            // Table & Column Mappings
            this.ToTable("VposBoricaRequests");
            this.Property(t => t.VposBoricaRequestId).HasColumnName("VposBoricaRequestId");
            this.Property(t => t.RequestIdentifier).HasColumnName("RequestIdentifier");
            this.Property(t => t.PaymentRequestId).HasColumnName("PaymentRequestId");
            this.Property(t => t.VposRedirectId).HasColumnName("VposRedirectId");
            this.Property(t => t.RedirectUrl).HasColumnName("RedirectUrl");
            this.Property(t => t.CreateDate).HasColumnName("CreateDate");
        }
    }
}
