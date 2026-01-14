using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class VposResult
    {
        public int VposResultId { get; set; }
        public Guid Gid { get; set; }
        public int PaymentRequestId { get; set; }
        public string PostData { get; set; }
        public string PostUrl { get; set; }
        public bool IsPaymentSuccessfull { get; set; }
        public bool IsPaymentCanceledByUser { get; set; }
        public DateTime RequestDate { get; set; }
    }

    public class VposResultMap : EntityTypeConfiguration<VposResult>
    {
        public VposResultMap()
        {
            // Primary Key
            this.HasKey(t => t.VposResultId);

            this.Property(t => t.VposResultId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.PostData)
                .IsRequired();

            this.Property(t => t.PostUrl)
                .IsRequired();

            // Table & Column Mappings
            this.ToTable("VposResults");
            this.Property(t => t.VposResultId).HasColumnName("VposResultId");
            this.Property(t => t.Gid).HasColumnName("Gid");
            this.Property(t => t.PaymentRequestId).HasColumnName("PaymentRequestId");
            this.Property(t => t.PostData).HasColumnName("PostData");
            this.Property(t => t.PostUrl).HasColumnName("PostUrl");
            this.Property(t => t.IsPaymentSuccessfull).HasColumnName("IsPaymentSuccessfull");
            this.Property(t => t.IsPaymentCanceledByUser).HasColumnName("IsPaymentCanceledByUser");
            this.Property(t => t.RequestDate).HasColumnName("RequestDate");
        }
    }
}
