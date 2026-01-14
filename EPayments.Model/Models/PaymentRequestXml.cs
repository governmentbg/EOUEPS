using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class PaymentRequestXml
    {
        public int PaymentRequestXmlId { get; set; }
        public string RequestContent { get; set; }
        public string ResponseContent { get; set; }
        public bool IsRequestAccepted { get; set; }
        public DateTime CreateDate { get; set; }
    }

    public class PaymentRequestXmlMap : EntityTypeConfiguration<PaymentRequestXml>
    {
        public PaymentRequestXmlMap()
        {
            // Primary Key
            this.HasKey(t => t.PaymentRequestXmlId);

            this.Property(t => t.PaymentRequestXmlId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.RequestContent)
                .IsRequired();

            this.Property(t => t.ResponseContent)
                .IsRequired();

            // Table & Column Mappings
            this.ToTable("PaymentRequestXmls");
            this.Property(t => t.PaymentRequestXmlId).HasColumnName("PaymentRequestXmlId");
            this.Property(t => t.RequestContent).HasColumnName("RequestContent");
            this.Property(t => t.ResponseContent).HasColumnName("ResponseContent");
            this.Property(t => t.IsRequestAccepted).HasColumnName("IsRequestAccepted");
            this.Property(t => t.CreateDate).HasColumnName("CreateDate");
        }
    }
}
