using System;
using System.Collections.Generic;
using System.Data.Entity.ModelConfiguration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Model.Models
{
    public class BoricaTransactionPaymentRequest
    {
        public int BoricaTransactionId { get; set; }
        public int PaymentRequestId { get; set; }
        public virtual BoricaTransaction BoricaTransaction { get; set; }
        public virtual PaymentRequest PaymentRequest { get; set; }
    }

    public class BoricaTransactionPaymentRequestMap : EntityTypeConfiguration<BoricaTransactionPaymentRequest>
    {
        public BoricaTransactionPaymentRequestMap()
        {
            this.HasKey(t => t.BoricaTransactionId);

            this.HasKey(t => t.PaymentRequestId);

            // Properties
            this.Property(t => t.BoricaTransactionId)
                .IsRequired();

            this.Property(t => t.PaymentRequestId)
                .IsRequired();

            //this.ToTable("BoricaTransactionPaymentRequest");
            //this.Property(t => t.BoricaTransactionId).HasColumnName("BoricaTransactionId");
            //this.Property(t => t.PaymentRequestId).HasColumnName("PaymentRequestId");
        }
    }
}