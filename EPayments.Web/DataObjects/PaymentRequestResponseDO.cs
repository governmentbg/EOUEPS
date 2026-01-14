namespace EPayments.Web.DataObjects
{
    public class PaymentRequestResponseDO
    {
        public AcceptedReceiptJsonDO AcceptedReceiptJson { get; set; }

        public UnacceptedReceiptJsonDO UnacceptedReceiptJson { get; set; }
    }
}