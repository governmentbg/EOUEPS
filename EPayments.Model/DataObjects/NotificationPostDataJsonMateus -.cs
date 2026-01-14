namespace EPayments.Model.DataObjects
{
    public class NotificationPostDataJsonMateus
    {
        public string ApplicantUin { get; set; }

        public string PaymentRequestIdentifier { get; set; }

        public string TransactionIdentifier { get; set; }

        public string Status { get; set; }

        public string ChangeTime { get; set; }

        public string PaidInstalmentSum { get; set; }

        public string DebtInstalmentId { get; set; }

        public string PaidInterestSum { get; set; }

        public string ТaxSubjectId { get; set; }
    }
}
