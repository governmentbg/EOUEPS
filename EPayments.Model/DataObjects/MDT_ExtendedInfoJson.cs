namespace EPayments.Model.DataObjects
{
    public class MDT_ExtendedInfoJson
    {
        public string PartidaNo { get; set; }
        public string RegisterNo { get; set; }
        public string PropertyAddress { get; set; }
        public string KindDebtRegName { get; set; }
        public int TaxPeriodYear { get; set; }
        public int InstNo { get; set; }

        public string regionClientId { get; set; }

        public string regionName { get; set; }

        public string DebtInstalmentId { get; set; }

        public decimal PaidInstalmentSum { get; set; }

        public decimal PaidInterestSum { get; set; }

        public string taxSubjectId { get; set; }
    }
}
