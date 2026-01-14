using System;

namespace EPayments.Model.DataObjects
{
    public class BoricaTransactionSettlementJson
    {
        public string AreaOfIssue { get; set; }
        public string AuthorizationCode { get; set; }
        public string CardBrand { get; set; }
        public bool DistributionDateSpecified { get; set; }
        public string OrderId { get; set; }
        public string ProductCategory { get; set; }
        public string Rrn { get; set; }
        public DateTime? SettlementDate { get; set; }
        public bool SettlementDateSpecified { get; set; }
        public decimal? Sum { get; set; }
        public bool SumSpecified { get; set; }
        public decimal? Tax { get; set; }
        public bool TaxSpecified { get; set; }
        public decimal? TaxBorica { get; set; }
        public bool TaxBoricaSpecified { get; set; }
        public string Tid { get; set; }
        public DateTime? TransactionDate { get; set; }
        public bool TransactionDateSpecified { get; set; }
    }
}
