using EPayments.Model.DataObjects;
using EPayments.Model.Models;
using Newtonsoft.Json;
using System;
using System.Linq.Expressions;

namespace EPayments.Data.ViewObjects.Admin
{
    public class BoricaTransactionVO
    {
        public int BoricaTransactionId { get; set; }

        public string Order { get; set; }

        public decimal Amount { get; set; }

        public decimal? Fee { get; set; }

        public decimal? Commission { get; set; }

        public DateTime TransactionDate { get; set; }

        public string Card { get; set; }

        public DateTime? SettlementDate { get; set; }

        public string StatusMessage { get; set; }

        public string BoricaTransactionSettlement { get; set; }

        public BoricaTransactionSettlementJson BoricaTransactionSettlementJson { get; set; }

        public static Expression<Func<BoricaTransaction, BoricaTransactionVO>> MapFrom =>
            t => new BoricaTransactionVO()
            {
                BoricaTransactionId = t.BoricaTransactionId,
                Order = t.Order,
                Amount = t.Amount,
                Fee = t.Fee,
                Commission = t.Commission,
                TransactionDate = t.TransactionDate,
                Card = t.Card,
                SettlementDate = t.SettlementDate,
                StatusMessage = t.StatusMessage,
                BoricaTransactionSettlement = t.BoricaTransactionSettlement
            };
    }
}
