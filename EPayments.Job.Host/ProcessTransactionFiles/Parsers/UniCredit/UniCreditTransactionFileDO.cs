using System;
using System.Collections.Generic;

namespace EPayments.Job.Host.ProcessTransactionFiles.Parsers.UniCredit
{
    public class UniCreditTransactionFileDO
    {
        public DateTime? BankStatementDate { get; set; } //:20: -> Дата на извлечението, формат YYMMDD
        public string BankStatementIban { get; set; } //:25: -> Идентификация на сметката - IBAN
        public string BankStatementNumber { get; set; } //:28: -> Номер на извлечение за годината

        public List<UniCreditTransactionDO> Transactions { get; set; }

        public UniCreditTransactionFileDO()
        {
            this.Transactions = new List<UniCreditTransactionDO>();
        }
    }
}
