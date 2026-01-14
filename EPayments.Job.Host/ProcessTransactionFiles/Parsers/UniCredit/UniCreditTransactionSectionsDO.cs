using EPayments.Common.Helpers;
using System;
using System.Collections.Generic;

namespace EPayments.Job.Host.ProcessTransactionFiles.Parsers.UniCredit
{
    public class UniCreditTransactionSectionsDO
    {
        public string Row20 { get; set; }
        public string Row25 { get; set; }
        public string Row28 { get; set; }
        public List<UniCreditTransactionEntryDO> TransactionEntries { get; set; }

        public UniCreditTransactionSectionsDO()
        {
            this.TransactionEntries = new List<UniCreditTransactionEntryDO>();
        }

        public UniCreditTransactionFileDO Parse()
        {
            UniCreditTransactionFileDO transactionFileDO = new UniCreditTransactionFileDO();

            if (!String.IsNullOrWhiteSpace(this.Row20))
            {
                transactionFileDO.BankStatementDate = Parser.DateParse(this.Row20, DateFormat.yy_MM_dd_NoSeparator);
            }

            transactionFileDO.BankStatementIban = this.Row25;
            transactionFileDO.BankStatementNumber = this.Row28;

            foreach (UniCreditTransactionEntryDO transactionEntryDO in this.TransactionEntries)
            {
                if (!transactionEntryDO.SkipTransaction())
                {
                    transactionFileDO.Transactions.Add(transactionEntryDO.Parse());
                }
            }

            return transactionFileDO;
        }
    }
}
