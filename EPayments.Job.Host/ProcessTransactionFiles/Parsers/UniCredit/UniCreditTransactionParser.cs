using EPayments.Common.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPayments.Job.Host.ProcessTransactionFiles.Parsers.UniCredit
{
    public class UniCreditTransactionParser
    {
        private string filePath;

        public UniCreditTransactionParser(string filePath)
        {
            this.filePath = filePath;
        }

        public UniCreditTransactionFileDO ExtractCreditTransactions()
        {
            UniCreditTransactionSectionsDO sectionsDO = new UniCreditTransactionSectionsDO();

            using (StreamReader reader = new StreamReader(this.filePath))
            {
                while (!reader.EndOfStream)
                {
                    string sectionRow = reader.ReadLine().Trim();

                    if (sectionRow.StartsWith(":20:"))
                    {
                        sectionsDO.Row20 = sectionRow.Substring(4);
                    }
                    else if (sectionRow.StartsWith(":25:"))
                    {
                        sectionsDO.Row25 = sectionRow.Substring(4);
                    }
                    else if (sectionRow.StartsWith(":28:"))
                    {
                        sectionsDO.Row28 = sectionRow.Substring(4);
                    }
                    else if (sectionRow.StartsWith(":61:"))
                    {
                        UniCreditTransactionEntryDO transactionEntryDO = new UniCreditTransactionEntryDO();
                        transactionEntryDO.Row61 = sectionRow.Substring(4);

                        sectionsDO.TransactionEntries.Add(transactionEntryDO);
                    }
                    else if (sectionRow.StartsWith(":86:"))
                    {
                        sectionsDO.TransactionEntries.Last().Row86 = sectionRow.Substring(4);
                    }
                }
            }

            return sectionsDO.Parse();
        }
    }
}
