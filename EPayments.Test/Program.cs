using EPayments.Common.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Dictionary<string, UniCreditTransactionFileDO> transactionFiles = new Dictionary<string, UniCreditTransactionFileDO>();

            foreach (string file in Directory.GetFiles(@"D:\Projects Files\EPayments\UniCredit transaction files"))
            {
                UniCreditTransactionParser parser = new UniCreditTransactionParser(file);
                transactionFiles.Add(file, parser.ExtractCreditTransactions());
            }

            //TODO: check if UniCreditTransactionFileDO IBAN match expected IBAN
        }
    }

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

    public class UniCreditTransactionDO
    {
        public DateTime? TransactionDate { get; set; } //:61: -> Запис за транзакция / Вальор
        public DateTime? TransactionAccountingDate { get; set; } //:61: -> Запис за транзакция / Счетоводна дата
        public decimal? TransactionAmount { get; set; } //:61: -> Запис за транзакция / Сума на транзакцията 
        public string TransactionReferenceId { get; set; } //:61: -> Запис за транзакция / Банкова референция за транзакцията

        public string InfoSystemTransactionType { get; set; } // :86: <first 3 characters> -> Свободно поле за повече информация / Вид на транзакцията в банката по номенклатура
        public string InfoSystemTransactionDesc { get; set; } // :86: +00 -> Свободно поле за повече информация / Описание на транзакцията

        public string InfoPaymentType { get; set; } // :86: +21-22 PAY -> Свободно поле за повече информация / Код за вид плащане
        public string InfoDocumentType { get; set; } // :86: +21-22 DOC -> Свободно поле за повече информация / Вид на документа
        public string InfoDocumentNumber { get; set; } // :86: +21-22 NUM -> Свободно поле за повече информация / Номер на документа
        public DateTime? InfoDocumentDate { get; set; } // :86: +21-22 DAT -> Свободно поле за повече информация / Дата на документа
        public DateTime? InfoPaymentPeriodBegining { get; set; } // :86: +21-22 BEG -> Свободно поле за повече информация / Начало на периода, за който се плаща
        public DateTime? InfoPaymentPeriodEnd { get; set; } // :86: +21-22 END -> Свободно поле за повече информация / Край на периода, за който се плаща
        public string InfoDebtorBulstat { get; set; } // :86: +21-22 BUL -> Свободно поле за повече информация / БУЛСТАТ идентификатор на задълженото лице
        public string InfoDebtorEgn { get; set; } // :86: +21-22 EGN -> Свободно поле за повече информация / ЕГН идентификатор на задълженото лице
        public string InfoDebtorLnch { get; set; } // :86: +21-22 LNC -> Свободно поле за повече информация / ЛНЧ идентификатор на задълженото лице
        public string InfoDebtorName { get; set; } // :86: +21-22 IZL -> Свободно поле за повече информация / Име на задълженото лице

        public string InfoPaymentDetailsRaw { get; set; } // :86: +21-22 -> Свободно поле за повече информация / Детайли по плащане

        public string InfoPaymentReason { get; set; } // :86: +23-+27 -> Свободно поле за повече информация / Oснование на превода

        public string InfoSenderBic { get; set; } // :86: +30 -> Свободно поле за повече информация / BIC на наредителя на превода.
        public string InfoSenderIban { get; set; } // :86: +31 -> Свободно поле за повече информация / IBAN на наредителя на превода.
        public string InfoSenderName { get; set; } // :86: +32-33 -> Свободно поле за повече информация / Име на наредителя на превода.
    }

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
                //TODO: test code
                if (String.IsNullOrWhiteSpace(transactionEntryDO.Row61) || String.IsNullOrWhiteSpace(transactionEntryDO.Row86))
                {
                    throw new Exception();
                }

                if (!transactionEntryDO.SkipTransaction())
                {
                    transactionFileDO.Transactions.Add(transactionEntryDO.Parse());
                }
            }

            return transactionFileDO;
        }
    }

    public class UniCreditTransactionEntryDO
    {
        public string Row61 { get; set; }
        public string Row86 { get; set; }

        public bool SkipTransaction()
        {
            char creditSymbol = this.Row61[10];

            return creditSymbol != 'C';
        }

        public UniCreditTransactionDO Parse()
        {
            UniCreditTransactionDO transactionDO = new UniCreditTransactionDO();

            transactionDO.TransactionDate = Parser.DateParse(this.Row61.Substring(0, 6), DateFormat.yy_MM_dd_NoSeparator);
            transactionDO.TransactionAccountingDate = Parser.DateParse(this.Row61.Substring(0, 2) + this.Row61.Substring(6, 4), DateFormat.yy_MM_dd_NoSeparator);

            int indexOfDecimalDelimiter = this.Row61.Substring(11).IndexOf(',');
            string amountString = this.Row61.Substring(11, indexOfDecimalDelimiter + 3).Replace(",", ".");

            transactionDO.TransactionAmount = Parser.TwoDecimalPlacesFormatStringToDecimal(amountString);

            int indexOfReferenceDelimiter = this.Row61.Substring(indexOfDecimalDelimiter + 3).IndexOf("//");
            if (indexOfReferenceDelimiter != -1)
            {
                transactionDO.TransactionReferenceId = this.Row61.Substring(indexOfDecimalDelimiter + 3 + indexOfReferenceDelimiter + 2);
            }

            transactionDO.InfoSystemTransactionType = this.Row86.Substring(0, 3);

            List<string> sectionSearchKeys = new List<string>()
            {
                "+00",
                "+21",
                "+22",
                "+23",
                "+24",
                "+25",
                "+26",
                "+27",
                "+30",
                "+31",
                "+32",
                "+33",
            };

            Dictionary<string, string> sectionDictionary = ExtractOrderedKeyValues(this.Row86.Substring(3), sectionSearchKeys);

            string str00 = sectionDictionary["+00"] ?? "";
            string str21Concat22 = (sectionDictionary["+21"] ?? "") + (sectionDictionary["+22"] ?? "");
            string str23Concat27 =
                (!String.IsNullOrWhiteSpace(sectionDictionary["+23"]) ? sectionDictionary["+23"] + " " : "") +
                (!String.IsNullOrWhiteSpace(sectionDictionary["+24"]) ? sectionDictionary["+24"] + " " : "") +
                (!String.IsNullOrWhiteSpace(sectionDictionary["+25"]) ? sectionDictionary["+25"] + " " : "") +
                (!String.IsNullOrWhiteSpace(sectionDictionary["+26"]) ? sectionDictionary["+26"] + " " : "") +
                (!String.IsNullOrWhiteSpace(sectionDictionary["+27"]) ? sectionDictionary["+27"] : "");
            string str30 = sectionDictionary["+30"] ?? "";
            string str31 = sectionDictionary["+31"] ?? "";
            string str32Concat33 = (sectionDictionary["+32"] ?? "") + (sectionDictionary["+33"] ?? "");

            transactionDO.InfoSystemTransactionDesc = str00;
            transactionDO.InfoPaymentDetailsRaw = str21Concat22;
            transactionDO.InfoPaymentReason = str23Concat27;
            transactionDO.InfoSenderBic = str30;
            transactionDO.InfoSenderIban = str31;
            transactionDO.InfoSenderName = str32Concat33;

            List<string> propertyKeys = new List<string>()
            {
                "/PAY",
                "/DOC",
                "/NUM",
                "/DAT",
                "/BEG",
                "/END",
                "/BUL",
                "/EGN",
                "/LNC",
                "/IZL",
            };

            Dictionary<string, string> propertyDictionary = ExtractOrderedKeyValues("/" + str21Concat22, propertyKeys);

            transactionDO.InfoPaymentType = propertyDictionary["/PAY"];
            transactionDO.InfoDocumentType = propertyDictionary["/DOC"];
            transactionDO.InfoDocumentNumber = propertyDictionary["/NUM"];
            if (!String.IsNullOrWhiteSpace(propertyDictionary["/DAT"]))
            {
                transactionDO.InfoDocumentDate = Parser.DateParse(propertyDictionary["/DAT"], DateFormat.yy_MM_dd_NoSeparator);
            }
            if (!String.IsNullOrWhiteSpace(propertyDictionary["/BEG"]))
            {
                transactionDO.InfoPaymentPeriodBegining = Parser.DateParse(propertyDictionary["/BEG"], DateFormat.yy_MM_dd_NoSeparator);
            }
            if (!String.IsNullOrWhiteSpace(propertyDictionary["/END"]))
            {
                transactionDO.InfoPaymentPeriodEnd = Parser.DateParse(propertyDictionary["/END"], DateFormat.yy_MM_dd_NoSeparator);
            }
            transactionDO.InfoDebtorBulstat = propertyDictionary["/BUL"];
            transactionDO.InfoDebtorEgn = propertyDictionary["/EGN"];
            transactionDO.InfoDebtorLnch = propertyDictionary["/LNC"];
            transactionDO.InfoDebtorName = propertyDictionary["/IZL"];

            return transactionDO;
        }

        private Dictionary<string, string> ExtractOrderedKeyValues(string stringContent, List<string> searchKeys)
        {
            Dictionary<string, string> sectionDictionary = new Dictionary<string, string>();
            foreach(string searchKey in searchKeys)
            {
                sectionDictionary.Add(searchKey, null);
            }

            while (searchKeys.Count > 0 && stringContent.Length > 0)
            {
                string currSectionKey = searchKeys.First();
                int indexOfCurrSectionKey = stringContent.IndexOf(currSectionKey);

                searchKeys.RemoveAt(0);

                if (indexOfCurrSectionKey != -1)
                {
                    stringContent = stringContent.Substring(indexOfCurrSectionKey + currSectionKey.Length);
                    int indexOfNextSectionKey = -1;

                    while (searchKeys.Count > 0 && stringContent.Length > 0)
                    {
                        string nextSectionKey = searchKeys[0];
                        indexOfNextSectionKey = stringContent.IndexOf(nextSectionKey);

                        if (indexOfNextSectionKey == -1)
                        {
                            searchKeys.RemoveAt(0);
                        }
                        else
                        {
                            break;
                        }
                    }

                    if (indexOfNextSectionKey == -1)
                    {
                        sectionDictionary[currSectionKey] = stringContent;
                        break;
                    }
                    else
                    {
                        sectionDictionary[currSectionKey] = stringContent.Substring(0, indexOfNextSectionKey);

                        stringContent = stringContent.Substring(indexOfNextSectionKey);
                    }
                }
            }

            return sectionDictionary;
        }
    }


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
