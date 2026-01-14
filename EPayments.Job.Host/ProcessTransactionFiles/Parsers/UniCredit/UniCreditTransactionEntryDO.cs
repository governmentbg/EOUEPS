using EPayments.Common.Helpers;
using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPayments.Job.Host.ProcessTransactionFiles.Parsers.UniCredit
{
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

            transactionDO.InfoPaymentMethod = TransactionRecordPaymentMethod.BankOrder;

            //При тип на транзакцията "Операция с карта"
            if (transactionDO.InfoSystemTransactionType != null && transactionDO.InfoSystemTransactionType.Trim().ToLower() == "ac1")
            {
                if (!String.IsNullOrWhiteSpace(transactionDO.InfoPaymentDetailsRaw))
                {
                    if (transactionDO.InfoPaymentDetailsRaw.TrimStart().StartsWith("ПОС "))
                    {
                        transactionDO.InfoPaymentMethod = TransactionRecordPaymentMethod.POS;
                    }
                    else if (transactionDO.InfoPaymentDetailsRaw.TrimStart().StartsWith("Плащане /импринтер/ ") ||
                             transactionDO.InfoPaymentDetailsRaw.TrimStart().StartsWith("EPAY "))
                    {
                        transactionDO.InfoPaymentMethod = TransactionRecordPaymentMethod.VPOS;
                    }

                    string authStr = "авт.код:";

                    int authStrIndex = transactionDO.InfoPaymentDetailsRaw.IndexOf(authStr);
                    if (authStrIndex != -1)
                    {
                        int authEndIndex = transactionDO.InfoPaymentDetailsRaw.Substring(authStrIndex + authStr.Length).IndexOf(" ");
                        if (authEndIndex != -1)
                        {
                            transactionDO.InfoAC1AuthorizationCode = transactionDO.InfoPaymentDetailsRaw.Substring(authStrIndex + authStr.Length, authEndIndex).Trim();
                        }
                    }

                    string panStr = "PAN:";

                    int panStrIndex = transactionDO.InfoPaymentDetailsRaw.IndexOf(panStr);
                    if (panStrIndex != -1)
                    {
                        transactionDO.InfoAC1BankCardInfo = transactionDO.InfoPaymentDetailsRaw.Substring(panStrIndex + panStr.Length).Trim();
                    }
                }
            }

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
}
