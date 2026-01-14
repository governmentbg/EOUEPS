using EPayments.Model.Enums;
using System;

namespace EPayments.Job.Host.ProcessTransactionFiles.Parsers.UniCredit
{
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

        public string InfoAC1AuthorizationCode { get; set; } // :86: +21-22 -> Авторизационен код (Само при тип на трамзакция AC1)
        public string InfoAC1BankCardInfo { get; set; } // :86: +21-22 ->  Данни за карта (Само при тип на трамзакция AC1)

        public TransactionRecordPaymentMethod InfoPaymentMethod { get; set; } // :86: +21-22 -> Начин на плащане

        public string InfoPaymentDetailsRaw { get; set; } // :86: +21-22 -> Свободно поле за повече информация / Детайли по плащане

        public string InfoPaymentReason { get; set; } // :86: +23-+27 -> Свободно поле за повече информация / Oснование на превода

        public string InfoSenderBic { get; set; } // :86: +30 -> Свободно поле за повече информация / BIC на наредителя на превода.
        public string InfoSenderIban { get; set; } // :86: +31 -> Свободно поле за повече информация / IBAN на наредителя на превода.
        public string InfoSenderName { get; set; } // :86: +32-33 -> Свободно поле за повече информация / Име на наредителя на превода.
    }
}
