using System;
using System.ComponentModel;

namespace EPayments.Model.Enums
{
    public enum CvposResponseCodesEnum
    {
        [Description("Одобрена или завършена успешно")]
        SuccessfullyCompleted = 0,
        [Description("Обърнете се към издателя на картата")]
        ReferToCardIssuer = 01,
        [Description("Вашата карта е ограничена")]
        PickUp = 04,
        [Description("Отказана транзакция")]
        DoNotHonour = 05,
        [Description("Грешка")]
        Error = 06,
        [Description("Невалидна транзакция")]
        InvalidTransaction = 12,
        [Description("Невалидна сума")]
        InvalidAmount = 13,
        [Description("Невалиден номер на картата")]
        NoSuchCard = 14,
        [Description("Анулиране от клиента")]
        CustomerCancellation = 17,
        [Description("Грешка във формата")]
        FormatError = 30,
        [Description("Придобиващ контакт с приемащ карта")]
        PickUpCardAcceptorContactAcquirer = 35,
        [Description("Карта с ограничение")]
        PickUpCardRestricted = 36,
        [Description("Сигурност на повикващия за приемане на карти")]
        PickUpCallAcquirerSecurity = 37,
        [Description("Допустимите опити за ПИН са надвишени")]
        PickUpAllowablePinTriesExceeded = 38,
        [Description("Няма кредитна сметка")]
        NoCreditAccount = 39,
        [Description("Заявената функция не се поддържа")]
        RequestedFunctionNotSupported = 40,
        [Description("Загубена карта")]
        PickUpLostCard = 41,
        [Description("Няма универсален акаунт")]
        NoUniversalAccount = 42,
        [Description("Открадната карта, вземане")]
        PickUpStolenCard = 43,
        [Description("Изтекла карта")]
        ExpiredCardTarget = 54,
        [Description("Неправилен личен идентификационен номер")]
        IncorrectPin = 55,
        [Description("Няма запис на карта")]
        NoCardRecord = 56,
        [Description("Транзакция не е разрешена на картодържателя")]
        TransactionNotPermittedToCardholder = 57,
        [Description("Транзакцията не е разрешена от терминал")]
        TransactionNotPermittedToTerminal = 58,
        [Description("Предполагаема измама")]
        SuspectedFraud = 59,
        [Description("Таймаут при издателя")]
        Timeout = 82,
        [Description("Отказано без причина")]
        NoReasonToDecline = 85,
        [Description("Криптографски отказ")]
        CryptographicFailure = 88,
        [Description("Неуспешно удостоверяване")]
        AuthenticationFailure = 89,
        [Description("Издателят или суичът не работи")]
        IssuerOrSwitchIsInoperative = 91,
        [Description("Примирете грешка")]
        ReconcileErrorAuthNotFound = 95,
        [Description("Неизправност на системата")]
        SystemMalfunction = 96,
    }
}
