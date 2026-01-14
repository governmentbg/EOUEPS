using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EPayments.Data.ViewObjects.Web.APGModels
{
    public class APGWPaymentResponseDataDO : APGWResponseBase
    {
        private string _rc;
        private string _approval;
        private string _currency;
        private string _rrn;
        private string _intRef;
        private string _paresStatus;
        private string _eci;
        private string _adCustomBorOrderId;

        [Display(Name = "ACTION")]
        public int Action { get; set; }

        [Display(Name = "RC")]
        public string Rc 
        { 
            get { return _rc; }
            set
            {
                if (!this.ValidateString(value, 4))
                {
                    throw new ArgumentException("Rc maximum length is 4.");
                }

                this._rc = value;
            }
        }

        [Display(Name = "APPROVAL")]
        public string Approval 
        {
            get { return this._approval == null ? "-" : this._approval; }
            set
            {
                if (!this.ValidateString(value, 6))
                {
                    throw new ArgumentException("Approval maximum length is 6.");
                }

                this._approval = value;
            }
        }

        [Display(Name = "TRTYPE")]
        public override TrTypeEnum TrType { get; set; }

        [DataType(DataType.Currency)]
        public decimal? AmountValue { get; set; }

        [Display(Name = "AMOUNT")]
        public string Amount
        {
            get
            {
                return this.AmountValue == null ? null : ((decimal)this.AmountValue).ToString("#0.00");
            }

            set
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    this.AmountValue = null;
                }
                else
                {
                    this.AmountValue = decimal.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
                }
            }
        }

        [Display(Name = "CURRENCY")]
        public string Currency 
        {
            get { return this._currency == null ? "-" : this._currency; }
            set
            {
                if (!this.ValidateString(value, 3))
                {
                    throw new ArgumentException("Currency maximum length is 3.");
                }

                this._currency = value;
            }
        }

        [Display(Name = "RRN")]
        public string Rrn 
        {
            get { return this._rrn == null ? "-" : this._rrn; }
            set
            {
                if (!this.ValidateString(value, 12))
                {
                    throw new ArgumentException("Rrn length is 12.");
                }

                this._rrn = value;
            }
        }

        [Display(Name = "INT_REF")]
        public string INT_REF 
        {
            get { return this._intRef == null ? "-" : this._intRef; }
            set
            {
                if (!this.ValidateString(value, 32))
                {
                    throw new ArgumentException("IntRef maximum length is 32.");
                }

                this._intRef = value;
            }
        }

        [Display(Name = "PARES_STATUS")]
        public string Pares_Status 
        { 
            get { return this._paresStatus == null ? "-" : this._paresStatus; }
            set
            {
                if (!this.ValidateString(value, 20))
                {
                    throw new ArgumentException("ParesStatus maximum length is 20.");
                }

                this._paresStatus = value;
            }
        }

        [Display(Name = "ECI")]
        public string Eci 
        { 
            get { return this._eci == null ? "-" : this._eci; }
            set
            {
                if (!this.ValidateString(value, 2))
                {
                    throw new ArgumentException("Eci maximum length is 2.");
                }

                this._eci = value;
            }
        }

        [Display(Name = "AD.CUST_BOR_ORDER_ID")]
        public string AdCustomBorOrderId
        {
            get { return this._adCustomBorOrderId; }
            set
            {
                if (!this.ValidateString(value, 22))
                {
                    throw new ArgumentException("AdCustomBorOrderId maximum length is 22.");
                }

                this._adCustomBorOrderId = value;
            }
        }

        [Display(Name = "NONCE")]
        public override string Nonce { get; set; }

        [Display(Name = "CARD")]
        public string Card { get; set; }

        [Display(Name = "LANG")]
        public string Lang { get; set; }

        [Display(Name = "STATUSMSG")]
        public string StatusMsg { get; set; }

        public override bool IsResponseValid()
        {
            var pSign = this.GetPSignData();
            
            if (string.IsNullOrEmpty(pSign))
            {
                throw new NullReferenceException(nameof(pSign));
            }

            int index = 0;
            
            ParsedPart nextPart = this.ParseNextPart(pSign, index, this.Action.ToString().Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateIntegers(this.Action, nextPart.Message))
            {
                return false;
            }

            index = nextPart.NextIndex;

            nextPart = this.ParseNextPart(pSign, index, Rc.Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateStrings(this.Rc, nextPart.Message))
            {
                return false;
            }

            index = nextPart.NextIndex;

            nextPart = this.ParseNextPart(pSign, index, Approval.Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateStrings(this.Approval, nextPart.Message))
            {
                return false;
            }

            index = nextPart.NextIndex;

            nextPart = this.ParseNextPart(pSign, index, Terminal.Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateStrings(this.Terminal, nextPart.Message))
            {
                return false;
            }

            index = nextPart.NextIndex;

            nextPart = this.ParseNextPart(pSign, index, ((int)this.TrType).ToString().Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateStrings(((int)this.TrType).ToString(), nextPart.Message))
            {
                return false;
            }

            index = nextPart.NextIndex;

            nextPart = this.ParseNextPart(pSign, index, this.AmountValue.ToString().Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateDecimal(this.AmountValue, nextPart.Message))
            {
                return false;
            }

            index = nextPart.NextIndex;

            nextPart = this.ParseNextPart(pSign, index, Currency.Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateStrings(this.Currency, nextPart.Message))
            {
                return false;
            }

            index = nextPart.NextIndex;

            nextPart = this.ParseNextPart(pSign, index, Order.Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateStrings(this.Order, nextPart.Message))
            {
                return false;
            }

            index = nextPart.NextIndex;

            nextPart = this.ParseNextPart(pSign, index, Rrn.Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateStrings(this.Rrn, nextPart.Message))
            {
                return false;
            }

            index = nextPart.NextIndex;

            nextPart = this.ParseNextPart(pSign, index, INT_REF.Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateStrings(this.INT_REF, nextPart.Message))
            {
                return false;
            }

            index = nextPart.NextIndex;

            nextPart = this.ParseNextPart(pSign, index, Pares_Status.Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateStrings(this.Pares_Status, nextPart.Message))
            {
                return false;
            }

            index = nextPart.NextIndex;

            nextPart = this.ParseNextPart(pSign, index, Eci.Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateStrings(this.Eci, nextPart.Message))
            {
                return false;
            }

            index = nextPart.NextIndex;

            nextPart = this.ParseNextPart(pSign, index, TimeStamp.Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateStrings(this.TimeStamp, nextPart.Message))
            {
                return false;
            }          

            index = nextPart.NextIndex;

            nextPart = this.ParseNextPart(pSign, index, Nonce.Length);

            if (nextPart == null)
            {
                return false;
            }

            if (!this.ValidateStrings(this.Nonce, nextPart.Message))
            {
                return false;
            }

            return true;
        }

        public void ParseFromDictionary(Dictionary<string, object> keyValuePairs)
        {
            try
            {
                foreach (string key in keyValuePairs.Keys)
                {
                    if (keyValuePairs[key] != null)
                    {
                        string value = keyValuePairs[key].ToString();

                        switch (key)
                        {
                            case "ACTION":
                                this.Action = int.Parse(value);
                                break;
                            case "RC":
                                this.Rc = value;
                                break;
                            case "STATUSMSG":
                                this.StatusMsg = value;
                                break;
                            case "TERMINAL":
                                this.Terminal = value;
                                break;
                            case "TRTYPE":
                                this.TrType = (TrTypeEnum)Enum.Parse(typeof(TrTypeEnum), value);
                                break;
                            case "AMOUNT":
                                if (!string.IsNullOrWhiteSpace(value))
                                {
                                    this.AmountValue = decimal.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
                                }
                                break;
                            case "CURRENCY":
                                this.Currency = value;
                                break;
                            case "ORDER":
                                this.Order = value;
                                break;
                            case "TIMESTAMP":
                                this.TimeStamp = value;
                                break;
                            case "APPROVAL":
                                this.Approval = value;
                                break;
                            case "RRN":
                                this.Rrn = value;
                                break;
                            case "INT_REF":
                                this.INT_REF = value;
                                break;
                            case "PARES_STATUS":
                                this.Pares_Status = value;
                                break;
                            case "ECI":
                                this.Eci = value;
                                break;
                            case "CARD":
                                this.Card = value;
                                break;
                            case "NONCE":
                                this.Nonce = value;
                                break;
                            case "P_SIGN":
                                this.P_Sign = value;
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Invalid apgw payment response.", ex);
            }
        }

        public string GetBoricaRCDescription(string errorCode)
        {
            Dictionary<string, string> responseCodes = new Dictionary<string, string>();
            responseCodes.Add("-1", "В заявката не е попълнено задължително поле");
            responseCodes.Add("-2", "Заявката съдържа поле с некоректно име");
            responseCodes.Add("-3", "Aвторизационният хост не отговаря или форматът на отговора е неправилен");
            responseCodes.Add("-4", "Няма връзка с авторизационния хост");
            responseCodes.Add("-5", "Грешка във връзката с авторизационния хост");
            responseCodes.Add("-6", "Грешка в конфигурацията на APGW");
            responseCodes.Add("-7", "Форматът на отговора от авторизационния хост е неправилен");
            responseCodes.Add("-10", "Грешка в поле 'Сума' в заявката");
            responseCodes.Add("-11", "Грешка в поле 'Валута' в заявката");
            responseCodes.Add("-12", "Грешка в поле 'Идентификатор на търговеца' в заявката");
            responseCodes.Add("-13", "Неправилен IP адрес на търговеца");
            responseCodes.Add("-15", "Грешка в поле 'RRN' в заявката");
            responseCodes.Add("-16", "В момента се изпълнява друга трансакция на терминала");
            responseCodes.Add("-17", "Отказан достъп до платежния сървър (напр. грешка при проверка на P_SIGN)");
            responseCodes.Add("-19", "Грешка в искането за автентикация или неуспешна автентикация");
            responseCodes.Add("-20", "Разрешената разлика между времето на сървъра на търговеца и e-Gateway сървъра е надвишена");
            responseCodes.Add("-21", "Трансакцията вече е била изпълнена");
            responseCodes.Add("-22", "Трансакцията съдържа невалидни данни за автентикация");
            responseCodes.Add("-23", "Невалиден контекст на транзакцията");
            responseCodes.Add("-24", "Заявката съдържа стойности за полета, които не могат да бъдат обработени. Например валутата е раз-лична от валутата на терминала или транзакцията е по-стара от 24 часа.");
            responseCodes.Add("-25", "Трансакцията е отказана (напр. от картодържателя)");
            responseCodes.Add("-26", "Невалиден BIN на картата");
            responseCodes.Add("-27", "Невалидно име на търговеца");
            responseCodes.Add("-28", "Невалидно допълнително поле (например AD.CUST_BOR_ORDER_ID)");
            responseCodes.Add("-29", "Невалиден отговор от ACS на издателя на картата");
            responseCodes.Add("-30", "Трансакцията е отказана");
            responseCodes.Add("-31", "Трансакцията е в процес на обработка");
            responseCodes.Add("-32", "Дублирана отказана трансакция");
            responseCodes.Add("-33", "Трансакцията е в процес на аутентикация на карто-държателя");
            responseCodes.Add("-40", "Трансакцията е в процес на обработка");
            responseCodes.Add("00", "Успешно завършена трансакция"); // Successfully completed
            responseCodes.Add("01", "Обърнете се към издателя на картата"); // Refer to card issuer
            responseCodes.Add("04", "Вашата карта е ограничена"); // PICK UP
            responseCodes.Add("05", "Отказана транзакция"); // Do not Honour
            responseCodes.Add("06", "Грешка"); // Error
            responseCodes.Add("12", "Невалидна транзакция");// Invalid transaction
            responseCodes.Add("13", "Невалидна сума"); // Invalid amount
            responseCodes.Add("14", "Невалиден номер на картата"); // No such card
            responseCodes.Add("15", "Няма такъв емитент"); // No such issuer
            responseCodes.Add("17", "Анулиране от клиента"); // Customer cancellation
            responseCodes.Add("30", "Грешка във формата"); // Format error
            responseCodes.Add("35", "Придобиващ контакт с приемащ карта"); // Pick-up, card acceptor contact acquirer
            responseCodes.Add("36", "Карта с ограничение"); // Pick up, card restricted
            responseCodes.Add("37", "Сигурност на повикващия за приемане на карти"); // Pick up, call acquirer security
            responseCodes.Add("38", "Допустимите опити за ПИН са надвишени"); // Pick up, Allowable PIN tries exceeded
            responseCodes.Add("39", "Няма кредитна сметка"); // No credit account
            responseCodes.Add("40", "Заявената функция не се поддържа"); // Requested function not supported
            responseCodes.Add("41", "Загубена карта"); // Pick up, lost card
            responseCodes.Add("42", "Няма универсален акаунт"); // No universal account
            responseCodes.Add("43", "Открадната карта, вземане"); // Pick up, stolen card
            responseCodes.Add("54", "Изтекла карта"); // Expired card / target
            responseCodes.Add("55", "Неправилен ПИН"); // Incorrect PIN
            responseCodes.Add("56", "Няма запис на карта"); // No card record
            responseCodes.Add("57", "Транзакция не е разрешена на картодържателя"); // Transaction not permitted to cardholder
            responseCodes.Add("58", "Транзакцията не е разрешена от терминал"); // Transaction not permitted to terminal
            responseCodes.Add("59", "Предполагаема измама"); // Suspected fraud
            responseCodes.Add("82", "Таймаут при издателя");
            responseCodes.Add("85", "Отказано без причина"); // No reason to decline
            responseCodes.Add("88", "Криптографски отказ"); // Cryptographic failure
            responseCodes.Add("89", "Неуспешно удостоверяване"); // Authentication failure
            responseCodes.Add("91", "Издателят или суичът не работи"); // Issuer or switch is inoperative
            responseCodes.Add("95", "Грешка при съгласуване"); // Reconcile error / Auth Not found
            responseCodes.Add("96", "Неизправност на системата"); // System Malfunction

            return responseCodes.ContainsKey(errorCode) ? responseCodes[errorCode] : "Грешка при трансакцията";
        }

        public string GetBoricaACDescription(string actionCode)
        {
            Dictionary<string, string> actionCodes = new Dictionary<string, string>();
            actionCodes.Add("1", "Дублирана трансакция");
            actionCodes.Add("2", "Отказана трансакция");
            actionCodes.Add("3", "Грешка при обработка на трансакцията");

            return actionCodes.ContainsKey(actionCode) ? actionCodes[actionCode] : "Грешка при трансакцията";
        }

        public string GetPSignData()
        {
            return AddNextKeyToPSign(Action.ToString()) +
                AddNextKeyToPSign(Rc) +
                AddNextKeyToPSign(Approval) +
                AddNextKeyToPSign(Terminal) +
                AddNextKeyToPSign(((int)TrType).ToString()) +
                AddNextKeyToPSign(AmountValue == null ? null : ((decimal)AmountValue).ToString("#0.00", System.Globalization.CultureInfo.InvariantCulture)) +
                AddNextKeyToPSign(Currency) +
                AddNextKeyToPSign(Order) +
                AddNextKeyToPSign(Rrn) +
                AddNextKeyToPSign(INT_REF) +
                AddNextKeyToPSign(Pares_Status) +
                AddNextKeyToPSign(Eci) +
                AddNextKeyToPSign(TimeStamp) +
                AddNextKeyToPSign(Nonce) +
                AddNextKeyToPSign("-");//"-" is for the RFU (Reserved for Future use) field which is not used and implemented currently
        }

        public BoricaTransaction UpdateBoricaTransactionSuccessPayment(BoricaTransaction boricaTransaction)
        {
            if (boricaTransaction == null)
            {
                return null;
            }

            boricaTransaction.TransactionStatusId = (int)BoricaTransactionStatusEnum.Paid;

            boricaTransaction = this.UpdateBoricaTransaction(boricaTransaction);
            boricaTransaction.StatusMessage = GetBoricaRCDescription(boricaTransaction.Rc);

            return boricaTransaction;
        }

        public BoricaTransaction UpdateBoricaTransactionPendingPayment(BoricaTransaction boricaTransaction)
        {
            if (boricaTransaction == null)
            {
                return null;
            }

            boricaTransaction.TransactionStatusId = (int)BoricaTransactionStatusEnum.Pending;

            boricaTransaction = this.UpdateBoricaTransaction(boricaTransaction);
            boricaTransaction.StatusMessage = GetBoricaRCDescription(boricaTransaction.Rc);

            return boricaTransaction;
        }

        public BoricaTransaction UpdateBoricaTransactionCanceledPayment(BoricaTransaction boricaTransaction)
        {
            if (boricaTransaction == null)
            {
                return null;
            }

            boricaTransaction.TransactionStatusId = (int)BoricaTransactionStatusEnum.Canceled;

            boricaTransaction = this.UpdateBoricaTransaction(boricaTransaction);

            if (boricaTransaction.Rc != "00")
                boricaTransaction.StatusMessage = GetBoricaRCDescription(boricaTransaction.Rc);
            else
                boricaTransaction.StatusMessage = GetBoricaACDescription(boricaTransaction.Action);

            return boricaTransaction;
        }

        private BoricaTransaction UpdateBoricaTransaction(BoricaTransaction boricaTransaction)
        {
            boricaTransaction.Terminal = this.Terminal;
            boricaTransaction.Action = this.Action.ToString();
            boricaTransaction.Rc = this.Rc;
            boricaTransaction.Approval = this.Approval;
            boricaTransaction.Rrn = this.Rrn;
            boricaTransaction.IntRef = this.INT_REF;
            boricaTransaction.StatusMessage = this.StatusMsg;
            boricaTransaction.Card = this.Card;
            boricaTransaction.Eci = this.Eci;
            boricaTransaction.ParesStatus = this.Pares_Status;

            return boricaTransaction;
        }
    }
}
