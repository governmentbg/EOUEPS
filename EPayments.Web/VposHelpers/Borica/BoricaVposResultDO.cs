using EPayments.Data.ViewObjects.Web.APGModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.VposHelpers.Borica
{
    public class BoricaVposResultDO
    {
        public string TransactionType { get; set; }
        public string DateTime { get; set; }
        public string Amount { get; set; }
        public string TerminalId { get; set; }
        public string RequestIdentifier { get; set; }
        public string ResultCode { get; set; }
        public string ProtocolVersion { get; set; }

        public BoricaVposResultDO(string eBoricaData)
        {
            this.TransactionType = eBoricaData.Substring(0, 2);
            this.DateTime = eBoricaData.Substring(2, 14);
            this.Amount= eBoricaData.Substring(16, 12);
            this.TerminalId = eBoricaData.Substring(28, 8);
            this.RequestIdentifier = eBoricaData.Substring(36, 15);
            this.ResultCode = eBoricaData.Substring(51, 2);
            this.ProtocolVersion = eBoricaData.Substring(53, 3);
        }

        public BoricaVposResultDO(APGWPaymentResponseDataDO apgwResponse)
        {
            this.TransactionType = apgwResponse.TrType.ToString();
            this.DateTime = apgwResponse.TimeStamp;
            this.Amount = apgwResponse.Amount;
            this.TerminalId = apgwResponse.Terminal;
            this.RequestIdentifier = apgwResponse.Order;
            this.ResultCode = apgwResponse.Rc;
            this.ProtocolVersion = null;
        }
    }
}