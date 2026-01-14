using EPayments.Common.Helpers;
using EPayments.Distributions.Enums;
using System;
using System.Globalization;
using System.Xml.Linq;

namespace EPayments.Distributions.Models.BNB
{
    /// <summary>
    /// d document
    /// </summary>
    public class BnbDocument
    {
        /// <summary>
        /// Document type, one of this options ПНКП, ПНВБ, БПН
        /// </summary>
        public DocumentEnum Doc { get; set; }

        /// <summary>
        /// Unique number of the document in the BnbAccount
        /// </summary>
        public string Nok { get; set; }

        /// <summary>
        /// ipol, name of the receiver
        /// </summary>
        public string IPol { get; set; }

        /// <summary>
        /// IBAN of the receiver
        /// </summary>
        public string Iban { get; set; } 

        /// <summary>
        /// BIC of the receiver
        /// </summary>
        public string Bic { get; set; }

        /// <summary>
        /// Payment type
        /// </summary>
        public string Vpp { get; set; }

        /// <summary>
        /// Currency description
        /// </summary>
        public string Cur { get; set; }

        /// <summary>
        /// Sum
        /// </summary>
        public decimal Su { get; set; }

        /// <summary>
        /// First description of the payment
        /// </summary>
        public string O1 { get; set; }

        /// <summary>
        /// Second description of the payment
        /// </summary>
        public string O2 { get; set; }

        public PaymentSystemEnum? Sys { get; set; }

        /// <summary>
        /// Tax description
        /// </summary>
        public string Tax { get; set; }

        /// <summary>
        /// Code of action
        /// </summary>
        public string KD { get; set; }

        /// <summary>
        /// Date of execution
        /// </summary>
        public DateTime? Dex { get; set; }

        /// <summary>
        /// Payment type of the sender
        /// </summary>
        public string Vpn { get; set; }

        public BnbBudget BnbBudget { get; set; }

        public XElement Create()
        {
            char[] notAllowedSymbols = { '-', ':' };
            
            string recieverName = this.IPol.TrimStart(notAllowedSymbols).Replace("\"", "").Replace(" - ", " ").Replace(" : ", " ");
            string o1 = this.O1.Trim(notAllowedSymbols);
            string o2 = this.O2.Trim(notAllowedSymbols);

            if (recieverName.Length > 30)
            {
                recieverName = recieverName.Substring(0, 30);
            }
            
            XElement xElement = new XElement("d",
                new XAttribute("doc", this.Doc.GetDescription()),
                new XAttribute("nok", this.Nok),
                new XAttribute("ipol", recieverName),
                new XAttribute("iban", this.Iban),
                new XAttribute("bic", this.Bic),
                new XAttribute("cur", this.Cur),
                new XAttribute("su", this.Su.ToString("#0.##", CultureInfo.InvariantCulture)),
                new XAttribute("o1", o1),
                new XAttribute("o2", o2)
                );

            if (!string.IsNullOrWhiteSpace(this.Vpp))
            {
                xElement.Add(new XAttribute("vpp", this.Vpp));
            }

            if (this.Sys != null)
            {
                xElement.Add(new XAttribute("sys", this.Sys.GetDescription()));
            }

            if (!string.IsNullOrWhiteSpace(this.Tax))
            {
                xElement.Add(new XAttribute("tax", this.Tax));
            }

            if (!string.IsNullOrWhiteSpace(this.KD))
            {
                xElement.Add(new XAttribute("kd", this.KD));
            }

            if (this.Dex != null)
            {
                xElement.Add(new XAttribute("dex", ((DateTime)this.Dex).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)));
            }

            if (!string.IsNullOrWhiteSpace(this.Vpn))
            {
                xElement.Add(new XAttribute("vpn", this.Vpn));
            }

            if (BnbBudget != null)
            {
                xElement.Add(BnbBudget.Create());
            }

            return xElement;
        }
    }
}
