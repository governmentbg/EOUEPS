using EPayments.Common.Helpers;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace EPayments.Distributions.Models.BNB
{
    /// <summary>
    /// a document
    /// </summary>
    public class BnbAccount
    {
        /// <summary>
        /// Bank account of the sender
        /// </summary>
        public string Acc { get; set; }

        /// <summary>
        /// Bic number of the sender
        /// </summary>
        public string Bic { get; set; }

        /// <summary>
        /// Type of the currency
        /// </summary>
        public string Cur { get; set; }

        /// <summary>
        /// Total sum of all BnbDocuments
        /// </summary>
        public decimal Do { get; set; }

        public List<BnbDocument> Documents { get; set; }

        public XElement Create()
        {
            XElement xElement = new XElement("a",
                new XAttribute("acc", this.Acc),
                new XAttribute("bic", this.Bic),
                new XAttribute("cur", this.Cur),
                new XAttribute("do", this.Do.ToString("#0.##", CultureInfo.InvariantCulture))
                );

            if (this.Documents != null && this.Documents.Count > 0)
            {
                xElement.Add(this.Documents.Select(d => d.Create()));
            }

            return xElement;
        }
    }
}
