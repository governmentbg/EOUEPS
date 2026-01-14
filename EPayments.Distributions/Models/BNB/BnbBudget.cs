using System;
using System.Globalization;
using System.Xml.Linq;

namespace EPayments.Distributions.Models.BNB
{
    public class BnbBudget
    {
        /// <summary>
        /// Document type
        /// </summary>
        public string Vd { get; set; }

        /// <summary>
        /// Document number
        /// </summary>
        public string Nd { get; set; }

        /// <summary>
        /// Date of the document
        /// </summary>
        public DateTime? Dd { get; set; }

        /// <summary>
        /// Start date of the period
        /// </summary>
        public DateTime? Db { get; set; }

        /// <summary>
        /// End date of the period
        /// </summary>
        public DateTime? De { get; set; }

        /// <summary>
        /// Name of the receiver
        /// </summary>
        public string Izl { get; set; }

        /// <summary>
        /// Bulstat
        /// </summary>
        public string Bul { get; set; }

        /// <summary>
        /// EGN
        /// </summary>
        public string Egn { get; set; }

        /// <summary>
        /// Person card - LNC, for foreigners
        /// </summary>
        public string Lnc { get; set; }

        public XElement Create()
        {
            string dateFormat = "yyyy-MM-dd";
            
            XElement xElement = new XElement("bud");

            //if (!string.IsNullOrEmpty(this.Vd))
            //{
            //    xElement.Add(new XAttribute("vd", this.Vd));
            //}

            if (!string.IsNullOrEmpty(this.Nd))
            {
                xElement.Add(new XAttribute("nd", this.Nd));
            }

            if (this.Dd != null)
            {
                xElement.Add(new XAttribute("dd", ((DateTime)this.Dd).ToString(dateFormat, CultureInfo.InvariantCulture)));
            }

            //if (this.Db != null)
            //{
            //    xElement.Add(new XAttribute("db", ((DateTime)this.Db).ToString(dateFormat, CultureInfo.InvariantCulture)));
            //}

            //if (this.De != null)
            //{
            //    xElement.Add(new XAttribute("de", ((DateTime)this.De).ToString(dateFormat, CultureInfo.InvariantCulture)));
            //}

            if (!string.IsNullOrEmpty(this.Izl))
            {
                xElement.Add(new XAttribute("izl", this.Izl));
            }

            if (!string.IsNullOrEmpty(this.Bul))
            {
                xElement.Add(new XAttribute("bul", this.Bul));
            }

            if (!string.IsNullOrEmpty(this.Egn))
            {
                xElement.Add(new XAttribute("egn", this.Egn));
            }

            if (!string.IsNullOrEmpty(this.Lnc))
            {
                xElement.Add(new XAttribute("lnc", this.Lnc));
            }

            return xElement;
        }
    }
}
