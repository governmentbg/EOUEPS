using System;
using System.Xml.Linq;

namespace EPayments.Distributions.Models.BNB
{
    /// <summary>
    /// h document
    /// </summary>
    public class BnbHeader
    {
        /// <summary>
        /// Unique number -> Id of Distribution
        /// </summary>
        public string RefId { get; set; } 

        /// <summary>
        /// DateTime of creation
        /// </summary>
        public DateTime TimeStamp { get; set; } 

        /// <summary>
        /// Bulstat of the sender
        /// </summary>
        public string Sender { get; set; }

        /// <summary>
        /// Name of the sender
        /// </summary>
        public string SenderName { get; set; }

        /// <summary>
        /// Reciever, must be constant
        /// </summary>
        public string Receiver => "BNBOnline";

        public XElement Create()
        {
            XElement element = new XElement("h",
                new XAttribute("refid", this.RefId),
                new XAttribute("timestamp", this.TimeStamp.ToString("yyyy-MM-ddTHH:mm:ss")),
                new XAttribute("sender", this.Sender),
                new XAttribute("sendername", this.SenderName),
                new XAttribute("receiver", this.Receiver));

            return element;
        }
    }
}
