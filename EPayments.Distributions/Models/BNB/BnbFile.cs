using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace EPayments.Distributions.Models.BNB
{
    public class BnbFile
    {
        public BnbHeader Header { get; set; }

        public List<BnbAccount> Accounts { get; set; }

        public XDocument CreateXmlDocument()
        {
            XElement root = new XElement("f");

            if (this.Header != null)
            {
                root.Add(this.Header.Create());
            }

            if (this.Accounts != null && this.Accounts.Count > 0)
            {
                root.Add(Accounts.Select(a => a.Create()));
            }

            XDocument document = new XDocument(new XDeclaration("1.0", "WINDOWS-1251", null), root);

            return document;
        }
    }
}
