using EPayments.Distributions.Models.BNB;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Xml.Linq;

namespace EPayments.Distributions.Interfaces
{
    public interface IBnbXmlDocumentCreator
    {
        XDocument CreateDocument(BnbFile bnbFile);

        List<string> ValidateDocument(XDocument document,
            string xsdDirectoryPath,
            string xsdName);

        XDocument SignDocument(XDocument document, RSA rsa);
    }
}
