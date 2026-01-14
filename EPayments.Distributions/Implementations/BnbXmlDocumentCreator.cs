using EPayments.Distributions.Interfaces;
using EPayments.Distributions.Models.BNB;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

namespace EPayments.Distributions.Implementations
{
    public class BnbXmlDocumentCreator : IBnbXmlDocumentCreator
    {
        public XDocument CreateDocument(BnbFile bnbFile)
        {
            XDocument document = bnbFile.CreateXmlDocument();
            
            return document;
        }

        public XDocument SignDocument(XDocument document, RSA rsa)
        {
            XmlDocument xmlDocument = this.ConvertToXmlDocument(document);

            Reference reference = new Reference(string.Empty);
            XmlDsigEnvelopedSignatureTransform xmlDsigEnvelopedSignatureTransform = new XmlDsigEnvelopedSignatureTransform();
            reference.AddTransform(xmlDsigEnvelopedSignatureTransform);

            SignedXml signedXml = new SignedXml(xmlDocument) { SigningKey = rsa };
            signedXml.AddReference(reference);
            signedXml.ComputeSignature();

            XmlElement xmlElement = signedXml.GetXml();
            xmlDocument.DocumentElement.AppendChild(xmlDocument.ImportNode(xmlElement, true));

            return ConvertToXDocument(xmlDocument);
        }

        public List<string> ValidateDocument(XDocument document, 
            string xsdDirectoryPath, 
            string xsdName)
        {
            XmlSchemaSet xmlSchemaSet = new XmlSchemaSet();
            xmlSchemaSet.Add(string.Empty, Path.Combine(xsdDirectoryPath, xsdName));
            List<string> errors = new List<string>();

            document.Validate(xmlSchemaSet, (o, ev) =>
            {
                errors.Add(ev.Message);
            });

            return errors;
        }

        private XmlDocument ConvertToXmlDocument(XDocument xDocument)
        {
            XmlDocument xmlDocument = new XmlDocument() { PreserveWhitespace = true };
            using (XmlReader xmlReader = xDocument.CreateReader())
            {
                xmlDocument.Load(xmlReader);
                return xmlDocument;
            }
        }

        private XDocument ConvertToXDocument(XmlDocument xmlDocument)
        {
            using (XmlNodeReader nodeReader = new XmlNodeReader(xmlDocument))
            {
                nodeReader.MoveToContent();

                return XDocument.Load(nodeReader);
            }
        }
    }
}
