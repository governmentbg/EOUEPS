using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Schema;

namespace EPayments.Service.Api.Common.XsdValidator
{
    public interface ISchemaValidator
    {
        List<string> ValidateXmlSchema(string xml);
    }
}
