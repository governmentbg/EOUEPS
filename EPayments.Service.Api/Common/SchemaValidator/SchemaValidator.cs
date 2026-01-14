using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Schema;
using log4net;

namespace EPayments.Service.Api.Common.XsdValidator
{
    internal class SchemaValidator : ISchemaValidator
    {
        private object _syncRoot = new object();
        private static volatile XmlSchemaSet _xmlSchemaSet = null;
        private static string _schemasDirectoryPath = null;
        private readonly ILog logger = LogManager.GetLogger(typeof(SchemaValidator));

        private static string SchemasDirecotryPath
        {
            get
            {
                if (_schemasDirectoryPath == null)
                {
                    _schemasDirectoryPath = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "bin", "schemas");
                }

                return _schemasDirectoryPath;
            }
        }

        public List<string> ValidateXmlSchema(string xml)
        {
            try
            {
                this.logger.Info($"1. BEFORE LOCK-BLOCK - validator: {this.GetHashCode()}; locker: {_syncRoot.GetHashCode()};");
                //validating over the same XmlSchemaSet is NOT THREADSAFE
                lock (_syncRoot)
                {
                    this.logger.Info($"2. ENTERED LOCK - validator: {this.GetHashCode()}; locker: {_syncRoot.GetHashCode()};");
                    if (_xmlSchemaSet == null)
                    {
                        this.logger.Info($"3.1. LOADING ALL SCHEMAS - validator: {this.GetHashCode()}; locker: {_syncRoot.GetHashCode()};");
                        _xmlSchemaSet = new XmlSchemaSet();
                        foreach (string schemaFile in Directory.GetFiles(SchemasDirecotryPath, "*.xsd", SearchOption.AllDirectories))
                        {
                            using (TextReader schemaReader = new StreamReader(schemaFile))
                            {
                                _xmlSchemaSet.Add(null, XmlReader.Create(schemaReader));
                            }
                        }
                        this.logger.Info($"3.2. LOADED {_xmlSchemaSet.Count} SCHEMAS - validator: {this.GetHashCode()}; locker: {_syncRoot.GetHashCode()};");
                    }

                    XmlReader reader = null;
                    List<Tuple<string, string>> errors = new List<Tuple<string, string>>();
#if DEBUG
                    List<Tuple<string, string>> warnings = new List<Tuple<string, string>>();
#endif

                    XmlReaderSettings settings = new XmlReaderSettings();
                    settings.ValidationType = ValidationType.Schema;
                    settings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
                    settings.ValidationEventHandler +=
                        delegate(object sender, ValidationEventArgs args)
                        {
                            if (args.Severity == XmlSeverityType.Error)
                            {
                                errors.Add(Tuple.Create(args.Message, reader.Name));
                            }
#if DEBUG
                            else
                            {
                                warnings.Add(Tuple.Create(args.Message, reader.Name));
                            }
#endif
                        };
                    settings.Schemas = _xmlSchemaSet;

                    this.logger.Info($"4. PARSING LOADED SCHEMAS ({settings.Schemas.Count}) - validator: {this.GetHashCode()}; locker: {_syncRoot.GetHashCode()};");
                    using (StringReader sr = new StringReader(xml))
                    using (reader = XmlReader.Create(sr, settings))
                    {
                        while (reader.Read())
                        {
                        }
                    }

                    return errors.Select(e => string.Format("{0}{1}", e.Item1, !String.IsNullOrWhiteSpace(e.Item2) ? String.Format(" (Element: {0})", e.Item2) : String.Empty)).ToList();
                }
            }
            catch (Exception ex)
            {
                this.logger.Info($"5. CAUGHT EXCEPTION - validator: {this.GetHashCode()}; locker: {_syncRoot.GetHashCode()};");
                return new List<string>
                {
                    ex.Message
                };
            }
        }
    }
}
