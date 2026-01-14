using System;

namespace EPayments.Documents.Serializer
{
    /// <summary>
    /// Имплементация на компонент извършващ сериализацията на документи
    /// </summary>
    public class DocumentSerializer : IDocumentSerializer
    {
        #region Public

        /// <summary>
        /// Конструктор на имплементацията
        /// </summary>
        public DocumentSerializer()
        {
        }

        /// <summary>
        /// Сериализира документа като стринг
        /// </summary>
        /// <typeparam name="T">тип на документа</typeparam>
        /// <param name="document">документ</param>
        /// <returns>стринг представляващ сериализирания документ</returns>
        public string XmlSerializeToString<T>(T document)
        {
            return XmlSerializeToStringInternal(typeof(T), document);
        }

        /// <summary>
        /// Сериализира документа като стринг
        /// </summary>
        /// <param name="document">документ</param>
        /// <returns>стринг представляващ сериализирания документ</returns>
        public string XmlSerializeObjectToString(object document)
        {
            return XmlSerializeToStringInternal(document.GetType(), document);
        }

        /// <summary>
        /// Сериализира документа като масив от байтове
        /// </summary>
        /// <typeparam name="T">тип на документа</typeparam>
        /// <param name="document">документ</param>
        /// <returns>масив от байтове представляващ сериализирания документ</returns>
        public byte[] XmlSerializeToBytes<T>(T document)
        {
            return XmlSerializeToBytesInternal(typeof(T), document);
        }

        /// <summary>
        /// Сериализира документа като масив от байтове
        /// </summary>
        /// <param name="document">документ</param>
        /// <returns>масив от байтове представляващ сериализирания документ</returns>
        public byte[] XmlSerializeObjectToBytes(object document)
        {
            return XmlSerializeToBytesInternal(document.GetType(), document);
        }

        /// <summary>
        /// Десериализира документа от стринг
        /// </summary>
        /// <typeparam name="T">тип на документа</typeparam>
        /// <param name="documentXml">стрингова репрезентация на XML на документа</param>
        /// <returns>десериализирания документ</returns>
        public T XmlDeserializeFromString<T>(string documentXml)
        {
            if (string.IsNullOrWhiteSpace(documentXml))
                throw new ArgumentNullException("documentXml should not be empty.");

            return (T)XmlDeserializeFromStringInternal(typeof(T), documentXml);
        }

        /// <summary>
        /// Десериализира документа от стринг
        /// </summary>
        /// <param name="objectType">тип на документа</param>
        /// <param name="documentXml">стрингова репрезентация на XML на документа</param>
        /// <returns>десериализирания документ</returns>
        public object XmlDeserializeFromString(Type objectType, string documentXml)
        {
            return XmlDeserializeFromStringInternal(objectType, documentXml);
        }

        /// <summary>
        /// Десериализира документа от масив от байтове
        /// </summary>
        /// <typeparam name="T">тип на документа</typeparam>
        /// <param name="bytes">битова репрезентация на XML на документа</param>
        /// <returns>десериализирания документ</returns>
        public T XmlDeserializeFromBytes<T>(byte[] bytes)
        {
            return (T)XmlDeserializeFromBytesInternal(typeof(T), bytes);
        }

        /// <summary>
        /// Десериализира документа от масив от байтове
        /// </summary>
        /// <param name="objectType">тип на документа</param>
        /// <param name="bytes">битова репрезентация на XML на документа</param>
        /// <returns>десериализирания документ</returns>
        public object XmlDeserializeFromBytes(Type objectType, byte[] bytes)
        {
            return XmlDeserializeFromBytesInternal(objectType, bytes);
        }

        #endregion

        #region Private

        private byte[] XmlSerializeToBytesInternal(Type type, object document)
        {
            if (document == null)
                throw new ArgumentNullException("document should not be null.");

            using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
            {
                using (System.Xml.XmlTextWriter xmlWriter = new System.Xml.XmlTextWriter(ms, DocumentSerializerSettings.DefaultEncoding))
                {
                    System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(type);
                    serializer.Serialize(xmlWriter, document);
                }

                return ms.ToArray();
            }
        }

        private string XmlSerializeToStringInternal(Type type, object document)
        {
            byte[] bytes = XmlSerializeToBytesInternal(type, document);

            using (System.IO.MemoryStream ms = new System.IO.MemoryStream(bytes))
            using (System.IO.StreamReader sr = new System.IO.StreamReader(ms, DocumentSerializerSettings.DefaultEncoding))
            {
                return sr.ReadToEnd();
            }
        }

        private object XmlDeserializeFromStringInternal(Type type, string xml)
        {
            using (System.IO.StringReader sr = new System.IO.StringReader(xml))
            {
                System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(type);
                return serializer.Deserialize(sr);
            }
        }

        private object XmlDeserializeFromBytesInternal(Type type, byte[] bytes)
        {
            using (System.IO.Stream s = new System.IO.MemoryStream(bytes))
            {
                System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(type);
                return serializer.Deserialize(s);
            }
        }

        #endregion
    }
}
