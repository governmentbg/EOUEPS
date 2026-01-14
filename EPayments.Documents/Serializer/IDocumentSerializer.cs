using System;

namespace EPayments.Documents.Serializer
{
    /// <summary>
    /// Интерфейс на компонент извършващ сериализацията на документи
    /// </summary>
    public interface IDocumentSerializer
    {
        /// <summary>
        /// Сериализира документа като стринг
        /// </summary>
        /// <typeparam name="T">тип на документа</typeparam>
        /// <param name="document">документ</param>
        /// <returns>стринг представляващ сериализирания документ</returns>
        string XmlSerializeToString<T>(T document);

        /// <summary>
        /// Сериализира документа като стринг
        /// </summary>
        /// <param name="document">документ</param>
        /// <returns>стринг представляващ сериализирания документ</returns>
        string XmlSerializeObjectToString(object document);

        /// <summary>
        /// Сериализира документа като масив от байтове
        /// </summary>
        /// <typeparam name="T">тип на документа</typeparam>
        /// <param name="document">документ</param>
        /// <returns>масив от байтове представляващ сериализирания документ</returns>
        byte[] XmlSerializeToBytes<T>(T document);

        /// <summary>
        /// Сериализира документа като масив от байтове
        /// </summary>
        /// <param name="document">документ</param>
        /// <returns>масив от байтове представляващ сериализирания документ</returns>
        byte[] XmlSerializeObjectToBytes(object document);

        /// <summary>
        /// Десериализира документа от стринг
        /// </summary>
        /// <typeparam name="T">тип на документа</typeparam>
        /// <param name="documentXml">стрингова репрезентация на XML на документа</param>
        /// <returns>десериализирания документ</returns>
        T XmlDeserializeFromString<T>(string documentXml);

        /// <summary>
        /// Десериализира документа от стринг
        /// </summary>
        /// <param name="objectType">тип на документа</param>
        /// <param name="documentXml">стрингова репрезентация на XML на документа</param>
        /// <returns>десериализирания документ</returns>
        object XmlDeserializeFromString(Type objectType, string documentXml);

        /// <summary>
        /// Десериализира документа от масив от байтове
        /// </summary>
        /// <typeparam name="T">тип на документа</typeparam>
        /// <param name="bytes">битова репрезентация на XML на документа</param>
        /// <returns>десериализирания документ</returns>
        T XmlDeserializeFromBytes<T>(byte[] bytes);

        /// <summary>
        /// Десериализира документа от масив от байтове
        /// </summary>
        /// <param name="objectType">тип на документа</param>
        /// <param name="bytes">битова репрезентация на XML на документа</param>
        /// <returns>десериализирания документ</returns>
        object XmlDeserializeFromBytes(Type objectType, byte[] bytes);
    }
}
