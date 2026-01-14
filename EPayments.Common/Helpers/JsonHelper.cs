using System;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace EPayments.Common.Helpers
{
    public static class JsonHelper
    {
        /// <summary>
        /// Check if a string can be parsed to a json object.
        /// </summary>
        /// <param name="str"></param>
        /// <returns>True if it's valid and False if not</returns>
        public static bool IsValidJson(string str)
        {
            if (string.IsNullOrWhiteSpace(str))
                return false;

            str = str.Trim();

            if (!(str.StartsWith("{") && str.EndsWith("}")))
                return false;

            try
            {
                var obj = JToken.Parse(str);
                return true;
            }
            catch (JsonReaderException)
            {
                return false;
            }
        }

        /// <summary>
        /// Attempts to repair a JSON string by escaping unescaped double quotes
        /// specifically within the "PropertyAddress" field's value.
        /// Uses efficient string operations and StringBuilder, avoiding Regex.
        /// </summary>
        /// <param name="jsonString">The potentially broken JSON string.</param>
        /// <returns>The repaired JSON string, or the original string if no repair was needed or possible.</returns>
        public static string RepairBrokenAdditionalInformationJsonString(string jsonString)
        {
            if (string.IsNullOrWhiteSpace(jsonString))
                return jsonString;

            var json = jsonString.Trim();

            if (!json.StartsWith("{") || !json.EndsWith("}"))
            {
                return jsonString;
            }

            const string propertyAddressKey = "\"propertyAddress\":\"";
            int keyIndex = json.IndexOf(propertyAddressKey, StringComparison.Ordinal);

            if (keyIndex == -1)
            {
                return jsonString;
            }

            int valueStartIndex = keyIndex + propertyAddressKey.Length;

            int valueEndIndex = -1;
            bool previousCharWasBackslash = false;

            for (int i = valueStartIndex; i < json.Length; i++)
            {
                char currentChar = json[i];
                if (currentChar == '"' && !previousCharWasBackslash)
                {
                    int nextCharIndex = i + 1;
                    while (nextCharIndex < json.Length && char.IsWhiteSpace(json[nextCharIndex]))
                    {
                        nextCharIndex++;
                    }

                    if (nextCharIndex < json.Length && (json[nextCharIndex] == ',' || json[nextCharIndex] == '}'))
                    {
                        valueEndIndex = i;
                        break;
                    }
                }
                previousCharWasBackslash = currentChar == '\\' && !previousCharWasBackslash;
            }

            if (valueEndIndex == -1)
            {
                return jsonString;
            }

            StringBuilder repairedJson = new StringBuilder(json.Length + 20);
            repairedJson.Append(json, 0, valueStartIndex);

            previousCharWasBackslash = false;
            bool repairNeeded = false;

            for (int i = valueStartIndex; i < valueEndIndex; i++)
            {
                char currentChar = json[i];
                if (currentChar == '"' && !previousCharWasBackslash)
                {
                    repairedJson.Append('\\');
                    repairedJson.Append('"');
                    repairNeeded = true;
                }
                else
                {
                    repairedJson.Append(currentChar);
                }

                previousCharWasBackslash = currentChar == '\\' && !previousCharWasBackslash;
            }

            repairedJson.Append(json, valueEndIndex, json.Length - valueEndIndex);

            if (repairNeeded)
            {
                return repairedJson.ToString();
            }
            else
            {
                return jsonString;
            }
        }
    }
}