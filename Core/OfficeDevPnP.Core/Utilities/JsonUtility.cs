using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Utilities
{
    /// <summary>
    /// Utility class that supports the serialization from Json to type and vice versa
    /// </summary>
    public static class JsonUtility
    {
        /// <summary>
        /// Serializes an object of type T to a json string
        /// </summary>
        /// <typeparam name="T">Type of obj</typeparam>
        /// <param name="obj">Object to serialize</param>
        /// <returns>json string</returns>
        public static string Serialize<T>(T obj)
        {
            return JsonConvert.SerializeObject(obj);
        }

        /// <summary>
        /// Deserializes a json string to an object of type T
        /// </summary>
        /// <typeparam name="T">Type of the returned object</typeparam>
        /// <param name="json">json string</param>
        /// <returns>Object of type T</returns>
        public static T Deserialize<T>(string json)
        {
            return JsonConvert.DeserializeObject<T>(json);
        }

        /// <summary>
        /// Try to deserialize a string into an object
        /// </summary>
        /// <typeparam name="T">Type of the returned object</typeparam>
        /// <param name="json">json string</param>
        /// <param name="value">Object of type T is the deserialization succeeded</param>
        /// <returns><c>true</c> if the deserialization was successfull otherwise <c>false</c></returns>
        public static bool TryDeserialize<T>(string json, out T value)
        {
            try
            {
                value = Deserialize<T>(json);
                return true;
            }
            catch (JsonReaderException)
            {
                value = default(T);
                return false;
            }
            catch (JsonSerializationException)
            {
                value = default(T);
                return false;
            }
        }
    }
}