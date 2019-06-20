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
            return JsonConvert.SerializeObject(obj) ;
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

    }
}
