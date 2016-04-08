using System;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;


namespace ExcelTextReplace
{
    public class JsonTools
    {
        // 从一个对象信息生成Json串
        public static string ObjectToJson(object obj)
        {
            return JsonConvert.SerializeObject(obj);
        }
        // 从一个Json串生成对象信息
        public static object JsonToObject(string jsonString, object obj)
        {
           
            return JsonConvert.DeserializeObject(jsonString, obj.GetType());
        }

        public static object JsonToObject(string jsonString)
        {
            return JsonConvert.DeserializeObject(jsonString);
        }

        public static Dictionary<string, string> JsonToDictionary(string jsonString)
        {
            Dictionary<string, string> dict =
                JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);

            return dict;
        }

        public static object JsonTest(string jsonString)
        {
            Dictionary<string, string> dict = 
                JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);

            foreach (var item in dict)
            {
                Console.WriteLine("key:{0},vlaue:{1}",item.Key ,item.Value);
            }


            return null;
        }

        public static List<Dictionary<string,string>> JsonToList(string jsonString)
        {
           // JToken entireJson = JToken.Parse(jsonString);
            //JArray inner = JArray.Parse(entireJson);

                //entireJson["questionAnswer"].Value<JArray>();
            //JArray inner = entireJson["questionAnswer"].Value<JArray>();
            List<Dictionary<string, string>> videogames = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(jsonString);
            /*foreach (var dict in videogames)
            {
                foreach (var item in dict)
                {
                    Console.WriteLine("key:{0},vlaue:{1}", item.Key, item.Value);
                }
            }*/
            return videogames;
        }
    }
}
