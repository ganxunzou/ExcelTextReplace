using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelTextReplace
{
    public class JSONTools
    {
        private const string PLACEHOLDER_DATA = "PlaceHolderData";
        private const string ROW_DATA = "RowData";
        private const string LOCATION_DATA = "LocationData";
        private const string SHEETINDEX = "SheetIndex";

        public static List<ReplaceDataVO> parseJsonStrToVos(string content)
        {

            List<ReplaceDataVO> rpDataVoLs = new List<ReplaceDataVO>();

            JArray jarr = (JArray)JsonConvert.DeserializeObject(content);

            ReplaceDataVO rpDataVo = null;

            foreach (JObject item in jarr)
            {
                rpDataVo = new ReplaceDataVO();
                rpDataVo.PlaceHolderData = getPlaceHolderLs((JObject)item[PLACEHOLDER_DATA]);
                rpDataVo.RowData = getRowDataVoLs((JArray)item[ROW_DATA]);
                rpDataVo.LocationData = getLocationDataVoLs((JArray)item[LOCATION_DATA]);
                rpDataVo.SheetIndex = int.Parse((string)item[SHEETINDEX]);

                rpDataVoLs.Add(rpDataVo);
            }
            return rpDataVoLs;
        }

        private static List<PlaceHolderVO> getPlaceHolderLs( JObject item)
        {
            List<PlaceHolderVO> phVoLs = new List<PlaceHolderVO>();

            PlaceHolderVO phVo = null;

            foreach (JProperty prop in item.Properties())
            {
                phVo = new PlaceHolderVO(); 
                phVo.Holder = prop.Name;
                phVo.Value = prop.Value.ToString();

                phVoLs.Add(phVo);
            }

            return phVoLs;
        }

        private static List<RowDataVO> getRowDataVoLs(JArray jarr)
        {
            List<RowDataVO> rowDataLs = new List<RowDataVO>();

            RowDataVO rowData = null;
            foreach (JObject item in jarr)
            {
                rowData = new RowDataVO();
                rowData.From = int.Parse((string)item["from"]);
                rowData.Rows = int.Parse((string)item["rows"]);

                rowDataLs.Add(rowData);
            }

            return rowDataLs;
        }

        private static List<LocationDataVO> getLocationDataVoLs(JArray jarr)
        {
            List<LocationDataVO> ltDataLs = new List<LocationDataVO>();

            LocationDataVO ltData = null;
            foreach (JObject item in jarr)
            {
                ltData = new LocationDataVO();
                ltData.Localtion = (string)item["location"];
                ltData.Value = (string)item["value"];

                ltDataLs.Add(ltData);
            }

            return ltDataLs;
        }

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
                Console.WriteLine("key:{0},vlaue:{1}", item.Key, item.Value);
            }


            return null;
        }

        public static List<Dictionary<string, string>> JsonToList(string jsonString)
        {
            // JToken entireJson = JToken.Parse(jsonString);
            //JArray inner = JArray.Parse(entireJson);

            //entireJson["questionAnswer"].Value<JArray>();
            //JArray inner = entireJson["questionAnswer"].Value<JArray>();
            List<Dictionary<string, string>> dict = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(jsonString);
            /*foreach (var dict in videogames)
            {
                foreach (var item in dict)
                {
                    Console.WriteLine("key:{0},vlaue:{1}", item.Key, item.Value);
                }
            }*/
            return dict;
        }
    }
}
