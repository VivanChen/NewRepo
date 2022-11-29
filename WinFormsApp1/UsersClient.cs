using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using WinFormsApp1.Model;
using static WinFormsApp1.Model.Sightseeing;
using static WinFormsApp1.Model.Weather;

namespace WinFormsApp1
{
    public class UsersClient
    {
        private HttpClient client;

        public UsersClient()
        {
            client = new HttpClient();
            client.Timeout = TimeSpan.FromMinutes(200);
        }


        public async Task<string> BatchPostAsync(string url, object data)
        {
            try
            {
                var buffer = System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(data));

                var byteContent = new ByteArrayContent(buffer);

                byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");

                var response = await client.PostAsync(url, byteContent).ConfigureAwait(false);

                var Result = await response.Content.ReadAsStringAsync();

                return Result;
            }
            catch (Exception)
            {
                throw;
            }

        }
        public Weather.Rootobject getJson(string uri)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(uri); //request請求
            req.Timeout = 10000; //request逾時時間
            req.Method = "GET"; //request方式
            HttpWebResponse respone = (HttpWebResponse)req.GetResponse(); //接收respone
            StreamReader streamReader = new StreamReader(respone.GetResponseStream(), Encoding.UTF8); //讀取respone資料
            string result = streamReader.ReadToEnd(); //讀取到最後一行
            respone.Close();
            streamReader.Close();
            Weather.Rootobject jsondata = JsonConvert.DeserializeObject<Weather.Rootobject>(result); //將資料轉為json物件
            return jsondata;

        }
        public Sightseeing.Rootobject GetSightseeing_json(string uri)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(uri); //request請求
            req.Timeout = 10000; //request逾時時間
            req.Method = "GET"; //request方式
            HttpWebResponse respone = (HttpWebResponse)req.GetResponse(); //接收respone
            StreamReader streamReader = new StreamReader(respone.GetResponseStream(), Encoding.UTF8); //讀取respone資料
            string result = streamReader.ReadToEnd(); //讀取到最後一行
            respone.Close();
            streamReader.Close();
            Sightseeing.Rootobject apiRequest = JsonConvert.DeserializeObject<Sightseeing.Rootobject>(result); //將資料轉為json物件

            return apiRequest;
        }
        private static object ToApiRequest(object requestObject)
        {
            switch (requestObject)
            {
                case JObject jObject: // objects become Dictionary<string,object>
                    return ((IEnumerable<KeyValuePair<string, JToken>>)jObject).ToDictionary(j => j.Key, j => ToApiRequest(j.Value));
                case JArray jArray: // arrays become List<object>
                    return jArray.Select(ToApiRequest).ToList();
                case JValue jValue: // values just become the value
                    return jValue.Value;
                default: // don't know what to do here
                    throw new Exception($"Unsupported type: {requestObject.GetType()}");
            }
        }

    }
}
