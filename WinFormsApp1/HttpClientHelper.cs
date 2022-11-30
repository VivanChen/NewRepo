using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using static WinFormsApp1.Model.Sightseeing;

namespace WinFormsApp1
{
    public class HttpClientHelper
    {
        private static readonly object LockObj = new object();
        private static HttpClient client = null;
        HttpClientHandler handler = new HttpClientHandler();
        public HttpClientHelper()
        {          
            client = new HttpClient() { BaseAddress = new Uri("https://www.google.com/") };

            //帮HttpClient热身
            client.SendAsync(new HttpRequestMessage
            {
                Method = new HttpMethod("HEAD"),
                RequestUri = new Uri("https://www.google.com/" + "/")
            })
                .Result.EnsureSuccessStatusCode();
            GetInstance();
        }
        public static HttpClient GetInstance()
        {

            if (client == null)
            {
                lock (LockObj)
                {
                    if (client == null)
                    {
                        client = new HttpClient();
                    }
                }
            }
            return client;
        }
        public async Task<string> PostAsync(string url, string strJson)//post異步請求方法
        {
            try
            {
                HttpContent content = new StringContent(strJson);
                content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");
                //由HttpClient發出異步Post請求
                HttpResponseMessage res = await client.PostAsync(url, content);
                if (res.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    string str = res.Content.ReadAsStringAsync().Result;
                    return str;
                }
                else
                    return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public string Post(string url, string strJson)//post同步請求方法
        {
            try
            {
                HttpContent content = new StringContent(strJson);
                content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");
                //client.DefaultRequestHeaders.Connection.Add("keep-alive");
                //由HttpClient發出Post請求
                Task<HttpResponseMessage> res = client.PostAsync(url, content);
                if (res.Result.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    string str = res.Result.Content.ReadAsStringAsync().Result;
                    return str;
                }
                else
                    return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        /// <summary>
        /// 使用HttpClient下載大型檔案時會異常緩慢
        /// 利用非同步將檔案下載成流數據(Steam)
        /// 讀取後再轉為物件
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public async Task<Rootobject> Get(string url)
        {
            try
            {
                var responseString = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);
                responseString.EnsureSuccessStatusCode();
                using (var stream = await responseString.Content.ReadAsStreamAsync())
                using (var streamReader = new StreamReader(stream))
                using (var jsonReader = new JsonTextReader(streamReader))
                {
                    var serializer = new JsonSerializer();
                    Rootobject p = serializer.Deserialize<Rootobject>(jsonReader);
                    return p;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

    }
}
