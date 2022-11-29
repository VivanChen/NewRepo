using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    public class HttpClientHelper
    {
        private static readonly object LockObj = new object();
        private static HttpClient client = null;
        HttpClientHandler handler = new HttpClientHandler();
        public HttpClientHelper()
        {
            handler.UseProxy = false;//不加這個會非常慢
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

        public string Get(string url)
        {
            try
            {
                var responseString = client.GetStringAsync(url);
                return responseString.Result;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

    }
}
