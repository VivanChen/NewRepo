using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinFormsApp1.Model;
using static WinFormsApp1.Model.Sightseeing;

namespace WinFormsApp1.Controller
{
    public class SightseeingController
    {
        public static Rootobject rootobject { get;  set; }
        public Rootobject Getapi()
        {
            var tasks = new List<Task<string>>();
            HttpClientHelper httpClient = new HttpClientHelper();
            string result = httpClient.Get(ConfigurationManager.AppSettings["Api_sightseeing"]);
            rootobject = JsonConvert.DeserializeObject<Rootobject>(result);
            return rootobject;
        }
    }
}
