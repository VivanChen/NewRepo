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
        public async Task<Rootobject> GetapiAsync()
        {
            var tasks = new List<Task<string>>();
            HttpClientHelper httpClient = new HttpClientHelper();
            return await httpClient.Get(ConfigurationManager.AppSettings["Api_sightseeing"]);;
        }
    }
}
