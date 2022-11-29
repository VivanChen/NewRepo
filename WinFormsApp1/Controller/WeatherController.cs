using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinFormsApp1.Function;
using WinFormsApp1.Model;
using static WinFormsApp1.Model.Weather;

namespace WinFormsApp1.Controller
{
    public class WeatherController
    {
        public Rootobject Getapi()
        {
            var tasks = new List<Task<string>>();
            Rootobject weather = new Rootobject();
            UsersClient usersClient = new UsersClient();
            weather = usersClient.getJson(ConfigurationManager.AppSettings["ApiUatBase"] + Function_Decrypt.DecryptDES(ConfigurationManager.AppSettings["API-Key"]) + "&format=JSON");
            return weather;


        }
    }
}
