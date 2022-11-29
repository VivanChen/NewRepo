using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinFormsApp1.Model;

namespace WinFormsApp1.Controller
{
    public class SchoolController
    {
        public List<School> Getschools()
        {
            List<School> newschools = new List<School>();
            StreamReader reader = new StreamReader(Path.Combine(Environment.CurrentDirectory, ConfigurationManager.AppSettings["DIR_Mode"],"1080416臺北市各級學校分布圖.csv"), Encoding.UTF8);
            while (reader.ReadLine() != null)
            {
                if (string.IsNullOrEmpty(reader.ReadLine())) { continue; }
                string[] splite = reader.ReadLine().Split(",");
                newschools.Add(new School() { Type = splite[0], 
                                              Name = splite[1], 
                                              Postal = splite[2], 
                                              Address = splite[3], 
                                              Telephone = splite[4], 
                                              Lat = splite[5], 
                                              Lon = splite[6] });
            }
            return newschools;
        }
    }
}
