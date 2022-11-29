using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Model
{
    public class Sightseeing
    {

        public class Rootobject
        {
            public XML_Head XML_Head { get; set; }
        }

        public class XML_Head
        {
            public string Listname { get; set; }
            public string Language { get; set; }
            public string Orgname { get; set; }
            public DateTime Updatetime { get; set; }
            public Infos Infos { get; set; }
        }

        public class Infos
        {
            public List<Info> Info { get; set; }
        }

        public class Info
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public string Zone { get; set; }
            public string Toldescribe { get; set; }
            public string Description { get; set; }
            public string Tel { get; set; }
            public string Add { get; set; }
            public string Zipcode { get; set; }
            public string Region { get; set; }
            public string Town { get; set; }
            public string Travellinginfo { get; set; }
            public string Opentime { get; set; }
            public string Picture1 { get; set; }
            public string Picdescribe1 { get; set; }
            public string Picture2 { get; set; }
            public string Picdescribe2 { get; set; }
            public string Picture3 { get; set; }
            public string Picdescribe3 { get; set; }
            public string Map { get; set; }
            public string Gov { get; set; }
            public float Px { get; set; }
            public float Py { get; set; }
            public string Orgclass { get; set; }
            public string Class1 { get; set; }
            public string Class2 { get; set; }
            public string Class3 { get; set; }
            public string Level { get; set; }
            public string Website { get; set; }
            public string Parkinginfo { get; set; }
            public float? Parkinginfo_Px { get; set; }
            public float? Parkinginfo_Py { get; set; }
            public string Ticketinfo { get; set; }
            public string Remarks { get; set; }
            public string Keyword { get; set; }
            public DateTime? Changetime { get; set; }
        }


    }
}
