using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Model
{
    /// <summary>
    /// 108年台北市各級學校分布圖
    /// Type=國小、國中、高中、高職、特教學校、市立大專校院
    /// Name=name
    /// Postal=郵遞區號
    /// Address=地址
    /// Lat=經度
    /// LON=緯度
    /// </summary>
    public class School
    {
        public string Type { get; set; }
        public string Name { get; set; }
        public string Postal { get; set; }
        public string Address { get; set; }
        public string Telephone { get; set; }
        public string Lat { get; set; }
        public string Lon { get; set; }
    }
}
