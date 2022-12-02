using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Model
{

    public class Register
    {
        //科目
        public string Type { get; set; }
        //測驗日期
        public string Testdate { get; set; }
        //編號
        public string Number { get; set; }
        //測驗時間
        public string Testtime { get; set; }
        //報名單位
        public string Unit { get; set; }
        //年度
        public string Year { get; set; }
        //地區
        public string Region { get; set; }
        //姓名
        public string Name { get; set; }
        //試場
        public string TestField { get; set; }
        //出身年月日
        public string Birthday { get; set; }
        //座號
        public string SeatNumber { get; set; }
        //身分證統一編號
        public string TextId { get; set; }
        //公司代碼
        public string CompanyCode { get; set; }
        //測驗地點
        public string Location { get; set; }
        //測驗地址
        public string Address { get; set; }
       
    }
}
