using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WinFormsApp1.Function;
using WinFormsApp1.Model;


namespace WinFormsApp1.Controller
{
    public class BusinessController : Business
    {
        List<Business> NewBusinesses = new List<Business>();
        public List<Business> GetBusinesses()
        {

            string zipFile = Path.Combine(Environment.CurrentDirectory, ConfigurationManager.AppSettings["DIR_Mode"], ConfigurationManager.AppSettings["Business_data"]+"ZIP"); // 待解壓縮的檔案
            string outDir = Path.Combine(Environment.CurrentDirectory, ConfigurationManager.AppSettings["DIR_Temp"]);  // 解壓縮目的目錄
            if (!File.Exists(Path.Combine(outDir, ConfigurationManager.AppSettings["Business_data"]+"CSV")))
            {
                Function_Zip.UnzipToDirectory(zipFile, outDir, true);
            }

            using (StreamReader streamReader = new StreamReader(Path.Combine(Environment.CurrentDirectory, ConfigurationManager.AppSettings["DIR_Temp"], ConfigurationManager.AppSettings["Business_data"]+"CSV"), Encoding.UTF8))
            {
                while (streamReader.Peek() >= 0)//由於該CSV檔應有格式編寫問題或是換行符號有少，若用ReadLine()判斷是否為Null來判斷是否讀取下一筆會有異常，筆數會少，改用Peek()來判斷
                {
                    //if (string.IsNullOrEmpty(streamReader.ReadLine())) { continue; }
                    string[] splite = streamReader.ReadLine().Split(",");
                    NewBusinesses.Add(new Business()
                    {
                        //由於來源資料有空白，且空白在字串內。使用正規表示式去除空白或多個空白
                        Business_address = Regex.Replace(splite[0], @"\s", ""),
                        Uniform_Numbers = Regex.Replace(splite[1], @"\s", ""),
                        Unified_number_office = Regex.Replace(splite[2], @"\s", ""),
                        Business_name = Regex.Replace(splite[3], @"\s", ""),
                        Capital = Regex.Replace(splite[4], @"\s", ""),
                        Date_Establishment = Regex.Replace(splite[5], @"\s", ""),
                        Organization_name = Regex.Replace(splite[6], @"\s", ""),
                        Uniform_Invoices = Regex.Replace(splite[7], @"\s", ""),
                        Industry_code = Regex.Replace(splite[8], @"\s", ""),
                        Name = Regex.Replace(splite[9], @"\s", ""),
                        IndustryCode1 = Regex.Replace(splite[10], @"\s", ""),
                        Name1 = Regex.Replace(splite[11], @"\s", ""),
                        IndustryCode2 = Regex.Replace(splite[12], @"\s", ""),
                        Name2 = Regex.Replace(splite[13], @"\s", ""),
                        IndustryCode3 = Regex.Replace(splite[14], @"\s", ""),
                        Name3 = Regex.Replace(splite[15], @"\s", ""),
                    });
                }
            }
                          
            return NewBusinesses;
        }



    }
}


