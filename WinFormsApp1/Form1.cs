using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinFormsApp1.Controller;
using WinFormsApp1.Function;
using WinFormsApp1.Model;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        AccountsController AccountController = new AccountsController();
        public Accounts Account { get; set; }
        public Form1(Accounts Accounts_souce)
        {
            Account = Accounts_souce;
            ControlBox = false;
            InitializeComponent();//讀取頁面相關
        }
        
        public string SaveFileName { get; set; }
        //建立List<School>物件以方便後續調用，宣告在此處該物件不會被自動回收
        List<School> NewSchool = new List<School>();

        /// <summary>
        /// WinForm.Form1啟動時會先載入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            if (!AccountController.Validate_account(Account, out string message))
            {
                MessageBox.Show("帳號密碼錯誤，請重新登入");
                Form3 f = new Form3();//產生Form1的物件，才可以使用它所提供的Method
                this.Hide();
                f.ShowDialog();
                this.Dispose();
            }
            else label5.Text = "您好 ： "+Account.Account;
            //讀取csv檔
            Getschools();
            //篩選出學校類型(TYPE)
            var query = NewSchool.GroupBy(a => a.Type).Select(a => a.Key);
            //將查詢結果放入comboBox
            foreach (var item in query)
            {
                comboBox1.Items.Add(item);
            }
        }

        /// <summary>
        /// 資料寫入dataGridView
        /// </summary>
        /// <param name="schools"></param>
        private void ShowText(IEnumerable<School> schools)
        {
            dataGridView1.DataSource = schools.ToList();
        }


        /// <summary>
        /// 取得來源資料
        /// </summary>
        private void Getschools()
        {
            //建立SchoolController物件
            SchoolController school = new SchoolController();
            //讀取csv檔
            NewSchool = school.Getschools();
        }

        /// <summary>
        /// 依學校類型查詢
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (NewSchool.Count == 0) { Getschools(); }
            if (comboBox1.SelectedItem == null) { MessageBox.Show("請先選擇學校類型"); return; }
            var query = NewSchool.Where(a => a.Type == comboBox1.SelectedItem.ToString());
            ShowText(query);
        }
        /// <summary>
        /// 依照行政區域查詢(模糊查詢)
        /// 可使用";"或","一次查詢多種條件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            List<School> Tamp = new List<School>();
            string[] conditions = new string[]{ ",", ";" };
            string condition = string.Empty;
            foreach (var item in conditions)
            {
                if (textBox1.Text.Contains(item)) { condition = item; } 
            }
            if (NewSchool.Count == 0) { Getschools(); }
            if (string.IsNullOrEmpty(textBox1.Text)) { MessageBox.Show("請先填寫區域內容"); return; }
            else
            {
                //switch寫法
                switch (condition)
                {                    
                    case ",":
                        dataGridView1.DataSource = null;
                        string[] splite1 = textBox1.Text.Split(",");
                        if (splite1.Length == 0) { MessageBox.Show("查無資料，請重新輸入"); return; }
                        foreach (string item in splite1)
                        {
                            Tamp.AddRange(from a in NewSchool.AsEnumerable() where a.Address.Substring(3, 3).ToUpper().Contains(item.ToUpper()) select a);
                        }
                        ShowText(Tamp);
                        break;
                    case ";":
                        dataGridView1.DataSource = null;
                        string[] splite2 = textBox1.Text.Split(";");
                        if (splite2.Length == 0) { MessageBox.Show("查無資料，請重新輸入"); return; }
                        foreach (string item in splite2)
                        {
                            Tamp.AddRange(from a in NewSchool.AsEnumerable() where a.Address.Substring(3, 3).ToUpper().Contains(item.ToUpper()) select a);
                        }
                        ShowText(Tamp);
                        break;
                    default:
                        dataGridView1.DataSource = null;
                        var query1 = NewSchool.Where(a => a.Address.Substring(3, 3).ToUpper().Contains(textBox1.Text.ToUpper()));
                        if (query1.Count() == 0) { MessageBox.Show("查無資料，請重新輸入"); return; }
                        ShowText(query1);
                        break;
                }
                //if else 寫法

                //if (textBox1.Text.Contains(","))
                //{
                //    dataGridView1.DataSource = null;
                //    string[] splite = textBox1.Text.Split(",");
                //    if (splite.Length == 0) { MessageBox.Show("查無資料，請重新輸入"); return; }
                //    foreach (string item in splite)
                //    {
                //        Tamp.AddRange(from a in NewSchool.AsEnumerable() where a.Address.Substring(3, 3).ToUpper().Contains(item.ToUpper()) select a);
                //    }
                //    ShowText(Tamp);
                //}
                //else if (textBox1.Text.Contains(";"))
                //{
                //    dataGridView1.DataSource = null;
                //    string[] splite = textBox1.Text.Split(";");
                //    if (splite.Length == 0) { MessageBox.Show("查無資料，請重新輸入"); return; }
                //    foreach (string item in splite)
                //    {
                //        Tamp.AddRange(from a in NewSchool.AsEnumerable() where a.Address.Substring(3, 3).ToUpper().Contains(item.ToUpper()) select a);
                //    }
                //    ShowText(Tamp);
                //}
                //else
                //{
                //    dataGridView1.DataSource = null;
                //    var query1 = NewSchool.Where(a => a.Address.Substring(3, 3).ToUpper().Contains(textBox1.Text.ToUpper()));
                //    if (query1.Count() == 0) { MessageBox.Show("查無資料，請重新輸入"); return; }
                //    ShowText(query1);
                //}

            }
        }
        /// <summary>
        /// 依照電話號碼查詢(完整查詢)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            string phone = textBox2.Text;
            string patte = @"^02-[0-9]{4}-[0-9]{4}";
            bool check = Regex.IsMatch(phone, patte);
            //檢查使用輸入的電話格式是否正確
            if (check)
            {
                var query = NewSchool.Where(a => a.Telephone.Trim() == textBox2.Text.Trim());
                if (query.Count() == 0) { MessageBox.Show("查無此號碼，請重新輸入"); }
                ShowText(query);
            }
            else
            {
                MessageBox.Show("電話號碼格式錯誤！範例：02-1234-5678");
            }
            
        }
        /// <summary>
        /// 根據區域分別輸出檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            //從schoolList找出區域群組後透過ToDictionary將分組重新組成Dictionary<string,List<School>>物件
            var query = NewSchool.GroupBy(p => p.Address.Substring(3, 3).ToUpper()).ToDictionary(p => p.Key, p => p.ToList());

            //解開Dictionary 迴圈查詢 Keyvalue<string,List<School>>
            foreach (var list in query)
            {
                //迴圈查詢Keyvalue下的值(School物件)
                foreach (var info in list.Value)
                {
                    string strRegion = list.Key;
                    if(!Export_csv(info, strRegion, out string message)) MessageBox.Show(message);
                }
            }
        }
        /// <summary>
        /// 當滑鼠點擊textBox1時，清空提示
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "可輸入多條件以,或;區隔條件  例如:北投區,大同區")
            {
                textBox1.Clear();
            }
        }
        /// <summary>
        /// 匯出Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            Button_Close();
            Function_Excel export = new Function_Excel();
            export.ExportExcel(this.SaveFileName, dataGridView1, "新細明體", 11);//默認文件名,DataGridView控件的名稱,字體,字號
            Button_Open();
        }
        /// <summary>
        /// 依照區域輸出csv檔於程式執行目錄
        /// </summary>
        /// <param name="info"></param>
        /// <param name="region"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static bool Export_csv(School info, string region, out string message)
        {
            bool check = false;
            try
            {
                //組合需要輸出字串
                string s = $"{info.Type},{info.Name},{info.Postal},{info.Address},{info.Telephone},{info.Lat},{info.Lon}\r\n";
                if (region != "")
                {
                    //若不指定路徑AppendAllText()會預設檔案路徑於程式執行目錄下
                    File.AppendAllText($"{region}.csv", s, Encoding.UTF8);
                    check = true;
                }
                message = "";
                return check;
            }
            //利用try catch 若有執行錯誤將錯誤訊息，用Out參數修飾詞傳遞
            catch (Exception ex)
            {
                message = ex.ToString();
                return check;
            }          
        }

        /// <summary>
        /// 關閉按鈕
        /// </summary>
        private void Button_Close()
        {
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
        }
        /// <summary>
        /// 打開按鈕
        /// </summary>
        private void Button_Open()
        {
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;
        }

        private void 台北市各級學校查詢ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Form1 f = new Form1(Account);//產生Form2的物件，才可以使用它所提供的Method
            this.Hide();
            f.ShowDialog();
            this.Dispose();
        }

        private void 台灣商家黃頁ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 f = new Form2(Account);//產生Form2的物件，才可以使用它所提供的Method
            this.Hide();
            f.ShowDialog();
            this.Dispose();
        }

        private void 天氣查詢ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Form4 f = new Form4(Account);//產生Form2的物件，才可以使用它所提供的Method
            this.Hide();
            f.ShowDialog();
            this.Dispose();
        }

        private void 登出ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            DialogResult Result = MessageBox.Show("你確定要登出嗎??", "請確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            if (Result == DialogResult.OK)
            {
                MessageBox.Show("您已登出");
                Form3 f = new Form3();
                this.Hide();
                f.ShowDialog();
                this.Dispose();
            }
            else if (Result == DialogResult.Cancel)
            {
                return;
            }
        }

        private void 退出程序ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            DialogResult Result = MessageBox.Show("你確定要退出嗎??", "請確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            if (Result == DialogResult.OK)
            {
                Application.Exit();
            }
            else if (Result == DialogResult.Cancel)
            {
                return;
            }
        }

        private void 旅遊景點查詢ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form5 f = new Form5(Account);//產生Form2的物件，才可以使用它所提供的Method
            this.Hide();
            f.ShowDialog();
            this.Dispose();
        }
        //TODo
        //1.修改功能
        //2.刪除功能
        //3.新增功能
        //4.分頁功能

    }
}
