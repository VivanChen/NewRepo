using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinFormsApp1.Controller;
using WinFormsApp1.Model;
using static WinFormsApp1.Model.Weather;

namespace WinFormsApp1
{
    /// <summary>
    /// 串接氣象API
    /// </summary>
    public partial class Form4 : Form
    {
        WeatherController weatherController = new WeatherController();
        AccountsController AccountController = new AccountsController();
        Rootobject rootobject = new Rootobject();
        public Accounts Account { get; set; }
        public Form4(Accounts Accounts_souce)
        {
            Account = Accounts_souce;
            InitializeComponent();
        }
        private void Form4_Load(object sender, EventArgs e)
        {
            if (!AccountController.Validate_account(Account, out string message))
            {
                MessageBox.Show("帳號密碼錯誤，請重新登入");
                Form3 f = new Form3();//產生Form1的物件，才可以使用它所提供的Method
                this.Hide();
                f.ShowDialog();
                this.Dispose();
            }
            else label4.Text = "您好 : "+Account.Account;

            rootobject = weatherController.Getapi();

            foreach (var item in rootobject.records.location)
            {
                comboBox1.Items.Add(item.locationName);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("請先選擇縣市");
                return;
            }
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            dataGridView1.Columns.Add("1", "縣市");
            dataGridView1.Columns.Add("2", "天氣狀況");
            dataGridView1.Columns.Add("3", "降雨機率");
            dataGridView1.Columns.Add("4", "最低溫度");
            dataGridView1.Columns.Add("5", "最高溫度");
            foreach (var item in rootobject.records.location)
            {

                if (comboBox1.SelectedItem.ToString() == "ALL")
                {
                    string[] Tamp = new string[]
                    {
                      item.locationName ,
                      item.weatherElement[0].time[0].parameter.parameterName ,
                      item.weatherElement[1].time[0].parameter.parameterName+"%" ,
                      item.weatherElement[2].time[0].parameter.parameterName ,
                      item.weatherElement[4].time[0].parameter.parameterName
                    };
                    dataGridView1.Rows.Add(Tamp);
                }
                else
                {
                    if (comboBox1.SelectedItem.ToString() == item.locationName)
                    {
                        string[] Tamp = new string[]
                        {
                          item.locationName ,
                          item.weatherElement[0].time[0].parameter.parameterName ,
                          item.weatherElement[1].time[0].parameter.parameterName+"%" ,
                          item.weatherElement[2].time[0].parameter.parameterName ,
                          item.weatherElement[4].time[0].parameter.parameterName
                        };
                        dataGridView1.Rows.Add(Tamp);
                    }
                }
            }
        }
        /// <summary>
        /// 登出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
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

        private void 台北市各級學校查詢ToolStripMenuItem_Click(object sender, EventArgs e)
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

        private void 天氣查詢ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form4 f = new Form4(Account);//產生Form2的物件，才可以使用它所提供的Method
            this.Hide();
            f.ShowDialog();
            this.Dispose();
        }

        private void 登出ToolStripMenuItem_Click(object sender, EventArgs e)
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

        private void 退出程序ToolStripMenuItem_Click(object sender, EventArgs e)
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
    }
}
