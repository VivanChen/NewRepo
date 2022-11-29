using Newtonsoft.Json.Linq;
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
using static WinFormsApp1.Model.Sightseeing;

namespace WinFormsApp1
{
    public partial class Form5 : Form
    {
        AccountsController AccountController = new AccountsController();
        SightseeingController sightseeingController = new SightseeingController();
        public static Rootobject rootobject { get;  set; }

        public Accounts Account { get; set; }
        public Form5(Accounts Accounts_souce)
        {
            Account = Accounts_souce;
            InitializeComponent();
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            if (!AccountController.Validate_account(Account, out string message))
            {
                MessageBox.Show("帳號密碼錯誤，請重新登入");
                Form3 f = new Form3();//產生Form1的物件，才可以使用它所提供的Method
                this.Hide();
                f.ShowDialog();
                this.Dispose();
            }
            else label1.Text = "您好 : " + Account.Account;
            rootobject = sightseeingController.Getapi();
            var query =rootobject.XML_Head.Infos.Info.Where(a=>a.Region!=null).GroupBy(a => a.Region).Select(a => a.Key);
            foreach (var item in query)
            {
                comboBox1.Items.Add(item);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var query = rootobject.XML_Head.Infos.Info.Where(p=>p.Region!=null && p.Region==comboBox1.SelectedItem.ToString()).ToList();

            dataGridView1.DataSource = query;
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

        private void 旅遊景點查詢ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form5 f = new Form5(Account);//產生Form2的物件，才可以使用它所提供的Method
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

    }
}