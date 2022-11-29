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

namespace WinFormsApp1
{
    public partial class Form2 : Form
    {
        AccountsController AccountController = new AccountsController();
        List<Business> NewBusinesses = new List<Business>();
        public Accounts Account { get; set; }
        public Form2(Accounts Accounts_souce)
        {
            Account = Accounts_souce;
            ControlBox = false;
            InitializeComponent();
        }
        BusinessController Controller = new BusinessController();

        private void button2_Click(object sender, EventArgs e)
        {
            Form1 f = new Form1(Account);//產生Form1的物件，才可以使用它所提供的Method
            this.Hide();
            f.ShowDialog();
            this.Dispose();
        }
        private void Form2_Load(object sender, EventArgs e)
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
        }
        /// <summary>
        /// 非同步依照區域查詢
        /// 解決 UI 畫面卡死
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void button1_Click(object sender, EventArgs e) //async關鍵字
        {
            userControl11.Visible = true;//主執行緒
            userControl11.Refresh();

            await Task.Run(() =>//建立獨立執行緒去執行讀取檔案 (awit關鍵字)
            {
                if (NewBusinesses.Count==0)
                {
                    NewBusinesses = Controller.GetBusinesses();                   
                }
            });
            var query = (NewBusinesses.Where(a => a.Business_address.Substring(1, 3).ToUpper().Trim() == comboBox1.SelectedItem.ToString())).OrderBy(a => a.Uniform_Numbers);
            ShowText(query);
            userControl11.Visible = false;

        }
        /// <summary>
        /// 修改dataGridView標頭英轉中
        /// </summary>
        private void ChangeTitle()
        {
            Business business = new Business();
            dataGridView1.Columns["Business_address"].HeaderText = "營業地址";
            dataGridView1.Columns["Uniform_Numbers"].HeaderText = "統一編號";
            dataGridView1.Columns["Unified_number_office"].HeaderText = "總機構統一編號";
            dataGridView1.Columns["Business_name"].HeaderText = "營業人名稱";
            dataGridView1.Columns["Capital"].HeaderText = "資本額";
            dataGridView1.Columns["Date_Establishment"].HeaderText = "設立日期";
            dataGridView1.Columns["Organization_name"].HeaderText = "組織別名稱";
            dataGridView1.Columns["Uniform_Invoices"].HeaderText = "使用統一發票";
            dataGridView1.Columns["Industry_code"].HeaderText = "行業代號";
            dataGridView1.Columns["Name"].HeaderText = "名稱";
            dataGridView1.Columns["IndustryCode1"].HeaderText = "行業代號1";
            dataGridView1.Columns["Name1"].HeaderText = "名稱1";
            dataGridView1.Columns["IndustryCode2"].HeaderText = "行業代號2";
            dataGridView1.Columns["Name2"].HeaderText = "名稱2";
            dataGridView1.Columns["IndustryCode3"].HeaderText = "行業代號3";
            dataGridView1.Columns["Name3"].HeaderText = "名稱3";
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

        }
        /// <summary>
        /// 顯示
        /// </summary>
        /// <param name="businesses"></param>
        private void ShowText(IEnumerable<Business> businesses)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = businesses.ToList();
            ChangeTitle();
        }
        /// <summary>
        /// 依照公司名稱模糊查詢
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            var query = NewBusinesses.Where(a => a.Business_name.ToUpper().Trim().Contains(textBox1.Text.ToUpper().Trim()));
            ShowText(query);
        }
        /// <summary>
        /// 自動編號
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //自動編號，與資料無關
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X,
               e.RowBounds.Location.Y,
               dataGridView1.RowHeadersWidth - 4,
               e.RowBounds.Height);

            TextRenderer.DrawText(e.Graphics,
                  (e.RowIndex + 1).ToString(),
                   dataGridView1.RowHeadersDefaultCellStyle.Font,
                   rectangle,
                   dataGridView1.RowHeadersDefaultCellStyle.ForeColor,
                   TextFormatFlags.Default | TextFormatFlags.Default);
        }

        private void 台北各級學校查詢ToolStripMenuItem_Click(object sender, EventArgs e)
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
