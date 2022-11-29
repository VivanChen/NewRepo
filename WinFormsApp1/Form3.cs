using Google.Authenticator;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinFormsApp1.Controller;
using WinFormsApp1.Model;

namespace WinFormsApp1
{
    /// <summary>
    /// 登入頁面
    /// Todo:帳號密碼+OTP(多重要素驗證 微軟、Google )
    /// Todo:帳密錯誤限制
    /// Todo:OTP驗證錯誤限制
    /// Todo:登入訊息紀錄儲存
    /// </summary>
    public partial class Form3 : Form
    {
        Timer Timer = new Timer();//計時器物件
        Accounts Accounts_souce = new Accounts();
        GoogleAnalyticsController GoogleController = new GoogleAnalyticsController();
        AccountsController AccountController = new AccountsController();
        private int duration = 60;
        int error_number = 0;

        public Form3()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 檢核帳號密碼
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            string message = string.Empty;
            if (string.IsNullOrEmpty(textBox1.Text) || string.IsNullOrEmpty(textBox2.Text))
            {
                //todo:改寫 MessageBox.Show 窗體位置於程式內置中

                MessageBox.Show("請正確填寫帳號或密碼");
                return;
            }
            Accounts_souce.Account = textBox1.Text;
            Accounts_souce.Password = textBox2.Text;
            if (AccountController.Validate_account(Accounts_souce, out message))
            {
                error_number = 0;
                panel3.Visible = true;
                button1.Enabled = false;
                GoogleController.CreateSecretKeyAndQrCode(Accounts_souce, out SetupCode code, out message);

                //QRCode圖片從記憶體轉到畫面上
                using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(code.QrCodeSetupImageUrl.Replace("data:image/png;base64,", ""))))
                    pictureBox1.Image = Image.FromStream(ms);
            }
            else
            {
                if (message != "")
                {
                    MessageBox.Show(message);
                }
                else
                {
                    if (error_number >= 2)
                    {
                        Timer.Tick += new EventHandler(Account_down);
                        Timer.Interval = 1000;
                        button1.Enabled = false;
                        Timer.Start();

                    }
                    MessageBox.Show("帳號密碼輸入錯誤，請重新輸入");
                    error_number++;

                }
            }
        }
        /// <summary>
        /// 帳號計時器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Account_down(object sender, EventArgs e)
        {

            if (duration == 0)
            {
                Timer.Stop();
                button1.Text = "登入";
                error_number = 0;
                duration = 60;
                button1.Enabled = true;
            }
            else if (duration > 0)
            {
                duration--;
                button1.Text = $"您已錯誤三次，請等待 : ({duration.ToString()})秒";
            }
        }
        /// <summary>
        /// OTP計時器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Otp_down(object sender, EventArgs e)
        {

            if (duration == 0)
            {
                Timer.Stop();
                button2.Text = "Otp驗證登入";
                error_number = 0;
                duration = 60;
                button2.Enabled = true;
            }
            else if (duration > 0)
            {
                duration--;
                button2.Text = $"您已錯誤三次，請等待 : ({duration.ToString()})秒";
            }
        }

        /// <summary>
        /// OTP驗證
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox3.Text))
            {
                MessageBox.Show("請填入驗證碼後，再按下OTP驗證登入");
            }
            Accounts_souce.otp = textBox3.Text.Trim();
            if (!GoogleController.ValidateGoogleAuthCode(Accounts_souce))
            {

                if (error_number >= 2)
                {
                    Timer.Tick += new EventHandler(Otp_down);
                    Timer.Interval = 1000;
                    button2.Enabled = false;
                    Timer.Start();

                }
                MessageBox.Show("OTP輸入錯誤，請重新輸入");
                error_number++;
            }
            else
            {
                Accounts_souce.Loging_Time = DateTime.Now.ToString();
                Form1 f = new Form1(Accounts_souce);
                this.Hide();
                f.ShowDialog();
                this.Dispose();
            }

        }
    }
}

