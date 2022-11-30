using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinFormsApp1
{
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {
            this.pictureBox1.Image = Image.FromFile(Path.Combine(Environment.CurrentDirectory, ConfigurationManager.AppSettings["DIR_Image"],"Lodeing.gif"));
            this.label1.Text = "檔案讀取中...請稍後";
        }

    }
}
