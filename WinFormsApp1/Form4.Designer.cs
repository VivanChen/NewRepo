
namespace WinFormsApp1
{
    partial class Form4
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel5 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.cSVToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aPIToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.功能ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.台北市各級學校查詢ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.台灣商家黃頁ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.天氣查詢ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.登出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.退出程序ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.旅遊景點查詢ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panel5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel5
            // 
            this.panel5.BackColor = System.Drawing.SystemColors.Highlight;
            this.panel5.Controls.Add(this.label4);
            this.panel5.Controls.Add(this.label8);
            this.panel5.Location = new System.Drawing.Point(0, 26);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(800, 70);
            this.panel5.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label4.Location = new System.Drawing.Point(703, 27);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(46, 17);
            this.label4.TabIndex = 12;
            this.label4.Text = "label4";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label8.Location = new System.Drawing.Point(7, 19);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(158, 31);
            this.label8.TabIndex = 5;
            this.label8.Text = "天氣查詢功能";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 136);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(31, 15);
            this.label1.TabIndex = 8;
            this.label1.Text = "縣市";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(27, 162);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(158, 23);
            this.button1.TabIndex = 9;
            this.button1.Text = "查詢";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "ALL"});
            this.comboBox1.Location = new System.Drawing.Point(64, 133);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 23);
            this.comboBox1.TabIndex = 10;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(191, 133);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 25;
            this.dataGridView1.Size = new System.Drawing.Size(585, 307);
            this.dataGridView1.TabIndex = 11;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cSVToolStripMenuItem,
            this.aPIToolStripMenuItem,
            this.功能ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 24);
            this.menuStrip1.TabIndex = 12;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // cSVToolStripMenuItem
            // 
            this.cSVToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.台北市各級學校查詢ToolStripMenuItem,
            this.台灣商家黃頁ToolStripMenuItem});
            this.cSVToolStripMenuItem.Name = "cSVToolStripMenuItem";
            this.cSVToolStripMenuItem.Size = new System.Drawing.Size(42, 20);
            this.cSVToolStripMenuItem.Text = "CSV";
            // 
            // aPIToolStripMenuItem
            // 
            this.aPIToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.天氣查詢ToolStripMenuItem,
            this.旅遊景點查詢ToolStripMenuItem});
            this.aPIToolStripMenuItem.Name = "aPIToolStripMenuItem";
            this.aPIToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.aPIToolStripMenuItem.Text = "API";
            // 
            // 功能ToolStripMenuItem
            // 
            this.功能ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.登出ToolStripMenuItem,
            this.退出程序ToolStripMenuItem});
            this.功能ToolStripMenuItem.Name = "功能ToolStripMenuItem";
            this.功能ToolStripMenuItem.Size = new System.Drawing.Size(43, 20);
            this.功能ToolStripMenuItem.Text = "功能";
            // 
            // 台北市各級學校查詢ToolStripMenuItem
            // 
            this.台北市各級學校查詢ToolStripMenuItem.Name = "台北市各級學校查詢ToolStripMenuItem";
            this.台北市各級學校查詢ToolStripMenuItem.Size = new System.Drawing.Size(182, 22);
            this.台北市各級學校查詢ToolStripMenuItem.Text = "台北市各級學校查詢";
            this.台北市各級學校查詢ToolStripMenuItem.Click += new System.EventHandler(this.台北市各級學校查詢ToolStripMenuItem_Click);
            // 
            // 台灣商家黃頁ToolStripMenuItem
            // 
            this.台灣商家黃頁ToolStripMenuItem.Name = "台灣商家黃頁ToolStripMenuItem";
            this.台灣商家黃頁ToolStripMenuItem.Size = new System.Drawing.Size(182, 22);
            this.台灣商家黃頁ToolStripMenuItem.Text = "台灣商家黃頁";
            this.台灣商家黃頁ToolStripMenuItem.Click += new System.EventHandler(this.台灣商家黃頁ToolStripMenuItem_Click);
            // 
            // 天氣查詢ToolStripMenuItem
            // 
            this.天氣查詢ToolStripMenuItem.Name = "天氣查詢ToolStripMenuItem";
            this.天氣查詢ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.天氣查詢ToolStripMenuItem.Text = "天氣查詢";
            this.天氣查詢ToolStripMenuItem.Click += new System.EventHandler(this.天氣查詢ToolStripMenuItem_Click);
            // 
            // 登出ToolStripMenuItem
            // 
            this.登出ToolStripMenuItem.Name = "登出ToolStripMenuItem";
            this.登出ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.登出ToolStripMenuItem.Text = "登出";
            this.登出ToolStripMenuItem.Click += new System.EventHandler(this.登出ToolStripMenuItem_Click);
            // 
            // 退出程序ToolStripMenuItem
            // 
            this.退出程序ToolStripMenuItem.Name = "退出程序ToolStripMenuItem";
            this.退出程序ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.退出程序ToolStripMenuItem.Text = "退出程序";
            this.退出程序ToolStripMenuItem.Click += new System.EventHandler(this.退出程序ToolStripMenuItem_Click);
            // 
            // 旅遊景點查詢ToolStripMenuItem
            // 
            this.旅遊景點查詢ToolStripMenuItem.Name = "旅遊景點查詢ToolStripMenuItem";
            this.旅遊景點查詢ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.旅遊景點查詢ToolStripMenuItem.Text = "旅遊景點查詢";
            this.旅遊景點查詢ToolStripMenuItem.Click += new System.EventHandler(this.旅遊景點查詢ToolStripMenuItem_Click);
            // 
            // Form4
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 461);
            this.ControlBox = false;
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel5);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form4";
            this.Text = "Form4";
            this.Load += new System.EventHandler(this.Form4_Load);
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem cSVToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aPIToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 功能ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 台北市各級學校查詢ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 台灣商家黃頁ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 天氣查詢ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 登出ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 退出程序ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 旅遊景點查詢ToolStripMenuItem;
    }
}