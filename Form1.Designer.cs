


namespace checkdataCollect
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.collectbtn = new System.Windows.Forms.Button();
            this.selectfolderbtn = new System.Windows.Forms.Button();
            this.txtboxPath = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.button1 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.button_AddData = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.Orange;
            this.label1.Location = new System.Drawing.Point(12, 58);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 19);
            this.label1.TabIndex = 0;
            this.label1.Text = "选择源文件:";
            // 
            // collectbtn
            // 
            this.collectbtn.BackColor = System.Drawing.Color.Transparent;
            this.collectbtn.Font = new System.Drawing.Font("微软雅黑", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.collectbtn.Location = new System.Drawing.Point(101, 210);
            this.collectbtn.Name = "collectbtn";
            this.collectbtn.Size = new System.Drawing.Size(93, 44);
            this.collectbtn.TabIndex = 1;
            this.collectbtn.Text = "汇总";
            this.collectbtn.UseVisualStyleBackColor = false;
            this.collectbtn.Click += new System.EventHandler(this.collectbtn_Click);
            // 
            // selectfolderbtn
            // 
            this.selectfolderbtn.BackColor = System.Drawing.Color.Transparent;
            this.selectfolderbtn.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.selectfolderbtn.Location = new System.Drawing.Point(563, 58);
            this.selectfolderbtn.Name = "selectfolderbtn";
            this.selectfolderbtn.Size = new System.Drawing.Size(47, 23);
            this.selectfolderbtn.TabIndex = 2;
            this.selectfolderbtn.Text = "浏览";
            this.selectfolderbtn.UseVisualStyleBackColor = false;
            this.selectfolderbtn.Click += new System.EventHandler(this.selectfolderbtn_Click);
            // 
            // txtboxPath
            // 
            this.txtboxPath.Enabled = false;
            this.txtboxPath.Font = new System.Drawing.Font("微软雅黑", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtboxPath.Location = new System.Drawing.Point(101, 58);
            this.txtboxPath.Name = "txtboxPath";
            this.txtboxPath.Size = new System.Drawing.Size(456, 21);
            this.txtboxPath.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.ForeColor = System.Drawing.Color.Orange;
            this.label2.Location = new System.Drawing.Point(539, 351);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 23);
            this.label2.TabIndex = 4;
            // 
            // progressBar1
            // 
            this.progressBar1.BackColor = System.Drawing.Color.White;
            this.progressBar1.Location = new System.Drawing.Point(4, 378);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(614, 23);
            this.progressBar1.TabIndex = 5;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Transparent;
            this.button1.Font = new System.Drawing.Font("微软雅黑", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(442, 210);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(115, 44);
            this.button1.TabIndex = 6;
            this.button1.Text = "退出程序";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.ForeColor = System.Drawing.Color.Orange;
            this.label3.Location = new System.Drawing.Point(9, 357);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 12);
            this.label3.TabIndex = 7;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.BackColor = System.Drawing.Color.Transparent;
            this.radioButton1.Checked = true;
            this.radioButton1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.radioButton1.ForeColor = System.Drawing.Color.Coral;
            this.radioButton1.Location = new System.Drawing.Point(101, 106);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(108, 25);
            this.radioButton1.TabIndex = 8;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "实验室数据";
            this.radioButton1.UseVisualStyleBackColor = false;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.BackColor = System.Drawing.Color.Transparent;
            this.radioButton2.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.radioButton2.ForeColor = System.Drawing.Color.Coral;
            this.radioButton2.Location = new System.Drawing.Point(229, 106);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(124, 25);
            this.radioButton2.TabIndex = 9;
            this.radioButton2.Text = "物理检查数据";
            this.radioButton2.UseVisualStyleBackColor = false;
            this.radioButton2.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            // 
            // button_AddData
            // 
            this.button_AddData.BackColor = System.Drawing.Color.Transparent;
            this.button_AddData.Font = new System.Drawing.Font("微软雅黑", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button_AddData.Location = new System.Drawing.Point(260, 210);
            this.button_AddData.Name = "button_AddData";
            this.button_AddData.Size = new System.Drawing.Size(129, 44);
            this.button_AddData.TabIndex = 10;
            this.button_AddData.Text = "添加数据";
            this.button_AddData.UseVisualStyleBackColor = false;
            this.button_AddData.Click += new System.EventHandler(this.button_AddData_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::checkdataCollect.Properties.Resources.shuidi01;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(622, 424);
            this.ControlBox = false;
            this.Controls.Add(this.button_AddData);
            this.Controls.Add(this.radioButton2);
            this.Controls.Add(this.radioButton1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtboxPath);
            this.Controls.Add(this.selectfolderbtn);
            this.Controls.Add(this.collectbtn);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(638, 440);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(638, 440);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button collectbtn;
        private System.Windows.Forms.Button selectfolderbtn;
        private System.Windows.Forms.TextBox txtboxPath;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.Button button_AddData;
    }
}

