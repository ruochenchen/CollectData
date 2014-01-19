namespace checkdataCollect
{
    partial class Form2
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
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button_Scan = new System.Windows.Forms.Button();
            this.button_Comfirm = new System.Windows.Forms.Button();
            this.button_Return = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "目标文件路径：";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(120, 10);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(404, 21);
            this.textBox1.TabIndex = 1;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(120, 51);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(404, 21);
            this.textBox2.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "选择源文件路径：";
            // 
            // button_Scan
            // 
            this.button_Scan.Location = new System.Drawing.Point(530, 49);
            this.button_Scan.Name = "button_Scan";
            this.button_Scan.Size = new System.Drawing.Size(49, 23);
            this.button_Scan.TabIndex = 4;
            this.button_Scan.Text = "浏览";
            this.button_Scan.UseVisualStyleBackColor = true;
            this.button_Scan.Click += new System.EventHandler(this.button_Scan_Click);
            // 
            // button_Comfirm
            // 
            this.button_Comfirm.Location = new System.Drawing.Point(388, 98);
            this.button_Comfirm.Name = "button_Comfirm";
            this.button_Comfirm.Size = new System.Drawing.Size(75, 23);
            this.button_Comfirm.TabIndex = 5;
            this.button_Comfirm.Text = "确定";
            this.button_Comfirm.UseVisualStyleBackColor = true;
            this.button_Comfirm.Click += new System.EventHandler(this.button_Comfirm_Click);
            // 
            // button_Return
            // 
            this.button_Return.Location = new System.Drawing.Point(501, 98);
            this.button_Return.Name = "button_Return";
            this.button_Return.Size = new System.Drawing.Size(75, 23);
            this.button_Return.TabIndex = 6;
            this.button_Return.Text = "返回";
            this.button_Return.UseVisualStyleBackColor = true;
            this.button_Return.Click += new System.EventHandler(this.button_Return_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(588, 134);
            this.ControlBox = false;
            this.Controls.Add(this.button_Return);
            this.Controls.Add(this.button_Comfirm);
            this.Controls.Add(this.button_Scan);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Name = "Form2";
            this.Text = "添加数据";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button_Scan;
        private System.Windows.Forms.Button button_Comfirm;
        private System.Windows.Forms.Button button_Return;
    }
}