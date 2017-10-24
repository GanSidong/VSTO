namespace WordAddIn
{
    partial class AddBookMarkForm
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
            this.tbToolTip = new System.Windows.Forms.TextBox();
            this.chHighlightColor = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "Tooltip:";
            // 
            // tbToolTip
            // 
            this.tbToolTip.Location = new System.Drawing.Point(15, 28);
            this.tbToolTip.Multiline = true;
            this.tbToolTip.Name = "tbToolTip";
            this.tbToolTip.Size = new System.Drawing.Size(257, 145);
            this.tbToolTip.TabIndex = 1;
            // 
            // chHighlightColor
            // 
            this.chHighlightColor.AutoSize = true;
            this.chHighlightColor.Location = new System.Drawing.Point(15, 179);
            this.chHighlightColor.Name = "chHighlightColor";
            this.chHighlightColor.Size = new System.Drawing.Size(72, 16);
            this.chHighlightColor.TabIndex = 2;
            this.chHighlightColor.Text = "添加高亮";
            this.chHighlightColor.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(46, 227);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 3;
            this.button1.Text = "取消";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.cancel_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(167, 227);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 4;
            this.button2.Text = "确定";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.sure_Click);
            // 
            // AddBookMarkForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.chHighlightColor);
            this.Controls.Add(this.tbToolTip);
            this.Controls.Add(this.label1);
            this.Name = "AddBookMarkForm";
            this.Text = "AddBookMarkForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbToolTip;
        private System.Windows.Forms.CheckBox chHighlightColor;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}