namespace WordAddIn
{
    partial class FloatingPanel
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
            this.tbTooltip = new System.Windows.Forms.TextBox();
            this.lkBtn = new System.Windows.Forms.LinkLabel();
            this.SuspendLayout();
            // 
            // tbTooltip
            // 
            this.tbTooltip.BackColor = System.Drawing.SystemColors.Menu;
            this.tbTooltip.Location = new System.Drawing.Point(12, 12);
            this.tbTooltip.Multiline = true;
            this.tbTooltip.Name = "tbTooltip";
            this.tbTooltip.Size = new System.Drawing.Size(212, 161);
            this.tbTooltip.TabIndex = 0;
            // 
            // lkBtn
            // 
            this.lkBtn.AutoSize = true;
            this.lkBtn.Location = new System.Drawing.Point(159, 176);
            this.lkBtn.Name = "lkBtn";
            this.lkBtn.Size = new System.Drawing.Size(41, 12);
            this.lkBtn.TabIndex = 1;
            this.lkBtn.TabStop = true;
            this.lkBtn.Text = "remove";
            this.lkBtn.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lkBtn_LinkClicked);
            // 
            // FloatingPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(236, 197);
            this.Controls.Add(this.lkBtn);
            this.Controls.Add(this.tbTooltip);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "FloatingPanel";
            this.Text = "FloatingPanel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbTooltip;
        private System.Windows.Forms.LinkLabel lkBtn;
    }
}