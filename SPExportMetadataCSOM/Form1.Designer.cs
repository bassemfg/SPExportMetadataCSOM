namespace SPExportMetadataCSOM
{
    partial class Form1
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
            this.txtUID = new System.Windows.Forms.TextBox();
            this.txtPwd = new System.Windows.Forms.TextBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtSiteURL = new System.Windows.Forms.TextBox();
            this.txtOutput = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // txtUID
            // 
            this.txtUID.Location = new System.Drawing.Point(111, 24);
            this.txtUID.Name = "txtUID";
            this.txtUID.Size = new System.Drawing.Size(282, 20);
            this.txtUID.TabIndex = 0;
            this.txtUID.Text = "spmigration@tampaairport.com";
            // 
            // txtPwd
            // 
            this.txtPwd.Location = new System.Drawing.Point(111, 70);
            this.txtPwd.Name = "txtPwd";
            this.txtPwd.PasswordChar = '*';
            this.txtPwd.Size = new System.Drawing.Size(282, 20);
            this.txtPwd.TabIndex = 1;
            this.txtPwd.Text = "BoxMigration@018";
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(27, 189);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 23);
            this.btnRun.TabIndex = 3;
            this.btnRun.Text = "Generate";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Username";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(24, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Password";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 123);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(61, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Site URL(s)";
            // 
            // txtSiteURL
            // 
            this.txtSiteURL.Location = new System.Drawing.Point(111, 116);
            this.txtSiteURL.Multiline = true;
            this.txtSiteURL.Name = "txtSiteURL";
            this.txtSiteURL.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtSiteURL.Size = new System.Drawing.Size(282, 66);
            this.txtSiteURL.TabIndex = 2;
            this.txtSiteURL.Text = "https://tampaairport.sharepoint.com/sites/MasterPlan";
            // 
            // txtOutput
            // 
            this.txtOutput.Location = new System.Drawing.Point(27, 218);
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.Size = new System.Drawing.Size(366, 220);
            this.txtOutput.TabIndex = 6;
            this.txtOutput.Text = "";
            // 
            // Form1
            // 
            this.AcceptButton = this.btnRun;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(418, 450);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.txtSiteURL);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.txtPwd);
            this.Controls.Add(this.txtUID);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtUID;
        private System.Windows.Forms.TextBox txtPwd;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtSiteURL;
        private System.Windows.Forms.RichTextBox txtOutput;
    }
}

