
namespace WinFormsApp1
{
    partial class FormLogin
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
            this.btnLogin = new System.Windows.Forms.Button();
            this.lblKullanici = new System.Windows.Forms.Label();
            this.lblPass = new System.Windows.Forms.Label();
            this.tbxPass = new System.Windows.Forms.TextBox();
            this.tbxUser = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnLogin
            // 
            this.btnLogin.Location = new System.Drawing.Point(121, 309);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(94, 29);
            this.btnLogin.TabIndex = 0;
            this.btnLogin.Text = "Giriş";
            this.btnLogin.UseVisualStyleBackColor = true;
            this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
            // 
            // lblKullanici
            // 
            this.lblKullanici.AutoSize = true;
            this.lblKullanici.Location = new System.Drawing.Point(73, 146);
            this.lblKullanici.Name = "lblKullanici";
            this.lblKullanici.Size = new System.Drawing.Size(92, 20);
            this.lblKullanici.TabIndex = 2;
            this.lblKullanici.Text = "Kullanıcı Adı";
            // 
            // lblPass
            // 
            this.lblPass.AutoSize = true;
            this.lblPass.Location = new System.Drawing.Point(73, 223);
            this.lblPass.Name = "lblPass";
            this.lblPass.Size = new System.Drawing.Size(39, 20);
            this.lblPass.TabIndex = 3;
            this.lblPass.Text = "Şifre";
            // 
            // tbxPass
            // 
            this.tbxPass.Location = new System.Drawing.Point(73, 246);
            this.tbxPass.Name = "tbxPass";
            this.tbxPass.PasswordChar = '●';
            this.tbxPass.Size = new System.Drawing.Size(196, 27);
            this.tbxPass.TabIndex = 4;
            this.tbxPass.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tbxPass_KeyDown);
            // 
            // tbxUser
            // 
            this.tbxUser.Location = new System.Drawing.Point(73, 169);
            this.tbxUser.Name = "tbxUser";
            this.tbxUser.Size = new System.Drawing.Size(196, 27);
            this.tbxUser.TabIndex = 5;
            this.tbxUser.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tbxUser_KeyDown);
            // 
            // FormLogin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(356, 411);
            this.Controls.Add(this.tbxUser);
            this.Controls.Add(this.tbxPass);
            this.Controls.Add(this.lblPass);
            this.Controls.Add(this.lblKullanici);
            this.Controls.Add(this.btnLogin);
            this.Name = "FormLogin";
            this.Text = "Anaokulu";
            this.Load += new System.EventHandler(this.FormLogin_Load);
            this.VisibleChanged += new System.EventHandler(this.FormLogin_VisibleChanged);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnLogin;
        private System.Windows.Forms.Label lblKullanici;
        private System.Windows.Forms.Label lblPass;
        private System.Windows.Forms.TextBox tbxPass;
        private System.Windows.Forms.TextBox tbxUser;
    }
}