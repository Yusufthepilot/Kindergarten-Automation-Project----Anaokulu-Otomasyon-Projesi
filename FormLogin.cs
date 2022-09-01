using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace WinFormsApp1
{
    public partial class FormLogin : Form
    {
        public static int mode;

        public FormLogin()
        {
            InitializeComponent();
            CenterToScreen();
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            SqlConnection connection =  new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=True");
            SqlCommand command = new SqlCommand("select count (*) as cnt from GirisBilgileri where KullaniciAdi=@usr and sifre=@pwd", connection);

            command.Parameters.AddWithValue("@usr", tbxUser.Text);
            command.Parameters.AddWithValue("@pwd", tbxPass.Text);

            try
            {
                connection.Open();
                if (command.ExecuteScalar().ToString() == "1")
                {
                    MessageBox.Show("Giriş Başarılı");
                    Form1 formAnaokulu = new Form1();

                    if (tbxUser.Text == "admin")
                    {
                        mode = 0;
                    }
                    else
                    {
                        mode = 1;
                    }

                    this.Hide();
                    formAnaokulu.Show();
                }
                else
                {
                    MessageBox.Show("Kullanıcı adı veya şifre hatalı. Lütfen tekrar deneyiniz.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            connection.Close();

            tbxUser.Text = "";
            tbxPass.Text = "";
        }

        private void FormLogin_Load(object sender, EventArgs e)
        {
            this.ActiveControl = tbxUser;
        }

        private void tbxUser_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnLogin.PerformClick();
            }
        }

        private void tbxPass_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnLogin.PerformClick();
            }
        }

        private void FormLogin_VisibleChanged(object sender, EventArgs e)
        {
            tbxUser.Focus();
        }
    }
}
