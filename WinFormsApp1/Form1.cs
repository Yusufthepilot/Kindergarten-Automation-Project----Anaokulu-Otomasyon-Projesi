using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using ClosedXML.Excel;
using System.IO;
using System.Data.OleDb;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        SqlCommand komut;
        SqlConnection baglanti;
        SqlDataReader oku;
        int ogrenci_id=0;
        int veli_id=0;
        int ogretmen_id=0;
        int personel_id=0;
        int sinif_id=0;
        int user_id=0;

        public Form1()
        {
            InitializeComponent();
            CenterToScreen();
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
        }

        private void ogrenci_kaydet_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=True");
            komut = new SqlCommand("Insert Into Ogrenci (tckn, ad, soyad, yas) values (@TC, @AD, @SOYAD, @YAS)", baglanti);
            komut.Parameters.AddWithValue("@TC", ogrenci_tc.Text);
            komut.Parameters.AddWithValue("@AD", ogrenci_ad.Text);
            komut.Parameters.AddWithValue("@SOYAD", ogrenci_soyad.Text);
            komut.Parameters.AddWithValue("@YAS", ogrenci_yas.Text);

            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                komut.ExecuteNonQuery();
                baglanti.Close();
                yenile1();
                temizle1();
                sinif_ogrenci_doldur();
                veli_ogrenci_doldur();
                MessageBox.Show("Öğrenci Kayıt Edildi");
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Tckn"))
                {
                    MessageBox.Show("Hatalı kimlik numarası girişi.");
                }
                else if (ex.Message.Contains("Conversion"))
                {
                    MessageBox.Show("Hatalı giriş yapıldı.");
                }
                else if (ex.Message.Contains("UQ"))
                {
                    MessageBox.Show("Öğrenci mevcut!");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void ogrenci_sil_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=True");
            komut = new SqlCommand("Delete From Ogrenci where OgrenciID=@id", baglanti);
            komut.Parameters.AddWithValue("@id", ogrenci_id );

            try
            {
                if (ogrenci_id != 0)
                {
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    oku = komut.ExecuteReader();
                    baglanti.Close();
                    yenile1();
                    yenile2();
                    yenile3();
                    yenile4();
                    temizle1();
                    sinif_ogrenci_doldur();
                    SinifDoldur();
                    MessageBox.Show("Öğrenci Silindi");
                    ogrenci_id = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ogrenci_guncelle_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Update Ogrenci set tckn=@TC, ad=@AD, soyad=@SOYAD, yas=@YAS where OgrenciID=@id", baglanti);
            komut.Parameters.AddWithValue("@id", ogrenci_id);
            komut.Parameters.AddWithValue("@TC", ogrenci_tc.Text);
            komut.Parameters.AddWithValue("@AD", ogrenci_ad.Text);
            komut.Parameters.AddWithValue("@SOYAD", ogrenci_soyad.Text);
            komut.Parameters.AddWithValue("@YAS", ogrenci_yas.Text);

            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                komut.ExecuteNonQuery();
                baglanti.Close();
                yenile1();
                temizle1();
                sinif_ogrenci_doldur();
                MessageBox.Show("Öğrenci Güncellendi");
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Tckn"))
                {
                    MessageBox.Show("Hatalı kimlik numarası girişi.");
                }
                else if (ex.Message.Contains("Conversion"))
                {
                    MessageBox.Show("Hatalı giriş yapıldı.");
                }
                else if (ex.Message.Contains("UQ"))
                {
                    MessageBox.Show("Öğrenci mevcut!");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private void yenile1()
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("select * From Ogrenci", baglanti);
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                DataTable tablo1 = new DataTable();
                tablo1.Load(oku);
                dataGridView1.DataSource = tablo1;
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void temizle1()
        {
            ogrenci_tc.Text = "";
            ogrenci_ad.Text = "";
            ogrenci_soyad.Text = "";
            ogrenci_yas.Text = "";

        }
        private void ogrenci_temizle_Click(object sender, EventArgs e)
        {
            temizle1();
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int rowIndex = e.RowIndex;
            DataGridViewRow row = dataGridView1.Rows[rowIndex];
            ogrenci_id = int.Parse(row.Cells[0].Value.ToString());
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Select * From Ogrenci where OgrenciID=@id", baglanti);
            komut.Parameters.AddWithValue("@id", ogrenci_id);


            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                oku.Read();
                ogrenci_tc.Text = oku["tckn"].ToString();
                ogrenci_ad.Text = oku["ad"].ToString();
                ogrenci_soyad.Text = oku["soyad"].ToString();
                ogrenci_yas.Text = oku["yas"].ToString();
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void veli_kaydet_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Insert Into Veli (ogrenciID, tckn, ad, soyad, telefon, email) values (@OgrenciID, @TC, @AD, @SOYAD, @TELEFON, @EMAIL)", baglanti);
            komut.Parameters.AddWithValue("@OgrenciID", veli_ogrenci.SelectedValue);
            komut.Parameters.AddWithValue("@TC", veli_tc.Text);
            komut.Parameters.AddWithValue("@AD", veli_ad.Text);
            komut.Parameters.AddWithValue("@SOYAD", veli_soyad.Text);
            komut.Parameters.AddWithValue("@TELEFON", veli_tel.Text);
            komut.Parameters.AddWithValue("@EMAIL", veli_mail.Text);

            try
            {
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    komut.ExecuteNonQuery();
                    baglanti.Close();
                    yenile2();
                    temizle2();
                    MessageBox.Show("Veli Kayıt Edildi");
                    veli_id = 0;
                
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Tckn"))
                {
                    MessageBox.Show("Hatalı kimlik numarası girişi.");
                }
                else if (ex.Message.Contains("Conversion"))
                {
                    MessageBox.Show("Hatalı giriş yapıldı.");
                }
                else if (ex.Message.Contains("Telefon"))
                {
                    MessageBox.Show("Telefon numarası eksik veya hatalı girildi.");
                }
                else if (ex.Message.Contains("eMail"))
                {
                    MessageBox.Show("E-posta adresi eksik veya hatalı girildi.");
                }
                else if (ex.Message.Contains("UQ"))
                {
                    MessageBox.Show("Veli mevcut!");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private void veli_guncelle_Click(object sender, EventArgs e)
        {

            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Update Veli set ogrenciID=@OgrenciID, tckn=@TC, ad=@AD, soyad=@SOYAD, telefon=@TELEFON, email=@EMAIL where KayitID=@id", baglanti);
            komut.Parameters.AddWithValue("@OgrenciID", veli_ogrenci.SelectedValue);
            komut.Parameters.AddWithValue("@id", veli_id);
            komut.Parameters.AddWithValue("@TC", veli_tc.Text);
            komut.Parameters.AddWithValue("@AD", veli_ad.Text);
            komut.Parameters.AddWithValue("@SOYAD", veli_soyad.Text);
            komut.Parameters.AddWithValue("@TELEFON", veli_tel.Text);
            komut.Parameters.AddWithValue("@EMAIL", veli_mail.Text);

            try
            {
                if (veli_id != 0)
                {
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    komut.ExecuteNonQuery();
                    baglanti.Close();
                    yenile2();
                    temizle2();
                    MessageBox.Show("Veli Güncellendi");
                    veli_id = 0;
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Tckn"))
                {
                    MessageBox.Show("Hatalı kimlik numarası girişi.");
                }
                else if (ex.Message.Contains("Conversion"))
                {
                    MessageBox.Show("Hatalı giriş yapıldı.");
                }
                else if (ex.Message.Contains("Telefon"))
                {
                    MessageBox.Show("Telefon numarası eksik veya hatalı girildi.");
                }
                else if (ex.Message.Contains("eMail"))
                {
                    MessageBox.Show("E-posta adresi eksik veya hatalı girildi.");
                }
                else if (ex.Message.Contains("UQ"))
                {
                    MessageBox.Show("Veli mevcut!");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void veli_sil_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Delete From Veli where KayitID=@id", baglanti);
            komut.Parameters.AddWithValue("@id", veli_id);

            try
            {
                if (veli_id != 0)
                {
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    oku = komut.ExecuteReader();
                    baglanti.Close();
                    baglanti.Close();
                    yenile2();
                    temizle2();
                    MessageBox.Show("Veli Silindi");
                    veli_id = 0;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void veli_temizle_Click(object sender, EventArgs e)
        {
            temizle2();
        }

        private void dataGridView2_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int rowIndex = e.RowIndex;
            DataGridViewRow row = dataGridView2.Rows[rowIndex];
            veli_id = int.Parse(row.Cells[0].Value.ToString());
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Select * From Veli where KayitID=@id", baglanti);
            komut.Parameters.AddWithValue("@id", veli_id);


            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                oku.Read();
                veli_tc.Text = oku["tckn"].ToString();
                veli_ad.Text = oku["ad"].ToString();
                veli_soyad.Text = oku["soyad"].ToString();
                veli_tel.Text = oku["telefon"].ToString();
                veli_mail.Text = oku["eMail"].ToString();
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void temizle2()
        {
            veli_tc.Text = "";
            veli_ad.Text = "";
            veli_soyad.Text = "";
            veli_tel.Text = "";
            veli_mail.Text = "";
        }
        private void yenile2()
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("select * From Veli", baglanti);
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                DataTable tablo2 = new DataTable();
                tablo2.Load(oku);
                dataGridView2.DataSource = tablo2;
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void ogretmen_kaydet_Click(object sender, EventArgs e)
        {

            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Insert Into Ogretmen (tckn, ad, soyad, telefon, eMail, Maas, BaslamaTarihi) values (@TC, @AD, @SOYAD, @TELEFON, @EMAIL, @MAAS, @BASLAMATARIHI)", baglanti);
            komut.Parameters.AddWithValue("@TC", ogretmen_tc.Text);
            komut.Parameters.AddWithValue("@AD", ogretmen_ad.Text);
            komut.Parameters.AddWithValue("@SOYAD", ogretmen_soyad.Text);
            komut.Parameters.AddWithValue("@TELEFON", ogretmen_tel.Text);
            komut.Parameters.AddWithValue("@EMAIL", ogretmen_mail.Text);
            komut.Parameters.AddWithValue("@MAAS", ogretmen_maas.Text);
            komut.Parameters.AddWithValue("@BASLAMATARIHI", dtpOgretmen.Value);

            try
            {
                if(ogretmen_id != 0)
                {
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    komut.ExecuteNonQuery();
                    baglanti.Close();
                    yenile3();
                    temizle3();
                    sinif_ogretmen_doldur();
                    MessageBox.Show("Öğretmen Kayıt Edildi");
                    ogretmen_id = 0;
                }
                
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Tckn"))
                {
                    MessageBox.Show("Hatalı kimlik numarası girişi.");
                }
                else if (ex.Message.Contains("Conversion")|| ex.Message.Contains("convert"))
                {
                    MessageBox.Show("Hatalı giriş yapıldı.");
                }
                else if (ex.Message.Contains("Telefon"))
                {
                    MessageBox.Show("Telefon numarası eksik veya hatalı girildi.");
                }
                else if (ex.Message.Contains("eMail"))
                {
                    MessageBox.Show("E-posta adresi eksik veya hatalı girildi.");
                }
                else if (ex.Message.Contains("UQ"))
                {
                    MessageBox.Show("Öğretmen mevcut!");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void ogretmen_sil_Click(object sender, EventArgs e)
        {

            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Delete From Ogretmen where OgretmenID=@id", baglanti);
            komut.Parameters.AddWithValue("@id", ogretmen_id);

            try
            {
                if(ogretmen_id != 0)
                {
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    oku = komut.ExecuteReader();
                    baglanti.Close();
                    yenile3();
                    temizle3();
                    sinif_ogretmen_doldur();
                    MessageBox.Show("Öğretmen Silindi");
                    ogretmen_id = 0;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void ogretmen_guncelle_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Update Ogretmen set tckn=@TC, ad=@AD, soyad=@SOYAD, telefon=@TELEFON, email=@MAIL, maas=@MAAS, baslamatarihi=@BASLAMATARIHI where OgretmenID=@id", baglanti);
            komut.Parameters.AddWithValue("@id", ogretmen_id);
            komut.Parameters.AddWithValue("@TC", ogretmen_tc.Text);
            komut.Parameters.AddWithValue("@AD", ogretmen_ad.Text);
            komut.Parameters.AddWithValue("@SOYAD", ogretmen_soyad.Text);
            komut.Parameters.AddWithValue("@TELEFON", ogretmen_tel.Text);
            komut.Parameters.AddWithValue("@MAIL", ogretmen_mail.Text);
            komut.Parameters.AddWithValue("@MAAS", Convert.ToDecimal(ogretmen_maas.Text));
            komut.Parameters.AddWithValue("@BASLAMATARIHI", dtpOgretmen.Value);

            try
            {
                if (ogretmen_id != 0)
                {
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    komut.ExecuteNonQuery();
                    baglanti.Close();
                    yenile3();
                    temizle3();
                    sinif_ogretmen_doldur();
                    MessageBox.Show("Öğretmen Güncellendi");
                    ogretmen_id = 0;
                }
                
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Tckn"))
                {
                    MessageBox.Show("Hatalı kimlik numarası girişi.");
                }
                else if (ex.Message.Contains("Conversion") || ex.Message.Contains("convert"))
                {
                    MessageBox.Show("Hatalı giriş yapıldı.");
                }
                else if (ex.Message.Contains("Telefon"))
                {
                    MessageBox.Show("Telefon numarası eksik veya hatalı girildi.");
                }
                else if (ex.Message.Contains("eMail"))
                {
                    MessageBox.Show("E-posta adresi eksik veya hatalı girildi.");
                }
                else if (ex.Message.Contains("UQ"))
                {
                    MessageBox.Show("Öğretmen mevcut!");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void dataGridView3_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            int rowIndex = e.RowIndex;
            DataGridViewRow row = dataGridView3.Rows[rowIndex];
            ogretmen_id = int.Parse(row.Cells[0].Value.ToString());
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Select * From Ogretmen where OgretmenID=@id", baglanti);
            komut.Parameters.AddWithValue("@id", ogretmen_id);

            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                oku.Read();
                ogretmen_tc.Text = oku["tckn"].ToString();
                ogretmen_ad.Text = oku["ad"].ToString();
                ogretmen_soyad.Text = oku["soyad"].ToString();
                ogretmen_tel.Text = oku["telefon"].ToString();
                ogretmen_mail.Text = oku["eMail"].ToString();
                ogretmen_maas.Text = oku["maas"].ToString();
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);  
            }
        }
        private void temizle3()
        {
            ogretmen_tc.Text = "";
            ogretmen_ad.Text = "";
            ogretmen_soyad.Text = "";
            ogretmen_tel.Text = "";
            ogretmen_mail.Text = "";
            ogretmen_maas.Text = "";
        }

        private void yenile3()
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("select * From Ogretmen", baglanti);
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                DataTable tablo3 = new DataTable();
                tablo3.Load(oku);
                dataGridView3.DataSource = tablo3;
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ogretmen_temizle_Click(object sender, EventArgs e)
        {
            temizle3();
        }
        private void personel_kaydet_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Insert Into Yonetim (tckn, ad, soyad, pozisyon, telefon, eMail, Maas ,BaslamaTarihi) values (@TC, @AD, @SOYAD, @POZISYON, @TELEFON, @EMAIL, @MAAS, @BASLAMATARIHI)", baglanti);
            komut.Parameters.AddWithValue("@TC", personel_tc.Text);
            komut.Parameters.AddWithValue("@AD", personel_ad.Text);
            komut.Parameters.AddWithValue("@SOYAD", personel_soyad.Text);
            komut.Parameters.AddWithValue("@POZISYON", cbxPersonel.SelectedValue); 
            komut.Parameters.AddWithValue("@TELEFON", personel_tel.Text);
            komut.Parameters.AddWithValue("@EMAIL", personel_mail.Text);
            komut.Parameters.AddWithValue("@MAAS", personel_maas.Text);
            komut.Parameters.AddWithValue("@BASLAMATARIHI", dtpPersonel.Value);

            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                    komut.ExecuteNonQuery();
                    MessageBox.Show("Personel Kayıt Edildi");
                }
                baglanti.Close();
                yenile4();
                temizle4();
                MessageBox.Show("Personel Kayıt Edildi");

            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Tckn"))
                {
                    MessageBox.Show("Hatalı kimlik numarası girişi.");
                }
                else if (ex.Message.Contains("Conversion") || ex.Message.Contains("convert"))
                {
                    MessageBox.Show("Hatalı giriş yapıldı.");
                }
                else if (ex.Message.Contains("Telefon"))
                {
                    MessageBox.Show("Telefon numarası eksik veya hatalı girildi.");
                }
                else if (ex.Message.Contains("eMail"))
                {
                    MessageBox.Show("E-posta adresi eksik veya hatalı girildi.");
                }
                else if (ex.Message.Contains("UQ"))
                {
                    MessageBox.Show("Personel mevcut!");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void temizle4()
        {
            personel_tc.Text = "";
            personel_ad.Text = "";
            personel_soyad.Text = "";
            personel_tel.Text = "";
            personel_mail.Text = "";
            personel_maas.Text = "";
        }

        private void yenile4()
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("exec sp_PersonelListele", baglanti);
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                DataTable tablo4 = new DataTable();
                tablo4.Load(oku);
                dataGridView4.DataSource = tablo4;
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void personel_sil_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Delete From Yonetim where PersonelID=@id", baglanti);
            komut.Parameters.AddWithValue("@id", personel_id);

            try
            {
                if (personel_id != 0)
                {
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    oku = komut.ExecuteReader();
                    baglanti.Close();
                    baglanti.Close();
                    yenile4();
                    temizle4();
                    MessageBox.Show("Personel Silindi");
                    personel_id = 0;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void personel_guncelle_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Update Yonetim set tckn=@TC, ad=@AD, soyad=@SOYAD, pozisyon=@POZISYON, telefon=@TELEFON, email=@EMAIL, maas=@MAAS, baslamatarihi=@BASLAMATARIHI where PersonelID=@id", baglanti);
            komut.Parameters.AddWithValue("@id", personel_id);
            komut.Parameters.AddWithValue("@TC", personel_tc.Text);
            komut.Parameters.AddWithValue("@AD", personel_ad.Text);
            komut.Parameters.AddWithValue("@SOYAD", personel_soyad.Text);
            komut.Parameters.AddWithValue("@POZISYON", cbxPersonel.SelectedValue);
            komut.Parameters.AddWithValue("@TELEFON", personel_tel.Text);
            komut.Parameters.AddWithValue("@EMAIL", personel_mail.Text);
            komut.Parameters.AddWithValue("@MAAS", Convert.ToDecimal(personel_maas.Text));
            komut.Parameters.AddWithValue("@BASLAMATARIHI", dtpPersonel.Value);

            try
            {
                if(personel_id != 0)
                {
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    komut.ExecuteNonQuery();
                    baglanti.Close();
                    yenile4();
                    temizle4();
                    MessageBox.Show("Personel Güncellendi");
                    personel_id = 0;
                }
                
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Tckn"))
                {
                    MessageBox.Show("Hatalı kimlik numarası girişi.");
                }
                else if (ex.Message.Contains("Conversion") || ex.Message.Contains("convert"))
                {
                    MessageBox.Show("Hatalı giriş yapıldı.");
                }
                else if (ex.Message.Contains("Telefon"))
                {
                    MessageBox.Show("Telefon numarası eksik veya hatalı girildi.");
                }
                else if (ex.Message.Contains("eMail"))
                {
                    MessageBox.Show("E-posta adresi eksik veya hatalı girildi.");
                }
                else if (ex.Message.Contains("UQ"))
                {
                    MessageBox.Show("Personel mevcut!");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void personel_temizle_Click(object sender, EventArgs e)
        {
            temizle4();
        }

        private void dataGridView4_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = e.RowIndex;
            DataGridViewRow row = dataGridView4.Rows[rowIndex];
            personel_id = int.Parse(row.Cells[0].Value.ToString());
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Select * From Yonetim where PersonelID=@id", baglanti);
            komut.Parameters.AddWithValue("@id", personel_id);


            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                oku.Read();
                personel_tc.Text = oku["tckn"].ToString();
                personel_ad.Text = oku["ad"].ToString();
                personel_soyad.Text = oku["soyad"].ToString();
                personel_tel.Text = oku["telefon"].ToString();
                personel_mail.Text = oku["eMail"].ToString();
                personel_maas.Text = oku["maas"].ToString();
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.ReadOnly = true;
            dataGridView2.ReadOnly = true;
            dataGridView3.ReadOnly = true;
            dataGridView4.ReadOnly = true;
            dataGridView5.ReadOnly = true;
            dgwAdminUsers.ReadOnly = true;

            dataGridView1.AllowUserToAddRows = false;
            dataGridView2.AllowUserToAddRows = false;
            dataGridView3.AllowUserToAddRows = false;
            dataGridView4.AllowUserToAddRows = false;
            dataGridView5.AllowUserToAddRows = false;
            dgwAdminUsers.AllowUserToAddRows = false;

            dgwtest.Hide();
            ogretmen_ara.Hide();

            oi_id.Enabled = false;
            on_id.Enabled = false;
            veli_oi_id.Enabled = false;

            oi_id.Hide();
            on_id.Hide();
            veli_oi_id.Hide();

            //label30.Hide();
            //ogretmen_baslangic.Hide();
            //label34.Hide();
            //personel_baslangic.Hide();

            sinif_temizle.Hide();

            sinif_ogrenci_doldur();
            sinif_ogretmen_doldur();
            veli_ogrenci_doldur();
            admin_user_doldur();
            cbxPersonelDoldur();

            yenile1();
            yenile2();
            yenile3();
            yenile4();
            SinifDoldur();
            
            if (FormLogin.mode != 0)
            {
                tabControl1.TabPages.Remove(tabAdmin);
            }
            
        }

        private void veli_ogrenci_doldur()
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Select * From Ogrenci", baglanti);
            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }

            SqlDataAdapter dataAdapter = new SqlDataAdapter(komut);
            DataTable dataTable = new DataTable();
            dataAdapter.Fill(dataTable);
            dataTable.Columns.Add("FullName", typeof(string), " Ad + ' ' + Soyad ");

            veli_ogrenci.ValueMember = "OgrenciID";
            veli_ogrenci.DisplayMember = "FullName";
            veli_ogrenci.DataSource = dataTable;
        }

        private void sinif_ogretmen_doldur()
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Select * From Ogretmen", baglanti);
            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
            
            SqlDataAdapter dataAdapter = new SqlDataAdapter(komut);
            DataTable dataTable = new DataTable();
            dataAdapter.Fill(dataTable);
            dataTable.Columns.Add("FullName", typeof(string), " Ad + ' ' + Soyad ");

            sinif_ogretmen.ValueMember = "OgretmenID";
            sinif_ogretmen.DisplayMember = "FullName";
            sinif_ogretmen.DataSource = dataTable;
        }

        
        private void sinif_ogrenci_doldur()
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Select * From Ogrenci", baglanti);
            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
            
            SqlDataAdapter dataAdapter = new SqlDataAdapter(komut);
            DataTable dataTable = new DataTable();
            dataAdapter.Fill(dataTable);
            dataTable.Columns.Add("FullName", typeof(string), " Ad + ' ' + Soyad ");

            sinif_ogrenci.ValueMember = "OgrenciID";
            sinif_ogrenci.DisplayMember = "FullName";
            sinif_ogrenci.DataSource = dataTable;
        }
        private void cbxPersonelDoldur()
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Select * From PersonelPozisyon", baglanti);
            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
            
            SqlDataAdapter dataAdapter = new SqlDataAdapter(komut);
            DataTable dataTable = new DataTable();
            dataAdapter.Fill(dataTable);
            
            cbxPersonel.ValueMember = "PozisyonID";
            cbxPersonel.DisplayMember = "Pozisyon";
            cbxPersonel.DataSource = dataTable;
        }

        private void sinif_kaydet_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Insert Into Sinif(OgrenciID,OgretmenID) values(@OGRENCIID,@OGRETMENID)", baglanti);

            komut.Parameters.AddWithValue("@OGRENCIID", sinif_ogrenci.SelectedValue);
            komut.Parameters.AddWithValue("@OGRETMENID", sinif_ogretmen.SelectedValue);

            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                komut.ExecuteNonQuery();
                baglanti.Close();
                SinifDoldur();
                temizle5();
                sinif_ogrenci_doldur();
                sinif_ogretmen_doldur();
                MessageBox.Show("Sinif Bilgisi Kayıt Edildi");
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("UQ"))
                {
                    MessageBox.Show("Bir öğrenci ancak bir öğretmene atanabilir.");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void SinifDoldur()
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("exec SinifListele", baglanti);
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                DataTable tablo5 = new DataTable();
                tablo5.Load(oku);
                dataGridView5.DataSource = tablo5;
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void temizle5()
        {
            oi_id.Text = "";
            on_id.Text = "";
          
        }

        private void sinif_güncelle_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Update Sinif set OgrenciID=@OGRENCIID, OgretmenID=@OGRETMENID where KayitID=@id", baglanti);
            komut.Parameters.AddWithValue("@id", sinif_id);
            komut.Parameters.AddWithValue("@OGRENCIID", Convert.ToInt16(oi_id.Text));
            komut.Parameters.AddWithValue("@OGRETMENID", Convert.ToInt16(on_id.Text));

            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                komut.ExecuteNonQuery();
                baglanti.Close();
                SinifDoldur();
                sinif_ogretmen_doldur();
                sinif_ogrenci_doldur();
                MessageBox.Show("Sinif Güncellendi");
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("UQ"))
                {
                    MessageBox.Show("Bir öğrenci ancak bir öğretmene atanabilir.");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void sinif_sil_Click(object sender, EventArgs e)
        {

            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Delete from Sinif where KayitID=@id",baglanti);
            komut.Parameters.AddWithValue("@id", sinif_id);
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                baglanti.Close();
                SinifDoldur();
                MessageBox.Show("Sınıf Kaydı Silindi");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void sinif_temizle_Click(object sender, EventArgs e)
        {
            temizle5();
        }

        private void dataGridView5_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            int rowIndex = e.RowIndex;
            DataGridViewRow row = dataGridView5.Rows[rowIndex];
            sinif_id = int.Parse(row.Cells[0].Value.ToString());
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Select * From Sinif where KayitID=@id", baglanti);
            komut.Parameters.AddWithValue("@id", sinif_id);

            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                oku.Read();
                oi_id.Text = oku["ogrenciID"].ToString();
                on_id.Text = oku["ogretmenID"].ToString();
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            var window = MessageBox.Show("Çıkmak istediğinize emin misiniz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            e.Cancel = (window == DialogResult.No);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.OpenForms[0].Show();
        }

        private void btnAdminSave_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Insert Into GirisBilgileri (KullaniciAdi, sifre) values (@USER, @PASS)", baglanti);
            komut.Parameters.AddWithValue("@USER", tbxAdminUser.Text);
            komut.Parameters.AddWithValue("@PASS", tbxAdminPass.Text);

            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                komut.ExecuteNonQuery();
                baglanti.Close();
                AdminTemizle();
                admin_user_doldur();
                
                MessageBox.Show("Kullanıcı Kayıt Edildi");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void admin_user_doldur()
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Select * From GirisBilgileri", baglanti);
            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
            oku = komut.ExecuteReader();
            DataTable tablo = new DataTable();
            tablo.Load(oku);

            dgwAdminUsers.DataSource = tablo;
        }

        private void btnAdminUpdate_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Update GirisBilgileri set KullaniciAdi=@USER, sifre=@PASS where id=@id", baglanti);
            komut.Parameters.AddWithValue("@id", user_id);
            komut.Parameters.AddWithValue("@USER", tbxAdminUserUpdate.Text);
            komut.Parameters.AddWithValue("@PASS", tbxAdminPassUpdate.Text);

            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                komut.ExecuteNonQuery();
                baglanti.Close();
                admin_user_doldur();
                AdminTemizle();
                MessageBox.Show("Kullanıcı Güncellendi");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgwAdminUsers_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int rowIndex = e.RowIndex;
            DataGridViewRow row = dgwAdminUsers.Rows[rowIndex];
            user_id = int.Parse(row.Cells[0].Value.ToString());
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Select * From GirisBilgileri where id=@id", baglanti);
            komut.Parameters.AddWithValue("@id", user_id);


            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                oku.Read();
                tbxAdminUserUpdate.Text = oku["KullaniciAdi"].ToString();
                tbxAdminPassUpdate.Text = oku["sifre"].ToString();
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAdminUserDelete_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("Delete From GirisBilgileri where id=@id", baglanti);
            komut.Parameters.AddWithValue("@id", user_id);

            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                baglanti.Close();
                admin_user_doldur();
                AdminTemizle();
                MessageBox.Show("Kullanıcı Silindi");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void AdminTemizle()
        {
            tbxAdminPass.Text = "";
            tbxAdminUser.Text = "";
            tbxAdminUserUpdate.Text = "";
            tbxAdminPassUpdate.Text = "";
        }
       

        private void btnBackup_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            SaveFileDialog saveDialog = new SaveFileDialog();
            
            saveDialog.Filter = "SQL Server Backup File (*.bak)|*.bak";
            saveDialog.DefaultExt = "bak";
            saveDialog.AddExtension = true;

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                String fileName = saveDialog.FileName;
                komut = new SqlCommand("BACKUP DATABASE [AnaokuluDB] TO  DISK = N'" + fileName + "' WITH NOFORMAT, NOINIT,  NAME = N'AnaokuluDB-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10", baglanti);
            }

            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                    oku = komut.ExecuteReader();
                }
                baglanti.Close();
                MessageBox.Show("Yedek başarıyla oluşturuldu.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnLoadBackup_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=master;Integrated Security=true");
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "SQL Server Backup File (*.bak)|*.bak";
            openFileDialog.DefaultExt = "bak";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                String fileName = openFileDialog.FileName;
                
                try
                {
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }

                    komut = new SqlCommand("ALTER DATABASE [AnaokuluDB] SET SINGLE_USER WITH ROLLBACK IMMEDIATE", baglanti);
                    komut.ExecuteNonQuery();
                    komut = new SqlCommand("RESTORE DATABASE [AnaokuluDB] FROM  DISK = N'" + fileName + "' WITH REPLACE", baglanti);
                    komut.ExecuteNonQuery();
                    komut = new SqlCommand("ALTER DATABASE [AnaokuluDB] SET MULTI_USER", baglanti);
                    komut.ExecuteNonQuery();

                    baglanti.Close();
                    MessageBox.Show("Yedekten dönüldü.");
                    
                    yenile1();
                    yenile2();
                    yenile3();
                    yenile4();
                    admin_user_doldur();
                    sinif_ogrenci_doldur();
                    sinif_ogretmen_doldur();
                    veli_ogrenci_doldur();
                    cbxPersonelDoldur();
                    SinifDoldur();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    komut = new SqlCommand("ALTER DATABASE [AnaokuluDB] SET MULTI_USER", baglanti);
                    komut.ExecuteNonQuery();
                }
            }
        }

        private void btnExcelOutOgrenci_Click(object sender, EventArgs e)
        {
            SqlConnection connection = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            SqlCommand command = new SqlCommand("Select * from Ogrenci", connection);

            try
            {
                connection.Open();
                DataTable data = new DataTable();

                using (SqlDataAdapter dataAdapter = new SqlDataAdapter(command))
                {
                    dataAdapter.Fill(data);
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add(data, "Excel Export"); 
                        
                        SaveFileDialog saveDialog = new SaveFileDialog();
                        saveDialog.Filter = "Excel Dosyası(*.xlsx)|*.xlsx";
                        saveDialog.DefaultExt = "xlsx";
                        saveDialog.AddExtension = true;

                        if (saveDialog.ShowDialog() == DialogResult.OK)
                        {
                            String fileName = saveDialog.FileName;
                            wb.SaveAs(fileName);
                            MessageBox.Show("Kayıtlar dışa aktarıldı.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void ImportDataFromExcel(string excelFilePath)
        {
            try
            {
                OleDbConnection oleDbConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;data source=" + excelFilePath + ";extended properties=" + "\"excel 8.0;hdr=yes;\"");
                oleDbConnection.Open();
                OleDbCommand oleDbCommand = new OleDbCommand("Select * from [Excel Export$]", oleDbConnection);
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter();
                dataAdapter.SelectCommand = oleDbCommand;
                DataTable dataTable = new DataTable();
                dataTable.Clear();
                dataAdapter.Fill(dataTable);
                dgwtest.DataSource = dataTable;
                oleDbConnection.Close();

                SqlConnection connection = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=True");

                foreach (DataGridViewRow row in dgwtest.Rows)
                {
                    yenile1();

                    SqlCommand command = new SqlCommand("Insert into Ogrenci (Tckn, Ad, Soyad, Yas) values (@tckn, @ad, @soyad, @yas)", connection);
                    command.Parameters.AddWithValue("tckn", row.Cells["Tckn"].Value.ToString());
                    command.Parameters.AddWithValue("ad", row.Cells["Ad"].Value.ToString());
                    command.Parameters.AddWithValue("soyad", row.Cells["Soyad"].Value.ToString());
                    command.Parameters.AddWithValue("yas", row.Cells["Yas"].Value);
                    
                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                }
                
                MessageBox.Show("Kayıtlar içeri aktarıldı.");
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("UQ"))
                {
                    MessageBox.Show("Kayıt tekrarı mevcut. Lütfen girdi bilgilerini kontrol ediniz.");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void btnImportTest_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            var result = openFileDialog.ShowDialog();

            if(result == DialogResult.OK)
            {
                ImportDataFromExcel(openFileDialog.FileName);
            }
        }

        private void veli_ogrenci_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void ogrenci_listele_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=True");
            komut = new SqlCommand("select * From Ogrenci Where Ad like '%" + ogrenci_ad.Text + "%' and Soyad like '%" + ogrenci_soyad.Text + "%'and Yas like '%" + ogrenci_yas.Text + "%'", baglanti);

            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                DataTable tablo1 = new DataTable();
                tablo1.Load(oku);
                dataGridView1.DataSource = tablo1;
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void veli_listele_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("select * From Veli Where Ad like '%" + veli_ad.Text + "%' and Soyad like'%" + veli_soyad.Text + "%'", baglanti);
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                DataTable tablo2 = new DataTable();
                tablo2.Load(oku);
                dataGridView2.DataSource = tablo2;
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void veli_ara_Click(object sender, EventArgs e)
        {
            {
                baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
                komut = new SqlCommand("select * From Veli Where  OgrenciID like '%" + veli_ogrenci.SelectedValue + "%'", baglanti);
            try
                {
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    oku = komut.ExecuteReader();
                    DataTable tablo2 = new DataTable();
                    tablo2.Load(oku);
                    dataGridView2.DataSource = tablo2;
                    baglanti.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void ogretmen_listele_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("select * From Ogretmen Where Ad like '%" + ogretmen_ad.Text + "%' and Soyad like '%" + ogretmen_soyad.Text + "%' and Maas like '%" + ogretmen_maas.Text + "%'", baglanti);
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                DataTable tablo3 = new DataTable();
                tablo3.Load(oku);
                dataGridView3.DataSource = tablo3;
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ogretmen_ara_Click(object sender, EventArgs e)
        {
            //baglanti = new SqlConnection("Data Source=DESKTOP-AEL6S94;Initial Catalog=AnaokuluDB;Integrated Security=true");
            //komut = new SqlCommand("select * From Ogretmen Where BaslamaTarihi  <=" + dtpOgretmen.Value.Date.ToString("ddMMyyyy") + ">= getdate()" ,  baglanti);
            //try
            //{
            //    if (baglanti.State == ConnectionState.Closed)
            //    {
            //        baglanti.Open();
            //    }
            //    oku = komut.ExecuteReader();
            //    DataTable tablo3 = new DataTable();
            //    tablo3.Load(oku);
            //    dataGridView3.DataSource = tablo3;
            //    baglanti.Close();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }

        private void personel_listele_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("select * From Yonetim Where Ad Like '%" + personel_ad.Text + "%' and Soyad Like '%" + personel_soyad.Text + "%' and Maas Like '%" + personel_maas.Text + "%'", baglanti);
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                DataTable tablo4 = new DataTable();
                tablo4.Load(oku);
                dataGridView4.DataSource = tablo4;
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void personel_Ara_Click(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=DESKTOP-88PI71V;Initial Catalog=AnaokuluDB;Integrated Security=true");
            komut = new SqlCommand("select * From Yonetim Where Pozisyon like '%" + cbxPersonel.SelectedValue + "%'", baglanti);
            
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                oku = komut.ExecuteReader();
                DataTable tablo4 = new DataTable();
                tablo4.Load(oku);
                dataGridView4.DataSource = tablo4;
                baglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnOgrenciExport_Click(object sender, EventArgs e)
        {
            DataTable dtGridSource = (DataTable)dataGridView1.DataSource;

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dtGridSource, "Öğrenci");

                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel Dosyası(*.xlsx)|*.xlsx";
                saveDialog.DefaultExt = "xlsx";
                saveDialog.AddExtension = true;

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    String fileName = saveDialog.FileName;
                    wb.SaveAs(fileName);
                    MessageBox.Show("Kayıtlar dışa aktarıldı.");
                }
            }
        }

        private void btnVeliExport_Click(object sender, EventArgs e)
        {
            DataTable dtGridSource = (DataTable)dataGridView2.DataSource;

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dtGridSource, "Veli");

                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel Dosyası(*.xlsx)|*.xlsx";
                saveDialog.DefaultExt = "xlsx";
                saveDialog.AddExtension = true;

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    String fileName = saveDialog.FileName;
                    wb.SaveAs(fileName);
                    MessageBox.Show("Kayıtlar dışa aktarıldı.");
                }
            }
        }

        private void btnOgretmenExport_Click(object sender, EventArgs e)
        {
            DataTable dtGridSource = (DataTable)dataGridView3.DataSource;

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dtGridSource, "Öğretmen");

                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel Dosyası(*.xlsx)|*.xlsx";
                saveDialog.DefaultExt = "xlsx";
                saveDialog.AddExtension = true;

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    String fileName = saveDialog.FileName;
                    wb.SaveAs(fileName);
                    MessageBox.Show("Kayıtlar dışa aktarıldı.");
                }
            }
        }

        private void btnPersonelExport_Click(object sender, EventArgs e)
        {
            DataTable dtGridSource = (DataTable)dataGridView4.DataSource;

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dtGridSource, "Personel");

                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel Dosyası(*.xlsx)|*.xlsx";
                saveDialog.DefaultExt = "xlsx";
                saveDialog.AddExtension = true;

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    String fileName = saveDialog.FileName;
                    wb.SaveAs(fileName);
                    MessageBox.Show("Kayıtlar dışa aktarıldı.");
                }
            }
        }
    }
}