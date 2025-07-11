using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using OfficeOpenXml;
using System.IO;
using System.Drawing.Imaging;

namespace Lab_Deneme
{
    public partial class MainForm : Form
    {
        private bool isAdmin;
        private bool isYonetici;  // Yönetici rolüne sahip mi kontrol edecek

        private bool isLotWarningShown = false; // Geçerli seçimler yapıldığında uyarıyı sıfırlıyoruz.

        private SqlConnection sqlConnection;

        public MainForm(bool isAdmin)
        {
            InitializeComponent();
            this.isAdmin = isAdmin;

            // MSSQL veritabanı bağlantısı
            string connectionString = "Server=DESKTOP-3O5O575;Database=Lab_deneme;Integrated Security=True;";
            // Bağlantıyı açıyoruz
            sqlConnection = new SqlConnection(connectionString);
            try
            {
                sqlConnection.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veritabanına bağlanılamadı: " + ex.Message);
            }

            InitializeControls();

            // Yetkiye göre sekmeleri ayarla
            if (!isAdmin)
            {
                tabControl1.TabPages.Remove(tabNewUser);  // Yönetici değilse "Yeni Kullanıcı Ekle" sekmesi kaldırılır
            }
        }

        // Form kapanırken bağlantıyı kapatma işlemi
        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State == System.Data.ConnectionState.Open)
            {
                sqlConnection.Close();
            }
        }

        // Yeni kullanıcı ekleme butonuna tıklandığında çalışacak
        private void btnAddNewUser_Click(object sender, EventArgs e)
        {
            string username = txtNewUsername.Text;
            string password = txtNewPassword.Text;
            string role = cmbUserRole.SelectedItem.ToString();

            // Şifreyi hashle
            string hashedPassword = HashPassword(password);

            // Veritabanına yeni kullanıcıyı ekle
            AddNewUser(username, hashedPassword, role);
        }

        // SHA-256 ile şifreyi hashleyen metod
        private string HashPassword(string password)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password));
                StringBuilder builder = new StringBuilder();
                for (int i = 0; i < bytes.Length; i++)
                {
                    builder.Append(bytes[i].ToString("x2"));
                }
                return builder.ToString();  // Hashlenmiş şifreyi döndür
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // Rol seçeneklerini elle tanımlıyoruz.
            cmbUserRole.Items.Add("admin");
            cmbUserRole.Items.Add("Laboratuvar");
            cmbUserRole.Items.Add("Yönetici");

            cmbTestTipi.Items.AddRange(new string[] { "POY", "DTY", "ATY", "FDY" });
            if (cmbTestTipi.Items.Count > 0)
            {
                cmbTestTipi.SelectedIndex = 0; // İlk test tipini seçiyoruz
            }
            LoadRumuzlar();
            cmbTestTipi.SelectedIndexChanged += cmbTestTipi_SelectedIndexChanged;

            tabPOY.Enter += new EventHandler(tabPOY_Enter);
            tabDTY.Enter += new EventHandler(tabDTY_Enter);
            tabATY.Enter += new EventHandler(tabATY_Enter);
            tabFDY.Enter += new EventHandler(tabFDY_Enter);

            // Test tiplerini ComboBox'a ekleyelim
            cmbTestTipiFiltre.Items.AddRange(new string[] { "POY", "DTY", "ATY", "FDY" });

            // Test tipi seçildiğinde Lot No'ları yükle
            cmbTestTipiFiltre.SelectedIndexChanged += cmbTestTipiFiltre_SelectedIndexChanged;
            cmbLotNoFiltre.SelectedIndexChanged += cmbLotNoFiltre_SelectedIndexChanged;
            dataGridViewFiltreSonucu.AutoGenerateColumns = true;

            cmbKarsilastirmaUrunTipi.SelectedIndexChanged -= cmbKarsilastirmaUrunTipi_SelectedIndexChanged;
            cmbKarsilastirmaRumuz.SelectedIndexChanged -= cmbKarsilastirmaRumuz_SelectedIndexChanged_LoadLot;

            // Verileri ComboBox'lara yükle
            LoadUrunTipleri();
            LoadRumuzlar();

            // Olayları yeniden etkinleştir
            cmbKarsilastirmaUrunTipi.SelectedIndexChanged += cmbKarsilastirmaUrunTipi_SelectedIndexChanged;
            cmbKarsilastirmaRumuz.SelectedIndexChanged += cmbKarsilastirmaRumuz_SelectedIndexChanged_LoadLot;

            cmbKarsilastirmaRumuz.SelectedIndexChanged += new EventHandler(cmbKarsilastirmaRumuz_SelectedIndexChanged_SomeOther);

            LoadTestNedenleri(cmbTestNedeniPOY);  // POY sekmesindeki Test Nedeni ComboBox
            LoadTestNedenleri(cmbTestNedeniDTY);  // DTY sekmesindeki Test Nedeni ComboBox
            LoadTestNedenleri(cmbTestNedeniATY);  // ATY sekmesindeki Test Nedeni ComboBox
            LoadTestNedenleri(cmbTestNedeniFDY);  // FDY sekmesindeki Test Nedeni ComboBox

            LoadUrunTipiRapor();
            cmbUrunTipiRapor.SelectedIndexChanged += cmbUrunTipiRapor_SelectedIndexChanged;
            cmbRumuzRapor.SelectedIndexChanged += cmbRumuzRapor_SelectedIndexChanged;

            LoadStaticComboBoxes();
        }
        private void InitializeControls()
        {
            // ComboBox başlangıç verileri (Sabit veriler)
            cmbTransparency.Items.AddRange(new string[] { "FULL MAT", "SEMIDULL", "S.BRIGHT" });
            cmbCrossSection.Items.AddRange(new string[] { "ROUND", "TRILOBAL", "HOLLOW", "4 CHANNEL", "6 CHANNEL" });
            cmbTube.Items.AddRange(new string[] { "PAPER", "PLASTIC" });
        }

        private void LoadUrunTipleri()
        {
            // Ürün tiplerini ComboBox'a manuel olarak ekliyoruz.
            cmbKarsilastirmaUrunTipi.Items.Clear();
            cmbKarsilastirmaUrunTipi.Items.Add("POY");
            cmbKarsilastirmaUrunTipi.Items.Add("DTY");
            cmbKarsilastirmaUrunTipi.Items.Add("ATY");
            cmbKarsilastirmaUrunTipi.Items.Add("FDY");
            cmbKarsilastirmaUrunTipi.SelectedIndex = 0;  // Varsayılan olarak ilk ürün tipini seçiyoruz
        }

        private void cmbKarsilastirmaUrunTipi_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadRumuzlarKarsilastirma();
        }

        private void LoadRumuzlarKarsilastirma()
        {
            if (cmbKarsilastirmaUrunTipi.SelectedItem == null)
                return;

            string urunTipi = cmbKarsilastirmaUrunTipi.SelectedItem.ToString();
            string query = $"SELECT DISTINCT Rumuz FROM {urunTipi}";

            try
            {
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable rumuzData = new DataTable();
                adapter.Fill(rumuzData);

                cmbKarsilastirmaRumuz.DataSource = rumuzData;
                cmbKarsilastirmaRumuz.DisplayMember = "Rumuz";
                cmbKarsilastirmaRumuz.ValueMember = "Rumuz";
                cmbKarsilastirmaRumuz.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Rumuzlar yüklenirken hata oluştu: " + ex.Message);
            }
        }

        private void LoadTestNedenleri(ComboBox comboBox)
        {
            try
            {
                string query = "SELECT * FROM TestNedenleri";
                SqlDataAdapter dataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                comboBox.DataSource = dataTable;
                comboBox.DisplayMember = "testNedeni";
                comboBox.ValueMember = "testNedeniID";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Test nedenleri yüklenirken hata oluştu: " + ex.Message);
            }
        }


        // Veritabanına yeni kullanıcı ekleyen metod
        private void AddNewUser(string username, string hashedPassword, string role)
        {
            string query = "INSERT INTO Users (Username, PasswordHash, Role) VALUES (@Username, @PasswordHash, @Role)";
            SqlCommand cmd = new SqlCommand(query, sqlConnection);
            cmd.Parameters.AddWithValue("@Username", username);
            cmd.Parameters.AddWithValue("@PasswordHash", hashedPassword);  // Hashlenmiş şifreyi kaydet
            cmd.Parameters.AddWithValue("@Role", role);
            try
            {
                cmd.ExecuteNonQuery();
                MessageBox.Show("Yeni kullanıcı başarıyla eklendi!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kullanıcı ekleme sırasında bir hata oluştu: " + ex.Message);
            }
        }

        private void btnKaydetPOY_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtLotNoPOY.Text))
                {
                    MessageBox.Show("Lütfen geçerli bir LotNo giriniz.");
                    return;
                }

                string query = "INSERT INTO POY (Tarih, musteriAdi, Rumuz, Dtex, flamanSayimi, kopma_Uzama, kopma_UzamaCV, Mukavemet, mukavemetCV, kaynama_Cekme_kaynar_Su, yag, kontrolEden, aciklama, LotNo, Img_Sayisi, Cekim_Kuvveti) " +
                               "VALUES (@Tarih, @musteriAdi, @Rumuz, @Dtex, @flamanSayimi, @kopma_Uzama, @kopma_UzamaCV, @Mukavemet, @mukavemetCV, @kaynama_Cekme, @yag, @kontrolEden, @aciklama, @LotNo, @Img_Sayisi, @Cekim_Kuvveti)";

                SqlCommand cmd = new SqlCommand(query, sqlConnection);

                cmd.Parameters.AddWithValue("@Tarih", dtpTarih.Value);
                cmd.Parameters.AddWithValue("@musteriAdi", txtMusteriAdi.Text);
                cmd.Parameters.AddWithValue("@Rumuz", cmbRumuzPOY.Text);
                cmd.Parameters.AddWithValue("@Dtex", ParseIntegerField(txtDtex.Text));
                cmd.Parameters.AddWithValue("@flamanSayimi", ParseFormattedFloat(txtFlamanSayimi.Text, 1));
                cmd.Parameters.AddWithValue("@kopma_Uzama", ParseFormattedFloat(txtKopmaUzama.Text, 1));
                cmd.Parameters.AddWithValue("@kopma_UzamaCV", ParseFormattedFloat(txtKopmaUzamaCV.Text, 1));
                cmd.Parameters.AddWithValue("@Mukavemet", ParseFormattedFloat(txtMukavemet.Text, 1));
                cmd.Parameters.AddWithValue("@mukavemetCV", ParseFormattedFloat(txtMukavemetCV.Text, 1));
                cmd.Parameters.AddWithValue("@kaynama_Cekme", ParseFormattedFloat(txtKaynamaCekmeKaynarSu.Text, 1));
                cmd.Parameters.AddWithValue("@yag", ParseFormattedFloat(txtYag.Text, 1));
                cmd.Parameters.AddWithValue("@kontrolEden", txtKontrolEden.Text);
                cmd.Parameters.AddWithValue("@aciklama", txtAciklama.Text);
                cmd.Parameters.AddWithValue("@LotNo", txtLotNoPOY.Text);
                cmd.Parameters.AddWithValue("@Img_Sayisi", ParseIntegerField(txtImgSayisi.Text));
                cmd.Parameters.AddWithValue("@Cekim_Kuvveti", ParseFormattedFloat(txtCekimKuvveti.Text, 1));

                cmd.ExecuteNonQuery();

                MessageBox.Show("POY verileri başarıyla kaydedildi.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veri kaydedilirken hata oluştu: " + ex.Message);
            }
        }
        private object ParseFormattedFloat(string text, int decimalPlaces = 1)
        {
            if (float.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out float value) ||
                float.TryParse(text, NumberStyles.Float, new CultureInfo("tr-TR"), out value))
            {
                // Ondalık kısmı iki basamakla sınırlı olarak ayarla
                return Math.Round(value, decimalPlaces);
            }
            return DBNull.Value;
        }

        private object ParseIntegerField(string text)
        {
            if (int.TryParse(text, out int value))
            {
                return value;
            }
            return DBNull.Value;
        }


        private void btnKaydetDTY_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtLotNoDTY.Text))
                {
                    MessageBox.Show("Lütfen geçerli bir LotNo giriniz.");
                    return;
                }

                string query = "INSERT INTO DTY (Tarih, musteriAdi, Rumuz, Dtex, kopma_Uzama, kopma_UzamaCV, Mukavemet, mukavemetCV, kaynama_Cekme_kaynar_Su, yag, kivrim_Kisaltmasi, kivrim_Modulu, kivrim_Kaliciligi, Img_sayi, Img_kalicilik, kontrolEden, aciklama, LotNo) " +
                               "VALUES (@Tarih, @musteriAdi, @Rumuz, @Dtex, @kopma_Uzama, @kopma_UzamaCV, @Mukavemet, @mukavemetCV, @kaynama_Cekme, @yag, @kivrim_Kisaltmasi, @kivrim_Modulu, @kivrim_Kaliciligi, @Img_sayi, @Img_kalicilik, @kontrolEden, @aciklama, @LotNo)";

                SqlCommand cmd = new SqlCommand(query, sqlConnection);

                cmd.Parameters.AddWithValue("@Tarih", dtpTarihDTY.Value);
                cmd.Parameters.AddWithValue("@musteriAdi", txtMusteriAdiDTY.Text);
                cmd.Parameters.AddWithValue("@Rumuz", cmbRumuzDTY.Text);
                cmd.Parameters.AddWithValue("@Dtex", ParseIntegerField(txtDtexDTY.Text));
                cmd.Parameters.AddWithValue("@kopma_Uzama", ParseFormattedFloat(txtKopmaUzamaDTY.Text, 1));
                cmd.Parameters.AddWithValue("@kopma_UzamaCV", ParseFormattedFloat(txtKopmaUzamaCVDTY.Text, 1));
                cmd.Parameters.AddWithValue("@Mukavemet", ParseFormattedFloat(txtMukavemetDTY.Text, 1));
                cmd.Parameters.AddWithValue("@mukavemetCV", ParseFormattedFloat(txtMukavemetCVDTY.Text, 1));
                cmd.Parameters.AddWithValue("@kaynama_Cekme", ParseFormattedFloat(txtKaynamaCekmeKaynarSuDTY.Text, 1));
                cmd.Parameters.AddWithValue("@yag", ParseFormattedFloat(txtYagDTY.Text, 1));
                cmd.Parameters.AddWithValue("@kivrim_Kisaltmasi", ParseFormattedFloat(txtKivrimKisaltmasiDTY.Text, 1));
                cmd.Parameters.AddWithValue("@kivrim_Modulu", ParseFormattedFloat(txtKivrimModuluDTY.Text, 1));
                cmd.Parameters.AddWithValue("@kivrim_Kaliciligi", ParseFormattedFloat(txtKivrimKaliciligiDTY.Text, 1));
                cmd.Parameters.AddWithValue("@Img_sayi", ParseIntegerField(txtImgSayiDTY.Text));
                cmd.Parameters.AddWithValue("@Img_kalicilik", ParseFormattedFloat(txtImgKalicilikDTY.Text, 1));
                cmd.Parameters.AddWithValue("@kontrolEden", txtKontrolEdenDTY.Text);
                cmd.Parameters.AddWithValue("@aciklama", txtAciklamaDTY.Text);
                cmd.Parameters.AddWithValue("@LotNo", txtLotNoDTY.Text);

                cmd.ExecuteNonQuery();

                MessageBox.Show("DTY verileri başarıyla kaydedildi.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veri kaydedilirken hata oluştu: " + ex.Message);
            }
        }
        private void btnKaydetATY_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtLotNoATY.Text))
                {
                    MessageBox.Show("Lütfen geçerli bir LotNo giriniz.");
                    return;
                }

                string query = "INSERT INTO ATY (Tarih, musteriAdi, Rumuz, Dtex, kopma_Uzama, kopma_UzamaCV, Mukavemet, mukavemetCV, kaynama_Cekme_kaynar_Su, yag, kontrolEden, aciklama, LotNo) " +
                               "VALUES (@Tarih, @musteriAdi, @Rumuz, @Dtex, @kopma_Uzama, @kopma_UzamaCV, @Mukavemet, @mukavemetCV, @kaynama_Cekme, @yag, @kontrolEden, @aciklama, @LotNo)";

                SqlCommand cmd = new SqlCommand(query, sqlConnection);

                cmd.Parameters.AddWithValue("@Tarih", dtpTarihATY.Value);
                cmd.Parameters.AddWithValue("@musteriAdi", txtMusteriAdiATY.Text);
                cmd.Parameters.AddWithValue("@Rumuz", cmbRumuzATY.Text);
                cmd.Parameters.AddWithValue("@Dtex", ParseIntegerField(txtDtexATY.Text));
                cmd.Parameters.AddWithValue("@kopma_Uzama", ParseFormattedFloat(txtKopmaUzamaATY.Text, 1));
                cmd.Parameters.AddWithValue("@kopma_UzamaCV", ParseFormattedFloat(txtKopmaUzamaCVATY.Text, 1));
                cmd.Parameters.AddWithValue("@Mukavemet", ParseFormattedFloat(txtMukavemetATY.Text, 1));
                cmd.Parameters.AddWithValue("@mukavemetCV", ParseFormattedFloat(txtMukavemetCVATY.Text, 1));
                cmd.Parameters.AddWithValue("@kaynama_Cekme", ParseFormattedFloat(txtKaynamaCekmeKaynarSuATY.Text, 1));
                cmd.Parameters.AddWithValue("@yag", ParseFormattedFloat(txtYagATY.Text, 1));
                cmd.Parameters.AddWithValue("@kontrolEden", txtKontrolEdenATY.Text);
                cmd.Parameters.AddWithValue("@aciklama", txtAciklamaATY.Text);
                cmd.Parameters.AddWithValue("@LotNo", txtLotNoATY.Text);

                cmd.ExecuteNonQuery();

                MessageBox.Show("ATY verileri başarıyla kaydedildi.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veri kaydedilirken hata oluştu: " + ex.Message);
            }
        }
        private void btnLoadPOYData_Click(object sender, EventArgs e)
        {
            try
            {
                string query = "SELECT * FROM POY";
                SqlDataAdapter dataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                dataGridViewVeriler.DataSource = dataTable;
                dataGridViewVeriler.Columns["Tarih"].DisplayIndex = 0;
                dataGridViewVeriler.Columns["LotNo"].DisplayIndex = 1;
                dataGridViewVeriler.Columns["Rumuz"].DisplayIndex = 2;
            }
            catch (Exception ex)
            {
                MessageBox.Show("POY verileri yüklenirken hata oluştu: " + ex.Message);
            }
        }

        private void btnLoadDTYData_Click(object sender, EventArgs e)
        {
            try
            {
                // DTY verilerini almak için SQL
                // 
                string query = "SELECT * FROM DTY";
                SqlDataAdapter dataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                // DataGridView'e veriyi yükle
                dataGridViewVeriler.DataSource = dataTable;
                dataGridViewVeriler.Columns["Tarih"].DisplayIndex = 0;  // İlk sütun
                dataGridViewVeriler.Columns["LotNo"].DisplayIndex = 1;  // İkinci sütun
                dataGridViewVeriler.Columns["Rumuz"].DisplayIndex = 2;  // Üçüncü sütun (Rumuz)
            }
            catch (Exception ex)
            {
                MessageBox.Show("DTY verileri yüklenirken hata oluştu: " + ex.Message);
            }
        }

        private void btnLoadATYData_Click(object sender, EventArgs e)
        {
            try
            {
                string query = "SELECT * FROM ATY";
                SqlDataAdapter dataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                dataGridViewVeriler.DataSource = dataTable;
                dataGridViewVeriler.Columns["Tarih"].DisplayIndex = 0;
                dataGridViewVeriler.Columns["LotNo"].DisplayIndex = 1;
                dataGridViewVeriler.Columns["Rumuz"].DisplayIndex = 2;
            }
            catch (Exception ex)
            {
                MessageBox.Show("ATY verileri yüklenirken hata oluştu: " + ex.Message);
            }
        }

        private void btnKaydetFDY_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtLotNoFDY.Text))
                {
                    MessageBox.Show("Lütfen geçerli bir LotNo giriniz.");
                    return;
                }

                string query = "INSERT INTO FDY (Tarih, musteriAdi, Rumuz, Dtex, kopma_Uzama, kopma_UzamaCV, Mukavemet, mukavemetCV, kaynama_Cekme_kaynar_Su, yag, kontrolEden, aciklama, testNedeniID, LotNo) " +
                               "VALUES (@Tarih, @musteriAdi, @Rumuz, @Dtex, @kopma_Uzama, @kopma_UzamaCV, @Mukavemet, @mukavemetCV, @kaynama_Cekme, @yag, @kontrolEden, @aciklama, @testNedeniID, @LotNo)";

                SqlCommand cmd = new SqlCommand(query, sqlConnection);

                cmd.Parameters.AddWithValue("@Tarih", dtpTarihFDY.Value);
                cmd.Parameters.AddWithValue("@musteriAdi", txtMusteriAdiFDY.Text);
                cmd.Parameters.AddWithValue("@Rumuz", cmbRumuzFDY.Text);
                cmd.Parameters.AddWithValue("@Dtex", ParseIntegerField(txtDtexFDY.Text));
                cmd.Parameters.AddWithValue("@kopma_Uzama", ParseFormattedFloat(txtKopmaUzamaFDY.Text, 1));
                cmd.Parameters.AddWithValue("@kopma_UzamaCV", ParseFormattedFloat(txtKopmaUzamaCVFDY.Text, 1));
                cmd.Parameters.AddWithValue("@Mukavemet", ParseFormattedFloat(txtMukavemetFDY.Text, 1));
                cmd.Parameters.AddWithValue("@mukavemetCV", ParseFormattedFloat(txtMukavemetCVFDY.Text, 1));
                cmd.Parameters.AddWithValue("@kaynama_Cekme", ParseFormattedFloat(txtKaynamaCekmeKaynarSuFDY.Text, 1));
                cmd.Parameters.AddWithValue("@yag", ParseFormattedFloat(txtYagFDY.Text, 1));
                cmd.Parameters.AddWithValue("@kontrolEden", txtKontrolEdenFDY.Text);
                cmd.Parameters.AddWithValue("@aciklama", txtAciklamaFDY.Text);
                cmd.Parameters.AddWithValue("@testNedeniID", cmbTestNedeniFDY.SelectedValue);
                cmd.Parameters.AddWithValue("@LotNo", txtLotNoFDY.Text);

                cmd.ExecuteNonQuery();

                MessageBox.Show("FDY verileri başarıyla kaydedildi.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veri kaydedilirken hata oluştu: " + ex.Message);
            }
        }
        private void btnLoadFDYData_Click(object sender, EventArgs e)
        {
            try
            {
                string query = "SELECT * FROM FDY";
                SqlDataAdapter dataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                dataGridViewVeriler.DataSource = dataTable;
                dataGridViewVeriler.Columns["Tarih"].DisplayIndex = 0;
                dataGridViewVeriler.Columns["LotNo"].DisplayIndex = 1;
                dataGridViewVeriler.Columns["Rumuz"].DisplayIndex = 2;
            }
            catch (Exception ex)
            {
                MessageBox.Show("FDY verileri yüklenirken hata oluştu: " + ex.Message);
            }
        }

        private void LoadRumuzlar()
        {
            try
            {
                string query = "SELECT rumuzID, rumuz FROM Rumuzlar";
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable rumuzData = new DataTable();
                adapter.Fill(rumuzData);

                cmbRumuz.DataSource = rumuzData;
                cmbRumuz.DisplayMember = "rumuz"; // Display name
                cmbRumuz.ValueMember = "rumuzID"; // Underlying ID
            }
            catch (Exception ex)
            {
                MessageBox.Show("Rumuzlar yüklenirken hata oluştu: " + ex.Message);
            }
        }



        private void btnKaydetLimit_Click(object sender, EventArgs e)
        {
            // ComboBox'ların seçili olup olmadığını kontrol ediyoruz
            if (cmbRumuz.SelectedIndex == -1 || cmbTestTipi.SelectedIndex == -1 || cmbParametre.SelectedIndex == -1)
            {
                MessageBox.Show("Lütfen tüm seçimleri yapınız.");
                return;
            }

            // Seçilen rumuzID'yi kontrol edin ve doğru şekilde alıyoruz
            if (cmbRumuz.SelectedValue == null || !int.TryParse(cmbRumuz.SelectedValue.ToString(), out int rumuzID))
            {
                MessageBox.Show("Geçerli bir rumuz seçiniz.");
                return;
            }

            // Seçilen parametreID'yi kontrol edin ve doğru şekilde alıyoruz
            if (cmbParametre.SelectedValue == null || !int.TryParse(cmbParametre.SelectedValue.ToString(), out int parametreID))
            {
                MessageBox.Show("Geçerli bir parametre giriniz.");
                return;
            }

            // Seçilen testTipi'yi kontrol edin ve doğru şekilde alıyoruz
            if (cmbTestTipi.SelectedItem == null || string.IsNullOrWhiteSpace(cmbTestTipi.SelectedItem.ToString()))
            {
                MessageBox.Show("Geçerli bir test tipi seçiniz.");
                return;
            }

            string testTipi = cmbTestTipi.SelectedItem.ToString();

            // Alt ve üst limitlerin geçerli olup olmadığını kontrol et
            if (!float.TryParse(txtAltLimit.Text, out float altLimit) || !float.TryParse(txtUstLimit.Text, out float ustLimit))
            {
                MessageBox.Show("Alt ve üst limitler geçerli sayılar olmalıdır.");
                return;
            }

            try
            {
                // Veritabanı bağlantısının açık olduğundan emin olun
                if (sqlConnection.State != ConnectionState.Open)
                {
                    sqlConnection.Open();
                }

                // Veritabanına limit ekleme sorgusu
                string query = "INSERT INTO Limitler (rumuzID, testTipi, parametreID, altLimit, ustLimit) " +
                               "VALUES (@rumuzID, @testTipi, @parametreID, @altLimit, @ustLimit)";
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@rumuzID", rumuzID);
                cmd.Parameters.AddWithValue("@testTipi", testTipi);
                cmd.Parameters.AddWithValue("@parametreID", parametreID);
                cmd.Parameters.AddWithValue("@altLimit", altLimit);
                cmd.Parameters.AddWithValue("@ustLimit", ustLimit);

                // Veriyi veritabanına kaydet
                cmd.ExecuteNonQuery();
                MessageBox.Show("Limit başarıyla kaydedildi!");

                // Giriş alanlarını temizle
                cmbRumuz.SelectedIndex = -1;
                cmbTestTipi.SelectedIndex = -1;
                cmbParametre.SelectedIndex = -1;
                txtAltLimit.Clear();
                txtUstLimit.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Limit kaydedilirken hata oluştu: " + ex.Message);
            }
        }

        private void LoadRumuzlarForPOY()
        {
            try
            {
                string query = "SELECT rumuzID, rumuz FROM Rumuzlar";
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable rumuzData = new DataTable();
                adapter.Fill(rumuzData);

                cmbRumuzPOY.DataSource = rumuzData;
                cmbRumuzPOY.DisplayMember = "rumuz";
                cmbRumuzPOY.ValueMember = "rumuzID";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Rumuzlar yüklenirken bir hata oluştu: " + ex.Message);
            }
        }
        private void tabPOY_Enter(object sender, EventArgs e)
        {
            LoadRumuzlarForPOY();
        }

        private void tabPage2_Enter(object sender, EventArgs e)
        {
            LoadRumuzlarForPOY();
        }
        // DTY sekmesi için Rumuzları yükleme fonksiyonu
        private void LoadRumuzlarForDTY()
        {
            try
            {
                string query = "SELECT rumuzID, rumuz FROM Rumuzlar";  // Veritabanından rumuzları çekiyoruz
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable rumuzData = new DataTable();
                adapter.Fill(rumuzData);

                // DTY sekmesindeki ComboBox'ı doldur
                cmbRumuzDTY.DataSource = rumuzData;
                cmbRumuzDTY.DisplayMember = "rumuz";  // Görünen kısım
                cmbRumuzDTY.ValueMember = "rumuzID";  // Değer olarak rumuzID'yi kullanıyoruz
            }
            catch (Exception ex)
            {
                MessageBox.Show("Rumuzlar yüklenirken bir hata oluştu: " + ex.Message);
            }
        }

        private void LoadRumuzlarForATY()
        {
            try
            {
                string query = "SELECT rumuzID, rumuz FROM Rumuzlar";  // Veritabanından rumuzları çekiyoruz
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable rumuzData = new DataTable();
                adapter.Fill(rumuzData);

                // ATY sekmesindeki ComboBox'ı doldur
                cmbRumuzATY.DataSource = rumuzData;
                cmbRumuzATY.DisplayMember = "rumuz";  // Görünen kısım
                cmbRumuzATY.ValueMember = "rumuzID";  // Değer olarak rumuzID'yi kullanıyoruz
            }
            catch (Exception ex)
            {
                MessageBox.Show("Rumuzlar yüklenirken bir hata oluştu: " + ex.Message);
            }
        }
        private void LoadRumuzlarForFDY()
        {
            try
            {
                string query = "SELECT rumuzID, rumuz FROM Rumuzlar";  // Veritabanından rumuzları çekiyoruz
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable rumuzData = new DataTable();
                adapter.Fill(rumuzData);

                // FDY sekmesindeki ComboBox'ı doldur
                cmbRumuzFDY.DataSource = rumuzData;
                cmbRumuzFDY.DisplayMember = "rumuz";  // Görünen kısım
                cmbRumuzFDY.ValueMember = "rumuzID";  // Değer olarak rumuzID'yi kullanıyoruz
            }
            catch (Exception ex)
            {
                MessageBox.Show("Rumuzlar yüklenirken bir hata oluştu: " + ex.Message);
            }
        }

        private void tabFDY_Enter(object sender, EventArgs e)
        {
            LoadRumuzlarForFDY();  // FDY sekmesi açıldığında rumuzlar yüklensin
        }



        private void tabATY_Enter(object sender, EventArgs e)
        {
            LoadRumuzlarForATY();  // ATY sekmesi açıldığında rumuzlar yüklensin
        }


        private void tabDTY_Enter(object sender, EventArgs e)
        {
            LoadRumuzlarForDTY();  // DTY sekmesi açıldığında rumuzlar yüklensin
        }



        private void cmbTestTipi_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbTestTipi.SelectedItem == null)
            {
                MessageBox.Show("Lütfen bir Test Tipi seçiniz.");
                return;
            }

            string selectedTest = cmbTestTipi.SelectedItem.ToString();

            try
            {
                string query = "SELECT ParametreID, Parametre FROM Parametreler WHERE TestTipi = @TestTipi";
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@TestTipi", selectedTest);

                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable parametreData = new DataTable();
                adapter.Fill(parametreData);

                cmbParametre.DataSource = parametreData;
                cmbParametre.DisplayMember = "Parametre";
                cmbParametre.ValueMember = "ParametreID";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Parametreler yüklenirken hata oluştu: " + ex.Message);
            }
        }
        private void cmbTestTipiFiltre_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbTestTipiFiltre.SelectedItem == null)
            {
                MessageBox.Show("Lütfen bir Test Tipi seçiniz.");
                return;
            }

            // Seçilen test tipine göre Lot No'ları getiriyoruz
            string selectedTestTipi = cmbTestTipiFiltre.SelectedItem.ToString();
            string query = $"SELECT DISTINCT LotNo FROM {selectedTestTipi}";

            try
            {
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable lotData = new DataTable();
                adapter.Fill(lotData);

                cmbLotNoFiltre.DataSource = lotData;
                cmbLotNoFiltre.DisplayMember = "LotNo";
                cmbLotNoFiltre.ValueMember = "LotNo";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lot No'lar yüklenirken hata oluştu: " + ex.Message);
            }
        }
        // Lot No seçildiğinde ilgili Rumuz'ları yükleyen metod
        private void cmbLotNoFiltre_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbTestTipiFiltre.SelectedItem == null || cmbLotNoFiltre.SelectedItem == null)
            {
                MessageBox.Show("Lütfen önce test tipi ve lot numarası seçiniz.");
                return;
            }

            // Seçilen Test Tipi ve Lot No
            string selectedTestTipi = cmbTestTipiFiltre.SelectedItem.ToString();
            string selectedLotNo = cmbLotNoFiltre.SelectedValue?.ToString() ?? cmbLotNoFiltre.Text;

            string query = $"SELECT DISTINCT Rumuz FROM {selectedTestTipi} WHERE LotNo = @LotNo";

            try
            {
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@LotNo", selectedLotNo);

                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable rumuzData = new DataTable();
                adapter.Fill(rumuzData);

                cmbRumuzFiltre.DataSource = rumuzData;
                cmbRumuzFiltre.DisplayMember = "Rumuz";
                cmbRumuzFiltre.ValueMember = "Rumuz";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Rumuzlar yüklenirken bir hata oluştu: " + ex.Message);
            }
        }
        private void btnFiltrele_Click(object sender, EventArgs e)
        {
            if (sqlConnection.State != ConnectionState.Open)
            {
                sqlConnection.Open();
            }

            // Test Tipi, Lot No ve Rumuz seçimlerini kontrol ediyoruz
            if (cmbTestTipiFiltre.SelectedItem == null || cmbLotNoFiltre.SelectedItem == null || cmbRumuzFiltre.SelectedItem == null)
            {
                MessageBox.Show("Lütfen tüm seçimleri yapınız.");
                return;
            }

            // Seçilen değerleri alıyoruz
            string selectedTestTipi = cmbTestTipiFiltre.SelectedItem.ToString();
            string selectedLotNo = cmbLotNoFiltre.Text;  // SelectedItem yerine Text kullanıyoruz
            string selectedRumuz = cmbRumuzFiltre.Text;

            try
            {
                // SQL sorgusunu oluşturuyoruz
                string query = $"SELECT * FROM {selectedTestTipi} WHERE LTRIM(RTRIM(UPPER(LotNo))) = LTRIM(RTRIM(UPPER(@LotNo))) AND LTRIM(RTRIM(UPPER(Rumuz))) = LTRIM(RTRIM(UPPER(@Rumuz)))";
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@LotNo", selectedLotNo);
                cmd.Parameters.AddWithValue("@Rumuz", selectedRumuz);

                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable filteredData = new DataTable();
                adapter.Fill(filteredData);

                // Eğer sonuç varsa DataGridView'e yükleyelim
                if (filteredData.Rows.Count > 0)
                {
                    dataGridViewFiltreSonucu.DataSource = filteredData;
                }
                else
                {
                    MessageBox.Show("Kayıt bulunamadı.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Filtreleme sırasında bir hata oluştu: " + ex.Message);
            }
        }
        private void dataGridViewFiltreSonucu_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnKarsilastir_Click(object sender, EventArgs e)
        {
            // Ürün tipi, rumuz ve lotların seçilmiş olup olmadığını kontrol edelim
            if (cmbKarsilastirmaUrunTipi.SelectedItem == null ||
                cmbKarsilastirmaRumuz.SelectedItem == null ||
                cmbKarsilastirmaLot1.SelectedItem == null ||
                cmbKarsilastirmaLot2.SelectedItem == null)
            {
                MessageBox.Show("Lütfen geçerli bir ürün tipi, rumuz ve lot numarası seçiniz.");
                return;
            }

            // Seçilen değerler
            string urunTipi = cmbKarsilastirmaUrunTipi.SelectedItem.ToString();
            string rumuz = cmbKarsilastirmaRumuz.Text;
            string lot1 = cmbKarsilastirmaLot1.Text;
            string lot2 = cmbKarsilastirmaLot2.Text;

            CompareLotsAltAlta(urunTipi, lot1, lot2, rumuz);
        }

        private void CompareLotsAltAlta(string urunTipi, string lot1, string lot2, string rumuz)
        {
            try
            {
                // Lotlar için SQL sorgusu
                string query = $"SELECT * FROM {urunTipi} WHERE (LotNo = @Lot1 OR LotNo = @Lot2) AND Rumuz = @Rumuz";
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@Lot1", lot1);
                cmd.Parameters.AddWithValue("@Lot2", lot2);
                cmd.Parameters.AddWithValue("@Rumuz", rumuz);

                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable karsilastirmaData = new DataTable();
                adapter.Fill(karsilastirmaData);

                if (karsilastirmaData.Rows.Count == 2)
                {
                    // DataGridView'e karşılaştırma sonuçlarını yükle
                    DataTable resultTable = new DataTable();
                    resultTable.Columns.Add("Parametre");
                    resultTable.Columns.Add($"Lot {lot1}");
                    resultTable.Columns.Add($"Lot {lot2}");

                    DataRow lot1Row = karsilastirmaData.Select($"LotNo = '{lot1}'").FirstOrDefault();
                    DataRow lot2Row = karsilastirmaData.Select($"LotNo = '{lot2}'").FirstOrDefault();

                    foreach (DataColumn column in karsilastirmaData.Columns)
                    {
                        if (column.ColumnName != "LotNo" && column.ColumnName != "Rumuz")
                        {
                            // Parametre adı
                            string parametre = column.ColumnName;

                            // Lot 1 ve Lot 2 değerleri
                            string lot1Value = lot1Row?[column.ColumnName]?.ToString() ?? "N/A";
                            string lot2Value = lot2Row?[column.ColumnName]?.ToString() ?? "N/A";

                            // Sonuç tablosuna ekle
                            DataRow resultRow = resultTable.NewRow();
                            resultRow["Parametre"] = parametre;
                            resultRow[$"Lot {lot1}"] = lot1Value;
                            resultRow[$"Lot {lot2}"] = lot2Value;
                            resultTable.Rows.Add(resultRow);
                        }
                    }

                    // DataGridView'e sonucu ekle
                    dataGridViewKarsilastirmaSonucu.DataSource = resultTable;
                }
                else
                {
                    MessageBox.Show("Seçilen lotlar arasında karşılaştırılacak veri bulunamadı.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Karşılaştırma sırasında bir hata oluştu: " + ex.Message);
            }
        }
        private void cmbKarsilastirmaRumuz_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadLotNolariKarsilastirma();
        }


        private void LoadLotNolariKarsilastirma()
        {
            if (cmbKarsilastirmaRumuz.SelectedItem == null || cmbKarsilastirmaUrunTipi.SelectedItem == null)
            {
                if (!isLotWarningShown)
                {
                    MessageBox.Show("Lütfen geçerli bir ürün tipi veya rumuz seçiniz.");
                    isLotWarningShown = true;
                }
                return;
            }

            string urunTipi = cmbKarsilastirmaUrunTipi.SelectedItem.ToString();
            string rumuz = cmbKarsilastirmaRumuz.Text;

            string query = $"SELECT DISTINCT LotNo FROM {urunTipi} WHERE Rumuz = @Rumuz";

            try
            {
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@Rumuz", rumuz);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable lotData = new DataTable();
                adapter.Fill(lotData);

                if (lotData.Rows.Count == 0)
                {
                    MessageBox.Show($"Seçilen {urunTipi} ve rumuz '{rumuz}' için lot numarası bulunamadı.");
                    return;
                }

                cmbKarsilastirmaLot1.DataSource = lotData.Copy();
                cmbKarsilastirmaLot1.DisplayMember = "LotNo";
                cmbKarsilastirmaLot1.ValueMember = "LotNo";
                cmbKarsilastirmaLot1.SelectedIndex = -1;

                cmbKarsilastirmaLot2.DataSource = lotData.Copy();
                cmbKarsilastirmaLot2.DisplayMember = "LotNo";
                cmbKarsilastirmaLot2.ValueMember = "LotNo";
                cmbKarsilastirmaLot2.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lot numaraları yüklenirken hata oluştu: " + ex.Message);
            }
        }
        private void cmbKarsilastirmaRumuz_SelectedIndexChanged_LoadLot(object sender, EventArgs e)
        {
            LoadLotNolariKarsilastirma();
        }

        //private void cmbKarsilastirmaUrunTipi_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    LoadRumuzlarKarsilastirma();
        //}

        // Rumuz seçildiğinde lot numaraları yüklenecek
        private void cmbKarsilastirmaRumuz_SelectedIndexChanged_SomeOther(object sender, EventArgs e)
        {
            LoadLotNolariKarsilastirma();
        }

        private void btnTestNedeniEkle_Click1(object sender, EventArgs e)
        {
            if (isAdmin || isYonetici)
            {
                string yeniTestNedeni = txtYeniTestNedeni.Text;
                if (!string.IsNullOrWhiteSpace(yeniTestNedeni))
                {
                    try
                    {
                        string query = "INSERT INTO TestNedenleri (testNedeni) VALUES (@testNedeni)";
                        SqlCommand cmd = new SqlCommand(query, sqlConnection);
                        cmd.Parameters.AddWithValue("@testNedeni", yeniTestNedeni);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Test nedeni başarıyla eklendi!");

                        // Tüm Test Nedeni ComboBox'larını güncelleyin
                        LoadTestNedenleri(cmbTestNedeniPOY);
                        LoadTestNedenleri(cmbTestNedeniDTY);
                        LoadTestNedenleri(cmbTestNedeniATY);
                        LoadTestNedenleri(cmbTestNedeniFDY);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Test nedeni eklenirken hata oluştu: " + ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("Test nedeni boş olamaz.");
                }
            }
            else
            {
                MessageBox.Show("Bu işlemi sadece admin veya yönetici yapabilir.");
            }
        }

        private void LoadTestSonuclari(string urunTipi, string rumuz, string lotNo)
        {
            string query = $"SELECT * FROM {urunTipi} WHERE Rumuz = @Rumuz AND LotNo = @LotNo";
            SqlCommand cmd = new SqlCommand(query, sqlConnection);
            cmd.Parameters.AddWithValue("@Rumuz", rumuz);
            cmd.Parameters.AddWithValue("@LotNo", lotNo);

            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataTable testSonuclariTable = new DataTable();
            adapter.Fill(testSonuclariTable);

            if (testSonuclariTable.Rows.Count == 0)
            {
                MessageBox.Show("Seçilen kriterlere uygun test sonucu bulunamadı.");
                return;
            }

            
        }


        private void btnRaporOlustur_Click(object sender, EventArgs e)
        {
            if (cmbUrunTipiRapor.SelectedIndex == -1 ||
                cmbRumuzRapor.SelectedIndex == -1 ||
                cmbLotNoRapor.SelectedIndex == -1)
            {
                MessageBox.Show("Lütfen gerekli seçimleri yapınız!");
                return;
            }

            // Kullanıcıdan alınan veriler
            var userData = new
            {
                Transparency = cmbTransparency.SelectedItem?.ToString(),
                CrossSection = cmbCrossSection.SelectedItem?.ToString(),
                Color = txtColor.Text,
                Tube = cmbTube.SelectedItem?.ToString()
            };

            // Debug: Parametre kontrolü
            MessageBox.Show($@"
Sorgu: SELECT * FROM {cmbUrunTipiRapor.SelectedItem} WHERE Rumuz = @Rumuz AND LotNo = @LotNo
Parametreler -> 
    Rumuz: {cmbRumuzRapor.SelectedValue?.ToString() ?? "Null"}, 
    LotNo: {cmbLotNoRapor.SelectedValue?.ToString() ?? "Null"}");

            // Test sonuçlarını al
            DataTable testResults = GetTestResultsForReport(
                cmbUrunTipiRapor.SelectedItem?.ToString(),
                cmbRumuzRapor.SelectedValue?.ToString(),
                cmbLotNoRapor.SelectedValue?.ToString()
            );

            // Eğer test sonucu yoksa kullanıcıya hata göster
            if (testResults == null || testResults.Rows.Count == 0)
            {
                MessageBox.Show("Sorgu başarılı çalıştı fakat sonuç döndürmedi.");
                return;
            }

            // Excel raporu oluştur
            CreateExcelReport(userData, testResults);
        }

        private DataTable GetTestSonucuByID(int testID, string urunTipi)
        {
            string query = $"SELECT * FROM {urunTipi} WHERE TestID = @TestID"; // Örnek sorgu, kendi veritabanınıza göre düzenleyin
            SqlCommand cmd = new SqlCommand(query, sqlConnection);
            cmd.Parameters.AddWithValue("@TestID", testID);

            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);
            return dataTable;
        }


        private void CreateExcelReport(dynamic userData, DataTable testResults)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // ComboBox'lardan seçilen değerleri doğru bir şekilde alıyoruz
            string rumuz = cmbRumuzRapor.SelectedValue?.ToString() ?? cmbRumuzRapor.Text;
            string lotNo = cmbLotNoRapor.SelectedValue?.ToString() ?? cmbLotNoRapor.Text;
            string tarih = DateTime.Now.ToString("yyyyMMdd");

            // Yeni dosya adı formatı: Rumuz_LotNo_Tarih.xlsx
            string newFileName = $"{rumuz}_{lotNo}_{tarih}.xlsx";
            string newFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), newFileName);

            string templatePath = "C:\\Users\\Laboratuvar\\Desktop\\Test Raporu Boş - deneme.xlsx";

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(templatePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Kullanıcı girdilerini bir eşleme (mapping) ile tanımlıyoruz
                    var userInputs = new[]
                    {
                new { Cell = "F13", Value = cmbUrunTipiRapor.Text },    // Product
                new { Cell = "F14", Value = rumuz },                   // Rumuz
                new { Cell = "F15", Value = lotNo },                   // Lot Number
                new { Cell = "F16", Value = cmbTransparency.SelectedItem?.ToString() }, // Transparency
                new { Cell = "F17", Value = cmbCrossSection.SelectedItem?.ToString() }, // Cross-Section
                new { Cell = "F18", Value = txtColor.Text },           // Color
                new { Cell = "F19", Value = cmbTube.SelectedItem?.ToString() } // Tube
            };

                    // Kullanıcı girdilerini yazdır
                    foreach (var input in userInputs)
                    {
                        worksheet.Cells[input.Cell].Value = string.IsNullOrWhiteSpace(input.Value) ? "  " : input.Value;
                    }

                    // Test Results kısmını doldurma
                    int startRow = 22; // Test Results'un başlangıç satırı
                    var propertyMappings = new (string ColumnName, string TestResultCell)[]
                    {
                ("Dtex", $"H{startRow}"),  // Linear Density
                ("Mukavemet", $"H{startRow + 1}"),  // Tenacity
                ("kopma_Uzama", $"H{startRow + 2}"),  // Elongation at Break
                ("kaynama_Cekme_kaynar_Su", $"H{startRow + 3}"),  // Shrinkage at Boiling
                ("kivrim_Kisaltmasi", $"H{startRow + 4}"),  // Crimp Contraction
                ("Img_sayi", $"H{startRow + 5}"),  // Intermingling
                ("yag", $"H{startRow + 6}")  // Amount of Finish
                    };

                    foreach (var (columnName, cellAddress) in propertyMappings)
                    {
                        if (testResults.Columns.Contains(columnName) && testResults.Rows.Count > 0)
                        {
                            worksheet.Cells[cellAddress].Value = testResults.Rows[0][columnName]?.ToString() ?? "  ";
                        }
                        else
                        {
                            worksheet.Cells[cellAddress].Value = "  "; // Veri yoksa N/A yazdır
                        }
                    }

                    // Raporu kaydet
                    package.SaveAs(new FileInfo(newFilePath));
                    MessageBox.Show($"Rapor başarıyla oluşturuldu: {newFilePath}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Rapor oluşturulurken bir hata oluştu: " + ex.Message);
            }
        }

        private DataTable GetTestResults(string urunTipi, string rumuz, string lotNo)
        {
            // SQL sorgusunu tanımlıyoruz
            string query = $@"
        SELECT * 
        FROM {urunTipi} 
        WHERE LTRIM(RTRIM(UPPER(Rumuz))) = LTRIM(RTRIM(UPPER(@Rumuz)))
        AND LTRIM(RTRIM(UPPER(LotNo))) = LTRIM(RTRIM(UPPER(@LotNo)))";

            try
            {
                // SQL komutunu oluşturuyoruz
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@Rumuz", rumuz.Trim());
                cmd.Parameters.AddWithValue("@LotNo", lotNo.Trim());

                // Verileri doldurmak için SqlDataAdapter kullanıyoruz
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                // Debug veya sonuç kontrolü için şu kısmı ekleyebilirsiniz
                if (dataTable.Rows.Count == 0)
                {
                    // Eğer sonuç döndürmediyse, kullanıcıya bilgi veriyoruz
                    MessageBox.Show("Sorgu başarılı çalıştı ancak sonuç döndürmedi.\n" +
                                    $"Ürün Tipi: {urunTipi}, Rumuz: {rumuz}, Lot No: {lotNo}");
                }
                else
                {
                    // Eğer sonuç döndürdüyse, ilk sonuç hakkında bilgi gösterelim
                    MessageBox.Show($"Sonuç bulundu! İlk kayıt:\nRumuz: {dataTable.Rows[0]["Rumuz"]}\n" +
                                    $"LotNo: {dataTable.Rows[0]["LotNo"]}");
                }

                // DataTable'ı geri döndürüyoruz
                return dataTable;
            }
            catch (Exception ex)
            {
                // Hata durumunda mesaj gösteriyoruz
                MessageBox.Show($"SQL Hatası: {ex.Message}");
                return null;
            }
        }


        private void LoadUrunTipiRapor()
        {
            cmbUrunTipiRapor.Items.Clear();
            cmbUrunTipiRapor.Items.AddRange(new string[] { "POY", "DTY", "ATY", "FDY" });
            cmbUrunTipiRapor.SelectedIndex = 0; // Varsayılan olarak ilk ürünü seçiyoruz
        }

        private void LoadStaticComboBoxes()
        {
            // Sabit ComboBox değerleri manuel olarak eklendi
            cmbTransparency.Items.AddRange(new string[] { "FULL MAT", "SEMIDULL", "S.BRIGHT" });
            cmbCrossSection.Items.AddRange(new string[] { "ROUND", "TRILOBAL", "HOLLOW", "4 CHANNEL", "6 CHANNEL" });
            cmbTube.Items.AddRange(new string[] { "PAPER", "PLASTIC" });
        }

        private void LoadRumuzlarRapor(string urunTipi)
        {
            string query = $"SELECT DISTINCT Rumuz FROM {urunTipi}";
            SqlCommand cmd = new SqlCommand(query, sqlConnection);

            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable rumuzData = new DataTable();
                adapter.Fill(rumuzData);

                MessageBox.Show($"Rumuzlar Yüklendi: {string.Join(", ", rumuzData.Rows.OfType<DataRow>().Select(r => r["Rumuz"].ToString()))}");

                cmbRumuzRapor.DataSource = rumuzData;
                cmbRumuzRapor.DisplayMember = "Rumuz";
                cmbRumuzRapor.ValueMember = "Rumuz";
                cmbRumuzRapor.SelectedIndex = -1; // Başlangıçta seçili değil
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Rumuzlar yüklenirken hata oluştu: {ex.Message}");
            }
        }

        private void LoadLotNoRapor(string urunTipi, string rumuz)
        {
            if (string.IsNullOrEmpty(urunTipi) || string.IsNullOrEmpty(rumuz)) return;

            string query = $"SELECT DISTINCT LotNo FROM {urunTipi} WHERE Rumuz = @Rumuz";
            SqlCommand cmd = new SqlCommand(query, sqlConnection);
            cmd.Parameters.AddWithValue("@Rumuz", rumuz);

            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable lotData = new DataTable();
                adapter.Fill(lotData);

                cmbLotNoRapor.DataSource = lotData;
                cmbLotNoRapor.DisplayMember = "LotNo";
                cmbLotNoRapor.ValueMember = "LotNo";
                cmbLotNoRapor.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lot No'lar yüklenirken hata oluştu: " + ex.Message);
            }
        }
        private void cmbUrunTipiRapor_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedUrunTipi = cmbUrunTipiRapor.SelectedItem?.ToString();
            if (!string.IsNullOrEmpty(selectedUrunTipi))
            {
                LoadRumuzlarRapor(selectedUrunTipi);
            }
        }
        private void cmbRumuzRapor_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedUrunTipi = cmbUrunTipiRapor.SelectedItem?.ToString();
            string selectedRumuz = cmbRumuzRapor.SelectedValue?.ToString();
            if (!string.IsNullOrEmpty(selectedUrunTipi) && !string.IsNullOrEmpty(selectedRumuz))
            {
                LoadLotNoRapor(selectedUrunTipi, selectedRumuz);
            }
        }
        private void LoadComboBoxData()
        {
            // Ürün Tipi, Rumuz ve Lot No'yu veritabanından çek
            LoadFromDatabase("POY", cmbUrunTipiRapor);
            LoadFromDatabase("Rumuzlar", cmbRumuzRapor);
            LoadFromDatabase("LotNo", cmbLotNoRapor);

            // Sabit verileri manuel olarak ekle
            cmbTransparency.Items.AddRange(new string[] { "FULL MAT", "SEMIDULL", "S.BRIGHT" });
            cmbCrossSection.Items.AddRange(new string[] { "ROUND", "TRILOBAL", "HOLLOW", "4 CHANNEL", "6 CHANNEL" });
            cmbTube.Items.AddRange(new string[] { "PAPER", "PLASTIC" }); // Tube için sabit değerler
        }

        private DataTable GetTestResultsForReport(string urunTipi, string rumuz, string lotNo)
        {
            try
            {
                // SQL sorgusu
                string query = $"SELECT * FROM {urunTipi} WHERE Rumuz = @Rumuz AND LotNo = @LotNo";
                SqlCommand cmd = new SqlCommand(query, sqlConnection);

                // Parametreleri ekliyoruz
                cmd.Parameters.AddWithValue("@Rumuz", rumuz.Trim());
                cmd.Parameters.AddWithValue("@LotNo", lotNo.Trim());

                // Debugging: Parametreleri ve sorguyu kontrol edin
                MessageBox.Show($"Sorgu: {query}\nParametreler -> Rumuz: {rumuz}, LotNo: {lotNo}");

                // Sorguyu çalıştır ve sonuçları DataTable'a yükle
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                // Eğer sonuç dönmezse bilgi mesajı göster
                if (dataTable.Rows.Count == 0)
                {
                    MessageBox.Show("Sorgu başarılı çalıştı fakat sonuç döndürmedi.\n" +
                                    $"Ürün Tipi: {urunTipi}, Rumuz: {rumuz}, Lot No: {lotNo}");
                }
                else
                {
                    // İlk kayıt hakkında bilgi ver
                    MessageBox.Show($"Sonuç bulundu! İlk kayıt:\nRumuz: {dataTable.Rows[0]["Rumuz"]}\n" +
                                    $"LotNo: {dataTable.Rows[0]["LotNo"]}");
                }

                return dataTable; // Sonuçları döndür
            }
            catch (Exception ex)
            {
                // Hata durumunda mesaj göster
                MessageBox.Show($"SQL Hatası: {ex.Message}");
                return null; // Hata durumunda null döndür
            }
        }


        private void LoadFromDatabase(string tableName, ComboBox comboBox)
        {
            try
            {
                string query = $"SELECT DISTINCT * FROM {tableName}";
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                comboBox.DataSource = dataTable;
                comboBox.DisplayMember = "Name"; // Örnek kolon adı
                comboBox.ValueMember = "ID";     // Örnek kolon adı
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hata: {ex.Message}");
            }
        }
    }
}
        // POY için LotNo ve Rumuz çekme işlemi

    