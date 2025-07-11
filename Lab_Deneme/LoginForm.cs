using System;
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace Lab_Deneme
{
    public partial class LoginForm : Form
    {
        private SqlConnection sqlConnection;

        public LoginForm()
        {
            InitializeComponent();

            // MSSQL veritabanı bağlantısı
            string connectionString = "Server=DESKTOP-Q50SUEF;Database=Lab_Deneme;Integrated Security=True;";
            sqlConnection = new SqlConnection(connectionString);
            sqlConnection.Open();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            string username = txtUsername.Text;
            string password = txtPassword.Text;

            try
            {
                // Veritabanı bağlantısı
                MessageBox.Show("Veritabanına bağlanılıyor...");

                string query = "SELECT PasswordHash, Role FROM Users WHERE Username = @Username";
                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.Parameters.AddWithValue("@Username", username);

                SqlDataReader reader = cmd.ExecuteReader();

                // Kullanıcı kontrolü
                if (reader.Read())
                {
                    string storedHash = reader["PasswordHash"].ToString();
                    string role = reader["Role"].ToString();

                    // Şifreyi hashleyip karşılaştırma
                    if (VerifyPasswordHash(password, storedHash))
                    {
                        MessageBox.Show("Giriş başarılı!");

                        bool isAdmin = role == "admin"; // Rolü belirle
                        MainForm mainForm = new MainForm(isAdmin);
                        mainForm.Show();
                        this.Hide();  // Giriş formunu gizle
                    }
                    else
                    {
                        MessageBox.Show("Hatalı şifre!");
                    }
                }
                else
                {
                    MessageBox.Show("Kullanıcı bulunamadı!");
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }
        }


        // Şifre hash doğrulama metodu
        private bool VerifyPasswordHash(string password, string storedHash)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password));
                StringBuilder builder = new StringBuilder();
                for (int i = 0; i < bytes.Length; i++)
                {
                    builder.Append(bytes[i].ToString("x2"));
                }
                string hashOfInput = builder.ToString();

                return StringComparer.OrdinalIgnoreCase.Compare(hashOfInput, storedHash) == 0;
            }
        }

        private void btnLogin_Click_1(object sender, EventArgs e)
        {

        }
    }
}
