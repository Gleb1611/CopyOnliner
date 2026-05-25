using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;

namespace CopyOnliner
{
    public partial class FormLogin : Form
    {
        private OleDbConnection connection;
        private bool isRegisterMode = false;
        private User currentUser;

        private Color darkBg = Color.FromArgb(15, 25, 35);
        private Color panelBg = Color.FromArgb(25, 35, 45);
        private Color accentColor = Color.FromArgb(0, 255, 100);
        private Color accentBlue = Color.FromArgb(0, 150, 255);
        private Color textColor = Color.FromArgb(220, 220, 220);

        public User CurrentUser => currentUser;

        public FormLogin()
        {
            InitializeComponent();
            SetupForm();
            CreateUsersTableIfNotExists();

            try
            {
                connection = new OleDbConnection(Form1.connectionString);
                connection.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка подключения: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetupForm()
        {
            this.Text = "Вход / Регистрация";
            this.Size = new Size(480, 600);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = darkBg;
        }

        private void CreateUsersTableIfNotExists()
        {
            try
            {
                using (OleDbConnection tempConn = new OleDbConnection(Form1.connectionString))
                {
                    tempConn.Open();
                    string createTableQuery = @"
                        CREATE TABLE Customers (
                            CustomerID COUNTER PRIMARY KEY,
                            FullName TEXT(100) NOT NULL,
                            Email TEXT(100) NOT NULL,
                            Password TEXT(50) NOT NULL
                        )";
                    try
                    {
                        OleDbCommand cmd = new OleDbCommand(createTableQuery, tempConn);
                        cmd.ExecuteNonQuery();
                    }
                    catch { }
                    tempConn.Close();
                }
            }
            catch { }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            DisplayLoginForm();
        }

        private void DisplayLoginForm()
        {
            this.Controls.Clear();

            Panel mainPanel = new Panel();
            mainPanel.Dock = DockStyle.Fill;
            mainPanel.BackColor = panelBg;
            mainPanel.Padding = new Padding(30);

            int yOffset = 30;

            Label lblTitle = new Label();
            lblTitle.Text = isRegisterMode ? "📝 РЕГИСТРАЦИЯ" : "🔐 ВХОД В СИСТЕМУ";
            lblTitle.Location = new Point(0, yOffset);
            lblTitle.Size = new Size(420, 40);
            lblTitle.Font = new Font("Segoe UI", 16, FontStyle.Bold);
            lblTitle.ForeColor = accentColor;
            lblTitle.TextAlign = ContentAlignment.MiddleCenter;
            mainPanel.Controls.Add(lblTitle);
            yOffset += 60;

            if (isRegisterMode)
            {
                // Полное имя
                AddLabel(mainPanel, "👤 Полное имя:", yOffset);
                TextBox txtFullName = AddTextBox(mainPanel, yOffset + 25);
                yOffset += 65;

                // Email
                AddLabel(mainPanel, "📧 Email:", yOffset);
                TextBox txtEmail = AddTextBox(mainPanel, yOffset + 25);
                yOffset += 65;

                // Пароль
                AddLabel(mainPanel, "🔒 Пароль:", yOffset);
                TextBox txtPassword = AddTextBox(mainPanel, yOffset + 25);
                txtPassword.PasswordChar = '*';
                yOffset += 65;

                // Подтверждение пароля
                AddLabel(mainPanel, "🔒 Подтвердите пароль:", yOffset);
                TextBox txtConfirmPassword = AddTextBox(mainPanel, yOffset + 25);
                txtConfirmPassword.PasswordChar = '*';
                yOffset += 70;

                Button btnRegister = new Button();
                btnRegister.Text = "✅ ЗАРЕГИСТРИРОВАТЬСЯ";
                btnRegister.Location = new Point(50, yOffset);
                btnRegister.Size = new Size(320, 45);
                btnRegister.BackColor = accentColor;
                btnRegister.ForeColor = Color.Black;
                btnRegister.Font = new Font("Segoe UI", 11, FontStyle.Bold);
                btnRegister.FlatStyle = FlatStyle.Flat;
                btnRegister.Cursor = Cursors.Hand;
                btnRegister.Click += (s, e) => RegisterUser(
                    txtFullName.Text, txtEmail.Text, txtPassword.Text, txtConfirmPassword.Text);
                mainPanel.Controls.Add(btnRegister);
                yOffset += 60;

                LinkLabel lnkSwitch = new LinkLabel();
                lnkSwitch.Text = "Уже есть аккаунт? Войти";
                lnkSwitch.Location = new Point(120, yOffset);
                lnkSwitch.Size = new Size(200, 25);
                lnkSwitch.Font = new Font("Segoe UI", 10);
                lnkSwitch.LinkColor = accentBlue;
                lnkSwitch.TextAlign = ContentAlignment.MiddleCenter;
                lnkSwitch.Click += (s, e) => { isRegisterMode = false; DisplayLoginForm(); };
                mainPanel.Controls.Add(lnkSwitch);
            }
            else
            {
                // Email
                AddLabel(mainPanel, "📧 Email:", yOffset);
                TextBox txtEmail = AddTextBox(mainPanel, yOffset + 25);
                yOffset += 65;

                // Пароль
                AddLabel(mainPanel, "🔒 Пароль:", yOffset);
                TextBox txtPassword = AddTextBox(mainPanel, yOffset + 25);
                txtPassword.PasswordChar = '*';
                yOffset += 70;

                Button btnLogin = new Button();
                btnLogin.Text = "🔓 ВОЙТИ";
                btnLogin.Location = new Point(50, yOffset);
                btnLogin.Size = new Size(320, 45);
                btnLogin.BackColor = accentColor;
                btnLogin.ForeColor = Color.Black;
                btnLogin.Font = new Font("Segoe UI", 11, FontStyle.Bold);
                btnLogin.FlatStyle = FlatStyle.Flat;
                btnLogin.Cursor = Cursors.Hand;
                btnLogin.Click += (s, e) => LoginUser(txtEmail.Text, txtPassword.Text);
                mainPanel.Controls.Add(btnLogin);
                yOffset += 60;

                LinkLabel lnkSwitch = new LinkLabel();
                lnkSwitch.Text = "Нет аккаунта? Зарегистрироваться";
                lnkSwitch.Location = new Point(110, yOffset);
                lnkSwitch.Size = new Size(200, 25);
                lnkSwitch.Font = new Font("Segoe UI", 10);
                lnkSwitch.LinkColor = accentBlue;
                lnkSwitch.TextAlign = ContentAlignment.MiddleCenter;
                lnkSwitch.Click += (s, e) => { isRegisterMode = true; DisplayLoginForm(); };
                mainPanel.Controls.Add(lnkSwitch);

                yOffset += 40;

                Button btnGuest = new Button();
                btnGuest.Text = "👤 ПРОДОЛЖИТЬ КАК ГОСТЬ";
                btnGuest.Location = new Point(50, yOffset);
                btnGuest.Size = new Size(320, 40);
                btnGuest.BackColor = Color.FromArgb(60, 60, 70);
                btnGuest.ForeColor = textColor;
                btnGuest.Font = new Font("Segoe UI", 10, FontStyle.Bold);
                btnGuest.FlatStyle = FlatStyle.Flat;
                btnGuest.Cursor = Cursors.Hand;
                btnGuest.Click += (s, e) => { currentUser = null; this.DialogResult = DialogResult.OK; this.Close(); };
                mainPanel.Controls.Add(btnGuest);
            }

            this.Controls.Add(mainPanel);
        }

        private void AddLabel(Panel panel, string text, int y)
        {
            Label label = new Label();
            label.Text = text;
            label.Location = new Point(30, y);
            label.Size = new Size(200, 25);
            label.Font = new Font("Segoe UI", 10);
            label.ForeColor = textColor;
            panel.Controls.Add(label);
        }

        private TextBox AddTextBox(Panel panel, int y)
        {
            TextBox textBox = new TextBox();
            textBox.Location = new Point(30, y);
            textBox.Size = new Size(360, 30);
            textBox.Font = new Font("Segoe UI", 11);
            textBox.BackColor = darkBg;
            textBox.ForeColor = textColor;
            textBox.BorderStyle = BorderStyle.FixedSingle;
            panel.Controls.Add(textBox);
            return textBox;
        }

        private void RegisterUser(string fullName, string email, string password, string confirmPassword)
        {
            if (string.IsNullOrWhiteSpace(fullName))
            {
                MessageBox.Show("Введите полное имя!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(email) || !email.Contains("@"))
            {
                MessageBox.Show("Введите корректный Email!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(password) || password.Length < 4)
            {
                MessageBox.Show("Пароль должен быть не менее 4 символов!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (password != confirmPassword)
            {
                MessageBox.Show("Пароли не совпадают!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string checkQuery = "SELECT COUNT(*) FROM Customers WHERE Email = ?";
                OleDbCommand checkCmd = new OleDbCommand(checkQuery, connection);
                checkCmd.Parameters.AddWithValue("?", email);
                int count = Convert.ToInt32(checkCmd.ExecuteScalar());

                if (count > 0)
                {
                    MessageBox.Show("Пользователь с таким Email уже существует!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string insertQuery = "INSERT INTO Customers (FullName, Email, Password) VALUES (?, ?, ?)";
                OleDbCommand cmd = new OleDbCommand(insertQuery, connection);
                cmd.Parameters.AddWithValue("?", fullName);
                cmd.Parameters.AddWithValue("?", email);
                cmd.Parameters.AddWithValue("?", password);

                cmd.ExecuteNonQuery();

                MessageBox.Show("Регистрация успешна! Теперь войдите в систему.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                isRegisterMode = false;
                DisplayLoginForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка регистрации: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoginUser(string email, string password)
        {
            if (string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(password))
            {
                MessageBox.Show("Введите Email и пароль!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string query = "SELECT CustomerID, FullName, Email FROM Customers WHERE Email = ? AND Password = ?";
                OleDbCommand cmd = new OleDbCommand(query, connection);
                cmd.Parameters.AddWithValue("?", email);
                cmd.Parameters.AddWithValue("?", password);

                OleDbDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    currentUser = new User();
                    currentUser.UserID = Convert.ToInt32(reader["CustomerID"]);
                    currentUser.Username = reader["FullName"].ToString();
                    currentUser.Email = reader["Email"].ToString();
                    currentUser.IsLoggedIn = true;
                    reader.Close();

                    MessageBox.Show($"Добро пожаловать, {currentUser.Username}!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    reader.Close();
                    MessageBox.Show("Неверный Email или пароль!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка входа: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (connection != null && connection.State == ConnectionState.Open)
                connection.Close();
            base.OnFormClosing(e);
        }
    }
}