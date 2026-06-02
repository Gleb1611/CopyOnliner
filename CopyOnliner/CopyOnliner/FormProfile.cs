using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace CopyOnliner
{
    public partial class FormProfile : Form
    {
        private User currentUser;
        private OleDbConnection connection;

        private Color darkBg = Color.FromArgb(10, 10, 10);
        private Color panelBg = Color.FromArgb(20, 20, 20);
        private Color accentColor = Color.FromArgb(0, 255, 100);
        private Color accentBlue = Color.FromArgb(0, 150, 255);
        private Color textColor = Color.FromArgb(200, 200, 200);

        public FormProfile(User user)
        {
            currentUser = user;
            InitializeComponent();
            SetupForm();

            try
            {
                connection = new OleDbConnection(Form1.connectionString);
                connection.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка подключения: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            DisplayProfile();
        }

        private void SetupForm()
        {
            this.Text = "Мой профиль";
            this.Size = new Size(600, 600);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = darkBg;
        }

        private void DisplayProfile()
        {
            Panel mainPanel = new Panel();
            mainPanel.Dock = DockStyle.Fill;
            mainPanel.BackColor = panelBg;
            mainPanel.Padding = new Padding(30);

            int yOffset = 20;

            Panel avatarPanel = new Panel();
            avatarPanel.Location = new Point(200, yOffset);
            avatarPanel.Size = new Size(100, 100);
            avatarPanel.BackColor = Color.FromArgb(0, 100, 50);
            avatarPanel.BorderStyle = BorderStyle.FixedSingle;
            mainPanel.Controls.Add(avatarPanel);

            Label lblAvatar = new Label();
            lblAvatar.Text = currentUser != null && currentUser.Username.Length > 0 ? currentUser.Username[0].ToString().ToUpper() : "?";
            lblAvatar.Location = new Point(0, 25);
            lblAvatar.Size = new Size(100, 50);
            lblAvatar.Font = new Font("Segoe UI", 36, FontStyle.Bold);
            lblAvatar.ForeColor = accentColor;
            lblAvatar.TextAlign = ContentAlignment.MiddleCenter;
            avatarPanel.Controls.Add(lblAvatar);
            yOffset += 120;

            AddInfoLabel(mainPanel, "👤 Полное имя:", currentUser?.Username ?? "Гость", yOffset);
            yOffset += 50;

            AddInfoLabel(mainPanel, "📧 Email:", currentUser?.Email ?? "Не указан", yOffset);
            yOffset += 50;

            AddInfoLabel(mainPanel, "📞 Телефон:", currentUser?.Phone ?? "Не указан", yOffset);
            yOffset += 60;

            Button btnOrders = new Button();
            btnOrders.Text = "📦 МОИ ЗАКАЗЫ";
            btnOrders.Location = new Point(50, yOffset);
            btnOrders.Size = new Size(200, 45);
            btnOrders.BackColor = accentBlue;
            btnOrders.ForeColor = Color.Black;
            btnOrders.Font = new Font("Segoe UI", 11, FontStyle.Bold);
            btnOrders.FlatStyle = FlatStyle.Flat;
            btnOrders.Cursor = Cursors.Hand;
            btnOrders.Click += BtnOrders_Click;
            mainPanel.Controls.Add(btnOrders);

            // Кнопка выхода из аккаунта
            Button btnLogout = new Button();
            btnLogout.Text = "🚪 ВЫЙТИ ИЗ АККАУНТА";
            btnLogout.Location = new Point(270, yOffset);
            btnLogout.Size = new Size(200, 45);
            btnLogout.BackColor = Color.FromArgb(180, 50, 50);
            btnLogout.ForeColor = Color.White;
            btnLogout.Font = new Font("Segoe UI", 11, FontStyle.Bold);
            btnLogout.FlatStyle = FlatStyle.Flat;
            btnLogout.Cursor = Cursors.Hand;
            btnLogout.Click += BtnLogout_Click;
            mainPanel.Controls.Add(btnLogout);
            yOffset += 60;

            Button btnClose = new Button();
            btnClose.Text = "ЗАКРЫТЬ";
            btnClose.Location = new Point(160, yOffset);
            btnClose.Size = new Size(200, 45);
            btnClose.BackColor = Color.FromArgb(60, 60, 70);
            btnClose.ForeColor = textColor;
            btnClose.Font = new Font("Segoe UI", 11, FontStyle.Bold);
            btnClose.FlatStyle = FlatStyle.Flat;
            btnClose.Cursor = Cursors.Hand;
            btnClose.Click += (s, e) => this.Close();
            mainPanel.Controls.Add(btnClose);

            this.Controls.Add(mainPanel);
        }

        private void AddInfoLabel(Panel panel, string label, string value, int y)
        {
            Label lblLabel = new Label();
            lblLabel.Text = label;
            lblLabel.Location = new Point(30, y);
            lblLabel.Size = new Size(150, 25);
            lblLabel.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblLabel.ForeColor = accentBlue;
            panel.Controls.Add(lblLabel);

            Label lblValue = new Label();
            lblValue.Text = value;
            lblValue.Location = new Point(190, y);
            lblValue.Size = new Size(300, 25);
            lblValue.Font = new Font("Segoe UI", 11);
            lblValue.ForeColor = textColor;
            panel.Controls.Add(lblValue);
        }

        private void BtnOrders_Click(object sender, EventArgs e)
        {
            if (currentUser == null)
            {
                MessageBox.Show("Вы вошли как гость. Заказы доступны только зарегистрированным пользователям.",
                    "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            FormOrders ordersForm = new FormOrders(currentUser);
            ordersForm.ShowDialog(this);
        }

        private void BtnLogout_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы уверены, что хотите выйти из аккаунта?", "Выход",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // Очищаем сессию
                Session.Logout();

                // Находим главную форму и обновляем кнопку профиля
                Form1 form1 = Application.OpenForms.OfType<Form1>().FirstOrDefault();
                if (form1 != null)
                {
                    form1.UpdateProfileButton();
                }

                MessageBox.Show("Вы успешно вышли из аккаунта.", "Выход",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Закрываем окно профиля
                this.Close();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (connection != null && connection.State == ConnectionState.Open)
            {
                connection.Close();
                connection.Dispose();
            }
            base.OnFormClosing(e);
        }
    }
}