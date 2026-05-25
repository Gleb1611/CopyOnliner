using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;

namespace CopyOnliner
{
    public partial class FormOrders : Form
    {
        private User currentUser;
        private OleDbConnection connection;
        private ListView listViewOrders;

        private Color darkBg = Color.FromArgb(10, 10, 10);        // Чистый черный
        private Color panelBg = Color.FromArgb(20, 20, 20);       // Темно-серый
        private Color accentColor = Color.FromArgb(0, 255, 100);   // Неоново-зеленый
        private Color accentBlue = Color.FromArgb(0, 150, 255);    // Неоново-синий
        private Color textColor = Color.FromArgb(200, 200, 200);   // Светло-серый текст

        public FormOrders(User user)
        {
            currentUser = user;
            InitializeComponent();
            SetupForm();

            try
            {
                connection = new OleDbConnection(Form1.connectionString);
                connection.Open();
                LoadOrders();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка подключения: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetupForm()
        {
            this.Text = "Мои заказы";
            this.Size = new Size(750, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = darkBg;

            listViewOrders = new ListView();
            listViewOrders.Dock = DockStyle.Fill;
            listViewOrders.View = View.Details;
            listViewOrders.FullRowSelect = true;
            listViewOrders.GridLines = true;
            listViewOrders.BackColor = panelBg;
            listViewOrders.ForeColor = textColor;
            listViewOrders.Font = new Font("Segoe UI", 10);

            listViewOrders.Columns.Add("№ заказа", 80, HorizontalAlignment.Center);
            listViewOrders.Columns.Add("Товар", 250, HorizontalAlignment.Left);
            listViewOrders.Columns.Add("Количество", 100, HorizontalAlignment.Center);
            listViewOrders.Columns.Add("Сумма", 120, HorizontalAlignment.Right);
            listViewOrders.Columns.Add("Дата", 150, HorizontalAlignment.Center);

            this.Controls.Add(listViewOrders);
        }

        private void LoadOrders()
        {
            try
            {
                listViewOrders.Items.Clear();

                string query = @"
                    SELECT OrderID, ConsoleName, Count, Price, OrderDate 
                    FROM Orders 
                    WHERE CustomersID = ?
                    ORDER BY OrderDate DESC";

                OleDbCommand cmd = new OleDbCommand(query, connection);
                cmd.Parameters.AddWithValue("?", currentUser.UserID);
                OleDbDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    ListViewItem item = new ListViewItem();
                    item.Text = reader["OrderID"].ToString();
                    item.SubItems.Add(reader["ConsoleName"].ToString());
                    item.SubItems.Add(reader["Count"].ToString());
                    item.SubItems.Add($"{Convert.ToDecimal(reader["Price"]):N0} ₽");
                    item.SubItems.Add(reader["OrderDate"] != DBNull.Value ? Convert.ToDateTime(reader["OrderDate"]).ToString("dd.MM.yyyy HH:mm") : "—");

                    listViewOrders.Items.Add(item);
                }

                reader.Close();

                if (listViewOrders.Items.Count == 0)
                {
                    Label lblEmpty = new Label();
                    lblEmpty.Text = "📦 У вас пока нет заказов";
                    lblEmpty.Location = new Point(250, 200);
                    lblEmpty.Size = new Size(250, 30);
                    lblEmpty.Font = new Font("Segoe UI", 12);
                    lblEmpty.ForeColor = textColor;
                    lblEmpty.TextAlign = ContentAlignment.MiddleCenter;
                    this.Controls.Add(lblEmpty);
                }

                foreach (ColumnHeader col in listViewOrders.Columns)
                {
                    col.Width = -2;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки заказов: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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