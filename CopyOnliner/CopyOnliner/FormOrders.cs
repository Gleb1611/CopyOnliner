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

        private Color darkBg = Color.FromArgb(10, 10, 10);
        private Color panelBg = Color.FromArgb(20, 20, 20);
        private Color accentColor = Color.FromArgb(0, 255, 100);
        private Color accentBlue = Color.FromArgb(0, 150, 255);
        private Color textColor = Color.FromArgb(200, 200, 200);

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

            this.Controls.Add(listViewOrders);
        }

        private void LoadOrders()
        {
            try
            {
                listViewOrders.Items.Clear();

                // Count в квадратных скобках, так как это зарезервированное слово
                string query = "SELECT OrderID, ConsoleName, [Count], Price FROM Orders WHERE CustomersID = ? ORDER BY OrderID DESC";

                OleDbCommand cmd = new OleDbCommand(query, connection);
                cmd.Parameters.AddWithValue("?", currentUser.UserID);

                OleDbDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    ListViewItem item = new ListViewItem();
                    item.Text = reader["OrderID"].ToString();
                    item.SubItems.Add(reader["ConsoleName"].ToString());
                    item.SubItems.Add(reader["Count"].ToString()); // Здесь можно без скобок, т.к. читаем по имени

                    decimal price = reader["Price"] != DBNull.Value ? Convert.ToDecimal(reader["Price"]) : 0;
                    item.SubItems.Add($"{price:N0} ₽");

                    listViewOrders.Items.Add(item);
                }

                reader.Close();

                if (listViewOrders.Items.Count == 0)
                {
                    ShowEmptyMessage();
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

        private void ShowEmptyMessage()
        {
            Label lblEmpty = new Label();
            lblEmpty.Text = "📦 У вас пока нет заказов";
            lblEmpty.Location = new Point(250, 200);
            lblEmpty.Size = new Size(250, 30);
            lblEmpty.Font = new Font("Segoe UI", 12);
            lblEmpty.ForeColor = textColor;
            lblEmpty.TextAlign = ContentAlignment.MiddleCenter;
            this.Controls.Add(lblEmpty);
            lblEmpty.BringToFront();
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