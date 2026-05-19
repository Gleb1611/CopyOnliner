using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;

namespace CopyOnliner
{
    public partial class Form2 : Form
    {
        private ConsoleItem consoleItem;
        private NumericUpDown numQuantity;
        private TextBox txtCustomerName;
        private TextBox txtCustomerPhone;
        private OleDbConnection connection;

        private Color darkBg = Color.FromArgb(15, 25, 35);
        private Color panelBg = Color.FromArgb(25, 35, 45);
        private Color accentColor = Color.FromArgb(0, 255, 100);
        private Color accentBlue = Color.FromArgb(0, 150, 255);
        private Color textColor = Color.FromArgb(220, 220, 220);
        private Color cardBg = Color.FromArgb(35, 45, 55);

        public Form2(ConsoleItem item)
        {
            consoleItem = item;
            InitializeComponent();
            SetupForm();
            DisplayConsoleDetails();

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
            this.Text = "Детали консоли";
            this.Size = new Size(550, 750);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = darkBg;
        }

        private void DisplayConsoleDetails()
        {
            Panel infoPanel = new Panel();
            infoPanel.Dock = DockStyle.Fill;
            infoPanel.BackColor = panelBg;
            infoPanel.Padding = new Padding(20);
            infoPanel.AutoScroll = true;

            int yOffset = 20;

            PictureBox pictureBox = new PictureBox();
            pictureBox.Location = new Point(20, yOffset);
            pictureBox.Size = new Size(450, 300);
            pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox.BackColor = cardBg;
            pictureBox.Image = consoleItem.Image ?? CreatePlaceholderImage();
            infoPanel.Controls.Add(pictureBox);
            yOffset += 320;

            AddLabel(infoPanel, $"{consoleItem.Brand} {consoleItem.Model}", yOffset, 40, accentColor, 18, true);
            yOffset += 50;

            string priceText = consoleItem.Price > 0 ? $"💰 Цена: {consoleItem.Price:N0} ₽" : "💰 Цена не указана";
            AddLabel(infoPanel, priceText, yOffset, 30, accentBlue, 14, true);
            yOffset += 45;

            AddLabel(infoPanel, $"🏪 Магазин: {consoleItem.ShopName ?? "Не указан"}", yOffset, 25, textColor, 11, false);
            yOffset += 40;

            Panel separator = new Panel();
            separator.BackColor = accentColor;
            separator.Location = new Point(20, yOffset);
            separator.Size = new Size(450, 2);
            infoPanel.Controls.Add(separator);
            yOffset += 20;

            AddLabel(infoPanel, "📊 ХАРАКТЕРИСТИКИ", yOffset, 25, accentColor, 11, true);
            yOffset += 30;

            string specs = "";
            if (!string.IsNullOrEmpty(consoleItem.ScreenSize)) specs += $"📱 Размер экрана: {consoleItem.ScreenSize}\"\n\n";
            if (!string.IsNullOrEmpty(consoleItem.Resolution)) specs += $"🎯 Разрешение: {consoleItem.Resolution}\n\n";
            if (!string.IsNullOrEmpty(consoleItem.Storage)) specs += $"💾 Память: {consoleItem.Storage} ГБ\n\n";
            if (!string.IsNullOrEmpty(consoleItem.RAM)) specs += $"🧠 ОЗУ: {consoleItem.RAM} ГБ\n\n";
            if (!string.IsNullOrEmpty(consoleItem.Processor)) specs += $"⚙ Процессор: {consoleItem.Processor}\n\n";
            if (!string.IsNullOrEmpty(consoleItem.BatteryLife)) specs += $"🔋 Батарея: {consoleItem.BatteryLife} часов\n\n";
            if (!string.IsNullOrEmpty(consoleItem.OS)) specs += $"💿 ОС: {consoleItem.OS}\n\n";
            if (!string.IsNullOrEmpty(consoleItem.Description)) specs += $"📝 Описание:\n{consoleItem.Description}";

            if (string.IsNullOrEmpty(specs)) specs = "Информация отсутствует";

            Label lblSpecs = new Label();
            lblSpecs.Text = specs;
            lblSpecs.Location = new Point(20, yOffset);
            lblSpecs.Size = new Size(450, 150);
            lblSpecs.Font = new Font("Segoe UI", 10);
            lblSpecs.ForeColor = textColor;
            lblSpecs.AutoSize = true;
            infoPanel.Controls.Add(lblSpecs);
            yOffset += lblSpecs.Height + 30;

            AddLabel(infoPanel, "🛒 ОФОРМЛЕНИЕ ЗАКАЗА", yOffset, 30, accentColor, 12, true);
            yOffset += 40;

            AddLabel(infoPanel, "👤 Ваше имя:", yOffset, 25, textColor, 10, false);
            yOffset += 30;

            txtCustomerName = new TextBox();
            txtCustomerName.Location = new Point(20, yOffset);
            txtCustomerName.Size = new Size(450, 30);
            txtCustomerName.Font = new Font("Segoe UI", 11);
            txtCustomerName.BackColor = darkBg;
            txtCustomerName.ForeColor = textColor;
            txtCustomerName.BorderStyle = BorderStyle.FixedSingle;
            txtCustomerName.PlaceholderText = "Введите ваше имя";
            infoPanel.Controls.Add(txtCustomerName);
            yOffset += 45;

            AddLabel(infoPanel, "📞 Телефон:", yOffset, 25, textColor, 10, false);
            yOffset += 30;

            txtCustomerPhone = new TextBox();
            txtCustomerPhone.Location = new Point(20, yOffset);
            txtCustomerPhone.Size = new Size(450, 30);
            txtCustomerPhone.Font = new Font("Segoe UI", 11);
            txtCustomerPhone.BackColor = darkBg;
            txtCustomerPhone.ForeColor = textColor;
            txtCustomerPhone.BorderStyle = BorderStyle.FixedSingle;
            txtCustomerPhone.PlaceholderText = "Введите ваш телефон";
            infoPanel.Controls.Add(txtCustomerPhone);
            yOffset += 45;

            AddLabel(infoPanel, "🔢 Количество:", yOffset, 25, textColor, 10, false);
            yOffset += 30;

            numQuantity = new NumericUpDown();
            numQuantity.Location = new Point(20, yOffset);
            numQuantity.Size = new Size(120, 30);
            numQuantity.Minimum = 1;
            numQuantity.Maximum = 99;
            numQuantity.Value = 1;
            numQuantity.BackColor = darkBg;
            numQuantity.ForeColor = textColor;
            numQuantity.Font = new Font("Segoe UI", 11);
            numQuantity.ValueChanged += NumQuantity_ValueChanged;
            infoPanel.Controls.Add(numQuantity);

            Label lblTotal = new Label();
            decimal total = consoleItem.Price * numQuantity.Value;
            lblTotal.Text = total > 0 ? $"💰 Итого: {total:N0} ₽" : "💰 Итого: Цена не указана";
            lblTotal.Location = new Point(160, yOffset);
            lblTotal.Size = new Size(310, 30);
            lblTotal.Font = new Font("Segoe UI", 12, FontStyle.Bold);
            lblTotal.ForeColor = accentColor;
            lblTotal.Tag = "totalLabel";
            infoPanel.Controls.Add(lblTotal);
            yOffset += 55;

            Button btnOrder = new Button();
            btnOrder.Text = "✅ ОФОРМИТЬ ЗАКАЗ";
            btnOrder.Location = new Point(125, yOffset);
            btnOrder.Size = new Size(240, 55);
            btnOrder.BackColor = accentColor;
            btnOrder.ForeColor = Color.Black;
            btnOrder.Font = new Font("Segoe UI", 12, FontStyle.Bold);
            btnOrder.FlatStyle = FlatStyle.Flat;
            btnOrder.FlatAppearance.BorderSize = 0;
            btnOrder.Cursor = Cursors.Hand;
            btnOrder.Click += BtnOrder_Click;
            infoPanel.Controls.Add(btnOrder);

            this.Controls.Add(infoPanel);
        }

        private void NumQuantity_ValueChanged(object sender, EventArgs e)
        {
            foreach (Control control in this.Controls)
            {
                if (control is Panel panel)
                {
                    foreach (Control innerControl in panel.Controls)
                    {
                        if (innerControl is Label label && label.Tag as string == "totalLabel")
                        {
                            decimal total = consoleItem.Price * numQuantity.Value;
                            label.Text = total > 0 ? $"💰 Итого: {total:N0} ₽" : "💰 Итого: Цена не указана";
                            break;
                        }
                    }
                }
            }
        }

        private void AddLabel(Panel panel, string text, int y, int height, Color color, float fontSize, bool bold)
        {
            Label label = new Label();
            label.Text = text;
            label.Location = new Point(20, y);
            label.Size = new Size(450, height);
            label.Font = new Font("Segoe UI", fontSize, bold ? FontStyle.Bold : FontStyle.Regular);
            label.ForeColor = color;
            panel.Controls.Add(label);
        }

        private Image CreatePlaceholderImage()
        {
            Bitmap placeholder = new Bitmap(450, 300);
            using (Graphics g = Graphics.FromImage(placeholder))
            {
                g.Clear(cardBg);
                using (Font font = new Font("Segoe UI", 48, FontStyle.Bold))
                {
                    string text = consoleItem.Brand.Length > 0 ? consoleItem.Brand[0].ToString() : "?";
                    SizeF textSize = g.MeasureString(text, font);
                    g.DrawString(text, font, new SolidBrush(accentColor),
                        (placeholder.Width - textSize.Width) / 2,
                        (placeholder.Height - textSize.Height) / 2);
                }
            }
            return placeholder;
        }

        private void BtnOrder_Click(object sender, EventArgs e)
        {
            if (consoleItem.Price <= 0)
            {
                MessageBox.Show("Цена на данный товар не указана. Заказ невозможен.",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtCustomerName.Text))
            {
                MessageBox.Show("Пожалуйста, введите ваше имя!",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtCustomerName.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(txtCustomerPhone.Text))
            {
                MessageBox.Show("Пожалуйста, введите ваш телефон!",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtCustomerPhone.Focus();
                return;
            }

            decimal totalPrice = consoleItem.Price * numQuantity.Value;
            DialogResult result = MessageBox.Show(
                $"Подтвердите заказ:\n\n" +
                $"Товар: {consoleItem.Brand} {consoleItem.Model}\n" +
                $"Цена за единицу: {consoleItem.Price:N0} ₽\n" +
                $"Количество: {numQuantity.Value}\n" +
                $"Общая сумма: {totalPrice:N0} ₽\n\n" +
                $"Покупатель: {txtCustomerName.Text}\n" +
                $"Телефон: {txtCustomerPhone.Text}\n\n" +
                $"Продолжить оформление?",
                "Подтверждение заказа",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                SaveOrderToDatabase(totalPrice);
            }
        }

        private void SaveOrderToDatabase(decimal totalPrice)
        {
            try
            {
                if (connection == null || connection.State != ConnectionState.Open)
                {
                    MessageBox.Show("Нет подключения к базе данных.",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string consoleName = $"{consoleItem.Brand} {consoleItem.Model}";
                int quantity = (int)numQuantity.Value;

                // Используем квадратные скобки для зарезервированных слов
                string insertQuery = @"
                    INSERT INTO Orders ([ConsoleName], [Count], [Price], [CustomersID], [ShopID])
                    VALUES (?, ?, ?, ?, ?)";

                OleDbCommand command = new OleDbCommand(insertQuery, connection);
                command.Parameters.AddWithValue("?", consoleName);
                command.Parameters.AddWithValue("?", quantity);
                command.Parameters.AddWithValue("?", totalPrice);
                command.Parameters.AddWithValue("?", 0);  // CustomersID
                command.Parameters.AddWithValue("?", 1);  // ShopID

                int rowsAffected = command.ExecuteNonQuery();

                if (rowsAffected > 0)
                {
                    MessageBox.Show(
                        $"✅ ЗАКАЗ УСПЕШНО ОФОРМЛЕН!\n\n" +
                        $"Товар: {consoleName}\n" +
                        $"Количество: {quantity}\n" +
                        $"Сумма: {totalPrice:N0} ₽\n\n" +
                        $"Мы свяжемся с вами по телефону {txtCustomerPhone.Text}",
                        "Заказ оформлен",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Ошибка при сохранении заказа. Попробуйте еще раз.",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении заказа: {ex.Message}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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