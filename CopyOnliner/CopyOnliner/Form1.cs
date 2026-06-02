using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CopyOnliner
{
    public partial class Form1 : Form
    {
        public static string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=OnlinerConsoles.accdb;";

        private OleDbConnection myConnection;
        private Panel searchPanel;
        private TextBox txtSearchModel;
        private Button btnProfile;
        private Button btnReset;
        private Label lblResults;
        private System.Windows.Forms.Timer searchTimer;
        private ListBox listBoxBrands;
        private CheckedListBox chkListShops;
        private Label lblBrand, lblModel, lblPrice, lblStorage, lblRAM, lblProcessor, lblShops, lblOS, lblSort, lblColor;
        private NumericUpDown numPriceFrom, numPriceTo;
        private CheckedListBox chkListStorage, chkListRAM;
        private ComboBox cmbProcessor;
        private ComboBox cmbOS;
        private CheckedListBox chkListColors;
        private CheckedListBox chkListSortOptions;
        private RadioButton rbAscending, rbDescending;
        private Button btnApply;

        // Путь к папке с изображениями
        private string imagesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Image");

        private Color darkBg = Color.FromArgb(245, 245, 245);
        private Color panelBg = Color.FromArgb(248, 248, 248);
        private Color accentColor = Color.FromArgb(0, 120, 215);
        private Color textColor = Color.FromArgb(51, 51, 51);

        public Form1()
        {
            InitializeComponent();
            SetupListView();
            SetupForm();
            InitializeSearchPanel();

            UpdateProfileButton();

            // Создаем папку Image если её нет
            if (!Directory.Exists(imagesPath))
            {
                Directory.CreateDirectory(imagesPath);
            }

            try
            {
                myConnection = new OleDbConnection(connectionString);
                myConnection.Open();
                LoadFilters();
                LoadData();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка подключения: {ex.Message}\n\nДобавляю демо-данные...", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                AddDemoData();
            }
        }

        private void SetupForm()
        {
            this.Text = "Каталог игровых консолей";
            this.BackColor = darkBg;
            this.WindowState = FormWindowState.Maximized;
            this.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            searchTimer = new System.Windows.Forms.Timer();
            searchTimer.Interval = 500;
            searchTimer.Tick += SearchTimer_Tick;
        }

        public void UpdateProfileButton()
        {
            if (btnProfile != null)
            {
                if (Session.IsLoggedIn)
                {
                    btnProfile.Text = $"👤 {Session.CurrentUser.Username}";
                    btnProfile.BackColor = accentColor;
                    btnProfile.ForeColor = Color.White;
                }
                else
                {
                    btnProfile.Text = "👤 ВОЙТИ";
                    btnProfile.BackColor = Color.FromArgb(240, 240, 240);
                    btnProfile.ForeColor = Color.FromArgb(102, 102, 102);
                }
            }
        }

        private void LoadFilters()
        {
            try
            {
                if (myConnection == null || myConnection.State != ConnectionState.Open)
                    return;

                string brandQuery = "SELECT DISTINCT Brand FROM Consoles WHERE Brand IS NOT NULL AND Brand <> '' ORDER BY Brand";
                OleDbCommand cmd = new OleDbCommand(brandQuery, myConnection);
                OleDbDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    string brand = reader["Brand"].ToString();
                    if (!string.IsNullOrEmpty(brand))
                        listBoxBrands.Items.Add(brand);
                }
                reader.Close();

                string storageQuery = "SELECT DISTINCT Storage FROM Consoles WHERE Storage IS NOT NULL AND Storage <> '' AND Storage <> 'N/A' ORDER BY Storage";
                cmd = new OleDbCommand(storageQuery, myConnection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string storage = reader["Storage"].ToString();
                    if (!string.IsNullOrEmpty(storage))
                        chkListStorage.Items.Add(storage);
                }
                reader.Close();

                string ramQuery = "SELECT DISTINCT RAM FROM Consoles WHERE RAM IS NOT NULL AND RAM <> '' AND RAM <> 'N/A' ORDER BY RAM";
                cmd = new OleDbCommand(ramQuery, myConnection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string ram = reader["RAM"].ToString();
                    if (!string.IsNullOrEmpty(ram))
                        chkListRAM.Items.Add(ram);
                }
                reader.Close();

                string processorQuery = "SELECT DISTINCT Processor FROM Consoles WHERE Processor IS NOT NULL AND Processor <> '' ORDER BY Processor";
                cmd = new OleDbCommand(processorQuery, myConnection);
                reader = cmd.ExecuteReader();
                cmbProcessor.Items.Add("Все");
                while (reader.Read())
                {
                    string processor = reader["Processor"].ToString();
                    if (!string.IsNullOrEmpty(processor))
                        cmbProcessor.Items.Add(processor);
                }
                reader.Close();
                cmbProcessor.SelectedIndex = 0;

                string osQuery = "SELECT DISTINCT OS FROM Consoles WHERE OS IS NOT NULL AND OS <> '' ORDER BY OS";
                cmd = new OleDbCommand(osQuery, myConnection);
                reader = cmd.ExecuteReader();
                if (cmbOS != null)
                {
                    cmbOS.Items.Clear();
                    cmbOS.Items.Add("Все");
                    while (reader.Read())
                    {
                        string os = reader["OS"].ToString();
                        if (!string.IsNullOrEmpty(os))
                            cmbOS.Items.Add(os);
                    }
                    cmbOS.SelectedIndex = 0;
                }
                reader.Close();

                string colorQuery = "SELECT DISTINCT Color FROM Consoles WHERE Color IS NOT NULL AND Color <> '' ORDER BY Color";
                cmd = new OleDbCommand(colorQuery, myConnection);
                reader = cmd.ExecuteReader();
                if (chkListColors != null)
                {
                    chkListColors.Items.Clear();
                    while (reader.Read())
                    {
                        string color = reader["Color"].ToString();
                        if (!string.IsNullOrEmpty(color))
                            chkListColors.Items.Add(color);
                    }
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка загрузки фильтров: {ex.Message}");
                listBoxBrands.Items.AddRange(new string[] { "Sony", "Microsoft", "Nintendo", "Valve", "ASUS" });
                chkListStorage.Items.AddRange(new string[] { "64", "128", "256", "512", "825", "1024" });
                chkListRAM.Items.AddRange(new string[] { "4", "8", "16", "32" });
                cmbProcessor.Items.AddRange(new string[] { "Все", "AMD Zen 2", "AMD APU", "AMD Z1 Extreme", "NVIDIA Tegra" });
                if (cmbOS != null) cmbOS.Items.AddRange(new string[] { "Все", "PlayStation OS", "Xbox OS", "Nintendo OS", "SteamOS", "Windows 11" });
                if (chkListColors != null) chkListColors.Items.AddRange(new string[] { "Черный", "Белый", "Серый", "Красный", "Синий", "Фиолетовый" });
                cmbProcessor.SelectedIndex = 0;
                if (cmbOS != null) cmbOS.SelectedIndex = 0;
            }
        }

        private string BuildOrderByClause()
        {
            List<string> orderByFields = new List<string>();
            string sortOrder = rbAscending.Checked ? "ASC" : "DESC";

            foreach (var item in chkListSortOptions.CheckedItems)
            {
                string field = item.ToString();
                switch (field)
                {
                    case "Бренд":
                        orderByFields.Add($"Brand {sortOrder}");
                        break;
                    case "Модель":
                        orderByFields.Add($"Model {sortOrder}");
                        break;
                    case "Цена":
                        orderByFields.Add($"Price {sortOrder}");
                        break;
                    case "Размер экрана":
                        orderByFields.Add($"ScreenSize {sortOrder}");
                        break;
                    case "Объём памяти":
                        orderByFields.Add($"Storage {sortOrder}");
                        break;
                    case "ОЗУ":
                        orderByFields.Add($"RAM {sortOrder}");
                        break;
                    case "Процессор":
                        orderByFields.Add($"Processor {sortOrder}");
                        break;
                    case "ОС":
                        orderByFields.Add($"OS {sortOrder}");
                        break;
                    case "Цвет":
                        orderByFields.Add($"Color {sortOrder}");
                        break;
                }
            }

            return orderByFields.Count > 0 ? string.Join(", ", orderByFields) : "";
        }

        private void InitializeSearchPanel()
        {
            searchPanel = new Panel();
            searchPanel.Dock = DockStyle.Left;
            searchPanel.Width = 340;
            searchPanel.BackColor = panelBg;
            searchPanel.Padding = new Padding(15);
            searchPanel.ForeColor = textColor;
            searchPanel.AutoScroll = true;

            int yOffset = 10;

            Panel headerPanel = new Panel();
            headerPanel.Height = 50;
            headerPanel.Dock = DockStyle.Top;
            headerPanel.BackColor = Color.White;
            headerPanel.Padding = new Padding(10, 5, 10, 5);

            btnProfile = new Button();
            btnProfile.Text = "👤 ВОЙТИ";
            btnProfile.Size = new Size(100, 35);
            btnProfile.Location = new Point(210, 8);
            btnProfile.BackColor = Color.FromArgb(240, 240, 240);
            btnProfile.ForeColor = Color.FromArgb(102, 102, 102);
            btnProfile.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            btnProfile.FlatStyle = FlatStyle.Flat;
            btnProfile.FlatAppearance.BorderSize = 1;
            btnProfile.FlatAppearance.BorderColor = Color.FromArgb(221, 221, 221);
            btnProfile.Cursor = Cursors.Hand;
            btnProfile.Click += BtnProfile_Click;

            headerPanel.Controls.Add(btnProfile);
            searchPanel.Controls.Add(headerPanel);

            yOffset = 60;

            Label lblFilters = new Label();
            lblFilters.Text = "Фильтры";
            lblFilters.Font = new Font("Segoe UI", 12, FontStyle.Bold);
            lblFilters.ForeColor = textColor;
            lblFilters.Location = new Point(15, yOffset);
            lblFilters.Size = new Size(310, 25);
            searchPanel.Controls.Add(lblFilters);
            yOffset += 35;

            // Бренд
            lblBrand = new Label();
            lblBrand.Text = "Бренд";
            lblBrand.Location = new Point(15, yOffset);
            lblBrand.Size = new Size(310, 20);
            lblBrand.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            lblBrand.ForeColor = textColor;
            searchPanel.Controls.Add(lblBrand);
            yOffset += 22;

            listBoxBrands = new ListBox();
            listBoxBrands.Location = new Point(15, yOffset);
            listBoxBrands.Size = new Size(310, 80);
            listBoxBrands.BackColor = Color.White;
            listBoxBrands.ForeColor = textColor;
            listBoxBrands.BorderStyle = BorderStyle.FixedSingle;
            listBoxBrands.SelectionMode = SelectionMode.MultiExtended;
            listBoxBrands.Font = new Font("Segoe UI", 9);
            searchPanel.Controls.Add(listBoxBrands);
            yOffset += 90;

            // Модель
            lblModel = new Label();
            lblModel.Text = "Модель";
            lblModel.Location = new Point(15, yOffset);
            lblModel.Size = new Size(310, 20);
            lblModel.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            lblModel.ForeColor = textColor;
            searchPanel.Controls.Add(lblModel);
            yOffset += 22;

            txtSearchModel = new TextBox();
            txtSearchModel.Location = new Point(15, yOffset);
            txtSearchModel.Size = new Size(310, 30);
            txtSearchModel.Font = new Font("Segoe UI", 10);
            txtSearchModel.BackColor = Color.White;
            txtSearchModel.ForeColor = textColor;
            txtSearchModel.BorderStyle = BorderStyle.FixedSingle;
            txtSearchModel.PlaceholderText = "Поиск по названию...";
            txtSearchModel.TextChanged += TxtSearch_TextChanged;
            searchPanel.Controls.Add(txtSearchModel);
            yOffset += 40;

            // Цена
            lblPrice = new Label();
            lblPrice.Text = "Цена, руб.";
            lblPrice.Location = new Point(15, yOffset);
            lblPrice.Size = new Size(310, 20);
            lblPrice.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            lblPrice.ForeColor = textColor;
            searchPanel.Controls.Add(lblPrice);
            yOffset += 22;

            Label lblPriceFrom = new Label();
            lblPriceFrom.Text = "от";
            lblPriceFrom.Location = new Point(15, yOffset);
            lblPriceFrom.Size = new Size(20, 27);
            lblPriceFrom.Font = new Font("Segoe UI", 9);
            lblPriceFrom.ForeColor = textColor;
            searchPanel.Controls.Add(lblPriceFrom);

            numPriceFrom = new NumericUpDown();
            numPriceFrom.Location = new Point(35, yOffset);
            numPriceFrom.Size = new Size(125, 27);
            numPriceFrom.BackColor = Color.White;
            numPriceFrom.ForeColor = textColor;
            numPriceFrom.Maximum = 1000000;
            numPriceFrom.ThousandsSeparator = true;
            numPriceFrom.Value = 0;
            searchPanel.Controls.Add(numPriceFrom);

            Label lblPriceTo = new Label();
            lblPriceTo.Text = "до";
            lblPriceTo.Location = new Point(170, yOffset);
            lblPriceTo.Size = new Size(20, 27);
            lblPriceTo.Font = new Font("Segoe UI", 9);
            lblPriceTo.ForeColor = textColor;
            searchPanel.Controls.Add(lblPriceTo);

            numPriceTo = new NumericUpDown();
            numPriceTo.Location = new Point(190, yOffset);
            numPriceTo.Size = new Size(135, 27);
            numPriceTo.BackColor = Color.White;
            numPriceTo.ForeColor = textColor;
            numPriceTo.Maximum = 1000000;
            numPriceTo.ThousandsSeparator = true;
            numPriceTo.Value = 0;
            searchPanel.Controls.Add(numPriceTo);
            yOffset += 40;

            // Память
            lblStorage = new Label();
            lblStorage.Text = "Объём памяти, ГБ";
            lblStorage.Location = new Point(15, yOffset);
            lblStorage.Size = new Size(310, 20);
            lblStorage.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            lblStorage.ForeColor = textColor;
            searchPanel.Controls.Add(lblStorage);
            yOffset += 22;

            chkListStorage = new CheckedListBox();
            chkListStorage.Location = new Point(15, yOffset);
            chkListStorage.Size = new Size(310, 60);
            chkListStorage.BackColor = Color.White;
            chkListStorage.ForeColor = textColor;
            chkListStorage.BorderStyle = BorderStyle.FixedSingle;
            chkListStorage.CheckOnClick = true;
            chkListStorage.Font = new Font("Segoe UI", 9);
            searchPanel.Controls.Add(chkListStorage);
            yOffset += 70;

            // ОЗУ
            lblRAM = new Label();
            lblRAM.Text = "ОЗУ, ГБ";
            lblRAM.Location = new Point(15, yOffset);
            lblRAM.Size = new Size(310, 20);
            lblRAM.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            lblRAM.ForeColor = textColor;
            searchPanel.Controls.Add(lblRAM);
            yOffset += 22;

            chkListRAM = new CheckedListBox();
            chkListRAM.Location = new Point(15, yOffset);
            chkListRAM.Size = new Size(310, 60);
            chkListRAM.BackColor = Color.White;
            chkListRAM.ForeColor = textColor;
            chkListRAM.BorderStyle = BorderStyle.FixedSingle;
            chkListRAM.CheckOnClick = true;
            chkListRAM.Font = new Font("Segoe UI", 9);
            searchPanel.Controls.Add(chkListRAM);
            yOffset += 70;

            // Процессор
            lblProcessor = new Label();
            lblProcessor.Text = "Процессор";
            lblProcessor.Location = new Point(15, yOffset);
            lblProcessor.Size = new Size(310, 20);
            lblProcessor.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            lblProcessor.ForeColor = textColor;
            searchPanel.Controls.Add(lblProcessor);
            yOffset += 22;

            cmbProcessor = new ComboBox();
            cmbProcessor.Location = new Point(15, yOffset);
            cmbProcessor.Size = new Size(310, 28);
            cmbProcessor.Font = new Font("Segoe UI", 10);
            cmbProcessor.BackColor = Color.White;
            cmbProcessor.ForeColor = textColor;
            cmbProcessor.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbProcessor.FlatStyle = FlatStyle.Flat;
            searchPanel.Controls.Add(cmbProcessor);
            yOffset += 40;

            // ОС
            lblOS = new Label();
            lblOS.Text = "Операционная система";
            lblOS.Location = new Point(15, yOffset);
            lblOS.Size = new Size(310, 20);
            lblOS.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            lblOS.ForeColor = textColor;
            searchPanel.Controls.Add(lblOS);
            yOffset += 22;

            cmbOS = new ComboBox();
            cmbOS.Location = new Point(15, yOffset);
            cmbOS.Size = new Size(310, 28);
            cmbOS.Font = new Font("Segoe UI", 10);
            cmbOS.BackColor = Color.White;
            cmbOS.ForeColor = textColor;
            cmbOS.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbOS.FlatStyle = FlatStyle.Flat;
            searchPanel.Controls.Add(cmbOS);
            yOffset += 40;

            // Цвета
            lblColor = new Label();
            lblColor.Text = "Цвет";
            lblColor.Location = new Point(15, yOffset);
            lblColor.Size = new Size(310, 20);
            lblColor.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            lblColor.ForeColor = textColor;
            searchPanel.Controls.Add(lblColor);
            yOffset += 22;

            chkListColors = new CheckedListBox();
            chkListColors.Location = new Point(15, yOffset);
            chkListColors.Size = new Size(310, 60);
            chkListColors.BackColor = Color.White;
            chkListColors.ForeColor = textColor;
            chkListColors.BorderStyle = BorderStyle.FixedSingle;
            chkListColors.CheckOnClick = true;
            chkListColors.Font = new Font("Segoe UI", 9);
            searchPanel.Controls.Add(chkListColors);
            yOffset += 70;

            // Сортировка
            lblSort = new Label();
            lblSort.Text = "Сортировка (выберите параметры)";
            lblSort.Location = new Point(15, yOffset);
            lblSort.Size = new Size(310, 20);
            lblSort.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            lblSort.ForeColor = textColor;
            searchPanel.Controls.Add(lblSort);
            yOffset += 22;

            chkListSortOptions = new CheckedListBox();
            chkListSortOptions.Location = new Point(15, yOffset);
            chkListSortOptions.Size = new Size(310, 100);
            chkListSortOptions.BackColor = Color.White;
            chkListSortOptions.ForeColor = textColor;
            chkListSortOptions.BorderStyle = BorderStyle.FixedSingle;
            chkListSortOptions.CheckOnClick = true;
            chkListSortOptions.Font = new Font("Segoe UI", 9);
            chkListSortOptions.Items.AddRange(new string[] {
                "Бренд", "Модель", "Цена", "Размер экрана", "Объём памяти", "ОЗУ", "Процессор", "ОС", "Цвет"
            });
            searchPanel.Controls.Add(chkListSortOptions);
            yOffset += 110;

            // Направление сортировки
            rbAscending = new RadioButton();
            rbAscending.Text = "По возрастанию ↑";
            rbAscending.Location = new Point(15, yOffset);
            rbAscending.Size = new Size(150, 25);
            rbAscending.ForeColor = textColor;
            rbAscending.Checked = true;
            searchPanel.Controls.Add(rbAscending);

            rbDescending = new RadioButton();
            rbDescending.Text = "По убыванию ↓";
            rbDescending.Location = new Point(170, yOffset);
            rbDescending.Size = new Size(150, 25);
            rbDescending.ForeColor = textColor;
            searchPanel.Controls.Add(rbDescending);
            yOffset += 35;

            // Кнопка применения
            btnApply = new Button();
            btnApply.Text = "ПОКАЗАТЬ ТОВАРЫ";
            btnApply.Location = new Point(15, yOffset);
            btnApply.Size = new Size(310, 45);
            btnApply.BackColor = accentColor;
            btnApply.ForeColor = Color.White;
            btnApply.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            btnApply.FlatStyle = FlatStyle.Flat;
            btnApply.Cursor = Cursors.Hand;
            btnApply.Click += BtnApply_Click;
            searchPanel.Controls.Add(btnApply);
            yOffset += 55;

            // Кнопка сброса
            btnReset = new Button();
            btnReset.Text = "СБРОСИТЬ ФИЛЬТРЫ";
            btnReset.Location = new Point(15, yOffset);
            btnReset.Size = new Size(310, 40);
            btnReset.BackColor = Color.White;
            btnReset.ForeColor = Color.FromArgb(102, 102, 102);
            btnReset.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            btnReset.FlatStyle = FlatStyle.Flat;
            btnReset.FlatAppearance.BorderSize = 1;
            btnReset.FlatAppearance.BorderColor = Color.FromArgb(221, 221, 221);
            btnReset.Cursor = Cursors.Hand;
            btnReset.Click += BtnReset_Click;
            searchPanel.Controls.Add(btnReset);
            yOffset += 55;

            lblResults = new Label();
            lblResults.Text = "Найдено: 0 товаров";
            lblResults.Location = new Point(15, yOffset);
            lblResults.Size = new Size(310, 30);
            lblResults.Font = new Font("Segoe UI", 10, FontStyle.Regular);
            lblResults.ForeColor = Color.FromArgb(136, 136, 136);
            lblResults.TextAlign = ContentAlignment.MiddleCenter;
            searchPanel.Controls.Add(lblResults);

            this.Controls.Add(searchPanel);
        }

        private void SetupListView()
        {
            imageList1.ImageSize = new Size(80, 80);
            imageList1.ColorDepth = ColorDepth.Depth32Bit;

            listView1.Dock = DockStyle.Fill;
            listView1.View = View.Details;
            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            listView1.MultiSelect = false;
            listView1.BackColor = Color.White;
            listView1.ForeColor = textColor;
            listView1.Font = new Font("Segoe UI", 10);
            listView1.BorderStyle = BorderStyle.None;

            listView1.Columns.Clear();
            listView1.Columns.Add("", 40, HorizontalAlignment.Center);
            listView1.Columns.Add("Товар", 280, HorizontalAlignment.Left);
            listView1.Columns.Add("Характеристики", 350, HorizontalAlignment.Left);
            listView1.Columns.Add("Цена", 120, HorizontalAlignment.Right);
            listView1.Columns.Add("Магазин", 120, HorizontalAlignment.Left);
            listView1.Columns.Add("Действия", 80, HorizontalAlignment.Center);

            listView1.DoubleClick += ListView1_DoubleClick;
        }

        private Image LoadImageByConsoleId(int consoleId)
        {
            if (!Directory.Exists(imagesPath))
            {
                Directory.CreateDirectory(imagesPath);
                return null;
            }

            string[] possibleExtensions = { ".png", ".jpg", ".jpeg", ".bmp", ".gif" };

            foreach (string ext in possibleExtensions)
            {
                string imagePath = Path.Combine(imagesPath, $"{consoleId}{ext}");
                if (File.Exists(imagePath))
                {
                    try
                    {
                        using (var fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                        {
                            var image = Image.FromStream(fs);
                            return new Bitmap(image, 80, 80);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка загрузки {imagePath}: {ex.Message}");
                    }
                }
            }

            return null;
        }

        private Image CreateBrandPlaceholderImage(string brand)
        {
            Bitmap placeholder = new Bitmap(80, 80);
            using (Graphics g = Graphics.FromImage(placeholder))
            {
                g.Clear(Color.FromArgb(245, 245, 245));
                using (Font font = new Font("Segoe UI", 24, FontStyle.Bold))
                {
                    string text = !string.IsNullOrEmpty(brand) ? brand[0].ToString() : "?";
                    SizeF textSize = g.MeasureString(text, font);
                    g.DrawString(text, font, new SolidBrush(Color.FromArgb(200, 200, 200)),
                        (placeholder.Width - textSize.Width) / 2,
                        (placeholder.Height - textSize.Height) / 2);
                }
            }
            return placeholder;
        }

        private void LoadData()
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                listView1.BeginUpdate();
                listView1.Items.Clear();
                imageList1.Images.Clear();

                Dictionary<int, int> imageCache = new Dictionary<int, int>();

                string query = @"
                    SELECT 
                        ConsoleID, Brand, Model, Description, 
                        OS, Weight, BatteryLife, ScreenSize, Resolution,
                        Storage, RAM, Processor, Price, Color
                    FROM Consoles
                    WHERE 1=1";

                var selectedBrands = listBoxBrands.SelectedItems;
                if (selectedBrands.Count > 0)
                {
                    string brands = "";
                    foreach (string brand in selectedBrands)
                    {
                        brands += $"'{brand.Replace("'", "''")}',";
                    }
                    brands = brands.TrimEnd(',');
                    query += $" AND Brand IN ({brands})";
                }

                if (!string.IsNullOrWhiteSpace(txtSearchModel.Text))
                {
                    query += $" AND Model LIKE '%{txtSearchModel.Text.Replace("'", "''")}%'";
                }

                if (numPriceFrom.Value > 0)
                {
                    query += $" AND Price >= {numPriceFrom.Value}";
                }
                if (numPriceTo.Value > 0)
                {
                    query += $" AND Price <= {numPriceTo.Value}";
                }

                var selectedStorage = new List<string>();
                foreach (var item in chkListStorage.CheckedItems)
                {
                    selectedStorage.Add(item.ToString());
                }
                if (selectedStorage.Count > 0)
                {
                    string storageList = "";
                    foreach (string storage in selectedStorage)
                    {
                        storageList += $"'{storage.Replace("'", "''")}',";
                    }
                    storageList = storageList.TrimEnd(',');
                    query += $" AND Storage IN ({storageList})";
                }

                var selectedRAM = new List<string>();
                foreach (var item in chkListRAM.CheckedItems)
                {
                    selectedRAM.Add(item.ToString());
                }
                if (selectedRAM.Count > 0)
                {
                    string ramList = "";
                    foreach (string ram in selectedRAM)
                    {
                        ramList += $"'{ram.Replace("'", "''")}',";
                    }
                    ramList = ramList.TrimEnd(',');
                    query += $" AND RAM IN ({ramList})";
                }

                if (cmbProcessor.SelectedIndex > 0 && cmbProcessor.SelectedItem.ToString() != "Все")
                {
                    query += $" AND Processor LIKE '%{cmbProcessor.SelectedItem.ToString().Replace("'", "''")}%'";
                }

                if (cmbOS != null && cmbOS.SelectedIndex > 0 && cmbOS.SelectedItem.ToString() != "Все")
                {
                    query += $" AND OS LIKE '%{cmbOS.SelectedItem.ToString().Replace("'", "''")}%'";
                }

                var selectedColors = new List<string>();
                foreach (var item in chkListColors.CheckedItems)
                {
                    selectedColors.Add(item.ToString());
                }
                if (selectedColors.Count > 0)
                {
                    string colorList = "";
                    foreach (string color in selectedColors)
                    {
                        colorList += $"'{color.Replace("'", "''")}',";
                    }
                    colorList = colorList.TrimEnd(',');
                    query += $" AND Color IN ({colorList})";
                }

                string orderBy = BuildOrderByClause();
                if (!string.IsNullOrEmpty(orderBy))
                {
                    query += $" ORDER BY {orderBy}";
                }
                else
                {
                    query += " ORDER BY ConsoleID";
                }

                if (myConnection == null || myConnection.State != ConnectionState.Open)
                {
                    AddDemoData();
                    return;
                }

                OleDbCommand command = new OleDbCommand(query, myConnection);
                OleDbDataReader reader = command.ExecuteReader();

                int imageIndex = 0;
                var processedConsoles = new HashSet<int>();

                while (reader.Read())
                {
                    int consoleId = Convert.ToInt32(reader["ConsoleID"]);

                    if (processedConsoles.Contains(consoleId))
                        continue;
                    processedConsoles.Add(consoleId);

                    ConsoleItem console = new ConsoleItem();
                    console.ConsoleId = consoleId;
                    console.Brand = reader["Brand"]?.ToString() ?? "";
                    console.Model = reader["Model"]?.ToString() ?? "";
                    console.Description = reader["Description"]?.ToString() ?? "";
                    console.OS = reader["OS"] != DBNull.Value ? reader["OS"].ToString() : "";
                    console.Weight = reader["Weight"] != DBNull.Value ? reader["Weight"].ToString() : "";
                    console.BatteryLife = reader["BatteryLife"] != DBNull.Value ? reader["BatteryLife"].ToString() : "";
                    console.ScreenSize = reader["ScreenSize"] != DBNull.Value ? reader["ScreenSize"].ToString() : "";
                    console.Resolution = reader["Resolution"] != DBNull.Value ? reader["Resolution"].ToString() : "";
                    console.Storage = reader["Storage"] != DBNull.Value ? reader["Storage"].ToString() : "";
                    console.RAM = reader["RAM"] != DBNull.Value ? reader["RAM"].ToString() : "";
                    console.Processor = reader["Processor"] != DBNull.Value ? reader["Processor"].ToString() : "";
                    console.Price = reader["Price"] != DBNull.Value ? Convert.ToDecimal(reader["Price"]) : 0;
                    console.Color = reader["Color"] != DBNull.Value ? reader["Color"].ToString() : "";
                    console.ShopName = "В наличии";

                    if (!imageCache.ContainsKey(consoleId))
                    {
                        Image img = LoadImageByConsoleId(consoleId);
                        if (img == null)
                        {
                            img = CreateBrandPlaceholderImage(console.Brand);
                        }
                        imageList1.Images.Add(img);
                        imageCache[consoleId] = imageIndex;
                        imageIndex++;
                        console.Image = img;
                    }
                    else
                    {
                        int cachedIndex = imageCache[consoleId];
                        console.Image = imageList1.Images[cachedIndex];
                    }

                    int actualImageIndex = imageCache[consoleId];

                    string specs = "";
                    if (!string.IsNullOrEmpty(console.ScreenSize) && console.ScreenSize != "N/A" && console.ScreenSize != "0")
                        specs += $"📱 {console.ScreenSize}\" • ";
                    if (!string.IsNullOrEmpty(console.Resolution))
                        specs += $"🎯 {console.Resolution} • ";
                    if (!string.IsNullOrEmpty(console.Storage) && console.Storage != "N/A" && console.Storage != "0")
                        specs += $"💾 {console.Storage} ГБ • ";
                    if (!string.IsNullOrEmpty(console.RAM) && console.RAM != "N/A" && console.RAM != "0")
                        specs += $"🧠 {console.RAM} ГБ • ";
                    if (!string.IsNullOrEmpty(console.Processor))
                        specs += $"⚙ {console.Processor} • ";
                    if (!string.IsNullOrEmpty(console.BatteryLife) && console.BatteryLife != "N/A" && console.BatteryLife != "0")
                        specs += $"🔋 {console.BatteryLife} ч";

                    specs = specs.TrimEnd(' ', '•');

                    if (string.IsNullOrEmpty(specs))
                        specs = "Характеристики не указаны";

                    string priceText = console.Price > 0 ? $"{console.Price:N0} ₽" : "Цена не указана";
                    string productName = $"{console.Brand} {console.Model}";

                    ListViewItem item = new ListViewItem();
                    item.Text = "";
                    item.ImageIndex = actualImageIndex;
                    item.Tag = console;

                    item.SubItems.Add(productName);
                    item.SubItems.Add(specs);
                    item.SubItems.Add(priceText);
                    item.SubItems.Add(console.ShopName);
                    item.SubItems.Add("Купить");

                    listView1.Items.Add(item);
                }

                reader.Close();
                listView1.EndUpdate();
                this.Cursor = Cursors.Default;
                lblResults.Text = $"Найдено: {processedConsoles.Count} товаров";
                this.Text = $"Каталог игровых консолей — {processedConsoles.Count} товаров";
            }
            catch (Exception ex)
            {
                listView1.EndUpdate();
                this.Cursor = Cursors.Default;
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AddDemoData();
            }
        }

        private void AddDemoData()
        {
            listView1.BeginUpdate();
            listView1.Items.Clear();
            imageList1.Images.Clear();

            Dictionary<int, int> imageCache = new Dictionary<int, int>();
            int imageIndex = 0;

            var demos = new[]
            {
                new { Id = 1, Brand = "Sony", Model = "PlayStation 5", Price = 49999, ScreenSize = "N/A", Resolution = "4K", Storage = "825", RAM = "16", Processor = "AMD Zen 2", BatteryLife = "N/A", OS = "PlayStation OS", Color = "Белый", Shop = "GAMEPARK", Description = "Новейшая консоль Sony" },
                new { Id = 2, Brand = "Microsoft", Model = "Xbox Series X", Price = 45999, ScreenSize = "N/A", Resolution = "4K", Storage = "1024", RAM = "16", Processor = "AMD Zen 2", BatteryLife = "N/A", OS = "Xbox OS", Color = "Черный", Shop = "Agroup", Description = "Самая мощная консоль Xbox" },
                new { Id = 3, Brand = "Nintendo", Model = "Switch OLED", Price = 29999, ScreenSize = "7", Resolution = "1280x720", Storage = "64", RAM = "4", Processor = "NVIDIA Tegra", BatteryLife = "4-9", OS = "Nintendo OS", Color = "Красный", Shop = "Newton", Description = "Гибридная консоль" },
                new { Id = 4, Brand = "Valve", Model = "Steam Deck", Price = 39999, ScreenSize = "7", Resolution = "1280x800", Storage = "512", RAM = "16", Processor = "AMD APU", BatteryLife = "2-8", OS = "SteamOS", Color = "Черный", Shop = "GAMEPARK", Description = "Портативный компьютер" },
                new { Id = 5, Brand = "ASUS", Model = "ROG Ally", Price = 44999, ScreenSize = "7", Resolution = "1920x1080", Storage = "512", RAM = "16", Processor = "AMD Z1 Extreme", BatteryLife = "3-6", OS = "Windows 11", Color = "Белый", Shop = "Agroup", Description = "Игровая портативная консоль" }
            };

            foreach (var demo in demos)
            {
                ConsoleItem console = new ConsoleItem();
                console.ConsoleId = demo.Id;
                console.Brand = demo.Brand;
                console.Model = demo.Model;
                console.Price = demo.Price;
                console.Description = demo.Description;
                console.OS = demo.OS;
                console.BatteryLife = demo.BatteryLife;
                console.ScreenSize = demo.ScreenSize;
                console.Resolution = demo.Resolution;
                console.Storage = demo.Storage;
                console.RAM = demo.RAM;
                console.Processor = demo.Processor;
                console.Color = demo.Color;
                console.ShopName = demo.Shop;

                if (!imageCache.ContainsKey(demo.Id))
                {
                    Image img = LoadImageByConsoleId(demo.Id);
                    if (img == null)
                    {
                        img = CreateBrandPlaceholderImage(demo.Brand);
                    }
                    imageList1.Images.Add(img);
                    imageCache[demo.Id] = imageIndex;
                    imageIndex++;
                    console.Image = img;
                }
                else
                {
                    int cachedIndex = imageCache[demo.Id];
                    console.Image = imageList1.Images[cachedIndex];
                }

                int actualImageIndex = imageCache[demo.Id];

                string specs = "";
                if (demo.ScreenSize != "N/A")
                    specs += $"📱 {demo.ScreenSize}\" • ";
                specs += $"🎯 {demo.Resolution} • 💾 {demo.Storage} ГБ • 🧠 {demo.RAM} ГБ • ⚙ {demo.Processor}";
                if (demo.BatteryLife != "N/A")
                    specs += $" • 🔋 {demo.BatteryLife} ч";

                ListViewItem item = new ListViewItem();
                item.Text = "";
                item.ImageIndex = actualImageIndex;
                item.Tag = console;

                item.SubItems.Add($"{demo.Brand} {demo.Model}");
                item.SubItems.Add(specs);
                item.SubItems.Add($"{demo.Price:N0} ₽");
                item.SubItems.Add(demo.Shop);
                item.SubItems.Add("Купить");

                listView1.Items.Add(item);
            }

            listView1.EndUpdate();
            lblResults.Text = $"Найдено: {demos.Length} товаров (Демо)";
            this.Text = $"Каталог игровых консолей — {demos.Length} товаров (Демо)";
        }

        private void ListView1_DoubleClick(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                ListViewItem selectedItem = listView1.SelectedItems[0];
                ConsoleItem console = selectedItem.Tag as ConsoleItem;
                if (console != null)
                {
                    Form2 detailsForm = new Form2(console);
                    detailsForm.ShowDialog(this);
                }
            }
        }

        private void BtnProfile_Click(object sender, EventArgs e)
        {
            if (Session.IsLoggedIn)
            {
                // Пользователь уже авторизован - показываем профиль
                FormProfile profileForm = new FormProfile(Session.CurrentUser);
                profileForm.ShowDialog(this);
                UpdateProfileButton();
            }
            else
            {
                // Пользователь не авторизован - показываем окно входа
                FormLogin loginForm = new FormLogin();
                if (loginForm.ShowDialog() == DialogResult.OK)
                {
                    UpdateProfileButton();
                }
            }
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            listBoxBrands.ClearSelected();
            txtSearchModel.Text = "";
            numPriceFrom.Value = 0;
            numPriceTo.Value = 0;

            for (int i = 0; i < chkListStorage.Items.Count; i++)
                chkListStorage.SetItemChecked(i, false);

            for (int i = 0; i < chkListRAM.Items.Count; i++)
                chkListRAM.SetItemChecked(i, false);

            for (int i = 0; i < chkListColors.Items.Count; i++)
                chkListColors.SetItemChecked(i, false);

            for (int i = 0; i < chkListSortOptions.Items.Count; i++)
                chkListSortOptions.SetItemChecked(i, false);

            if (cmbProcessor != null) cmbProcessor.SelectedIndex = 0;
            if (cmbOS != null) cmbOS.SelectedIndex = 0;

            rbAscending.Checked = true;

            LoadData();
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            if (searchTimer != null)
            {
                searchTimer.Stop();
                searchTimer.Start();
            }
        }

        private void SearchTimer_Tick(object sender, EventArgs e)
        {
            searchTimer.Stop();
            LoadData();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (searchTimer != null)
            {
                searchTimer.Stop();
                searchTimer.Dispose();
            }
            if (myConnection != null && myConnection.State == ConnectionState.Open)
                myConnection.Close();
            if (imageList1?.Images != null)
                foreach (Image img in imageList1.Images) img?.Dispose();
        }


    }

    public class ConsoleItem
    {
        public int ConsoleId { get; set; }
        public string Brand { get; set; }
        public string Model { get; set; }
        public string Description { get; set; }
        public string OS { get; set; }
        public string Weight { get; set; }
        public string BatteryLife { get; set; }
        public string ScreenSize { get; set; }
        public string Resolution { get; set; }
        public string Storage { get; set; }
        public string RAM { get; set; }
        public string Processor { get; set; }
        public decimal Price { get; set; }
        public string Color { get; set; }
        public string ShopName { get; set; }
        public Image Image { get; set; }
    }
}