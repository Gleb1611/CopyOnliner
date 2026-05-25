using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;

namespace CopyOnliner
{
    public partial class Form1 : Form
    {
        public static string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=OnlinerConsoles.accdb;";
        public static User CurrentUser { get; set; }

        private OleDbConnection myConnection;
        private Panel searchPanel;
        private TextBox txtSearchModel;
        private Button btnProfile;
        private Button btnReset;
        private Label lblResults;
        private System.Windows.Forms.Timer searchTimer;
        private ListBox listBoxBrands;

        // Фильтр по магазинам
        private CheckedListBox chkListShops;

        private Label lblBrand;
        private Label lblModel;
        private Label lblPrice;
        private Label lblScreenSize;
        private Label lblResolution;
        private Label lblStorage;
        private Label lblRAM;
        private Label lblProcessor;
        private Label lblOS;
        private Label lblColor;
        private Label lblShops;

        private NumericUpDown numPriceFrom;
        private NumericUpDown numPriceTo;
        private NumericUpDown numScreenSizeFrom;
        private NumericUpDown numScreenSizeTo;
        private TextBox txtResolution;
        private CheckedListBox chkListStorage;
        private CheckedListBox chkListRAM;
        private CheckedListBox chkListColors;
        private ComboBox cmbProcessor;
        private ComboBox cmbOS;
        private CheckedListBox chkListSortOptions;
        private RadioButton rbAscending;
        private RadioButton rbDescending;

        private Button btnApply;

        // Геймерская цветовая палитра
        private Color darkBg = Color.FromArgb(15, 25, 35);
        private Color panelBg = Color.FromArgb(25, 35, 45);
        private Color accentColor = Color.FromArgb(0, 255, 100);
        private Color accentBlue = Color.FromArgb(0, 150, 255);
        private Color textColor = Color.FromArgb(220, 220, 220);
        private Color cardBg = Color.FromArgb(35, 45, 55);
        private Color selectedBg = Color.FromArgb(0, 100, 50);

        private const int CardWidth = 340;
        private const int CardHeight = 290;

        public Form1()
        {
            InitializeComponent();

            // Включаем двойную буферизацию для устранения мерцания
            this.SetStyle(ControlStyles.DoubleBuffer |
                          ControlStyles.UserPaint |
                          ControlStyles.AllPaintingInWmPaint |
                          ControlStyles.ResizeRedraw, true);
            this.UpdateStyles();

            this.BackColor = darkBg;
            this.WindowState = FormWindowState.Maximized;
            this.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            searchTimer = new System.Windows.Forms.Timer();
            searchTimer.Interval = 500;
            searchTimer.Tick += SearchTimer_Tick;

            SetupListView();
            InitializeSearchPanel();

            UpdateProfileButton();

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

        private void UpdateProfileButton()
        {
            if (btnProfile != null)
            {
                if (CurrentUser != null && CurrentUser.IsLoggedIn)
                {
                    btnProfile.Text = $"👤 {CurrentUser.Username}";
                    btnProfile.BackColor = accentColor;
                }
                else
                {
                    btnProfile.Text = "👤 ВОЙТИ";
                    btnProfile.BackColor = accentBlue;
                }
            }
        }

        private void LoadFilters()
        {
            try
            {
                // Загружаем бренды
                string brandQuery = "SELECT DISTINCT Brand FROM Consoles WHERE Brand IS NOT NULL ORDER BY Brand";
                OleDbCommand cmd = new OleDbCommand(brandQuery, myConnection);
                OleDbDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    listBoxBrands.Items.Add(reader["Brand"].ToString());
                }
                reader.Close();

                // Загружаем магазины
                string shopQuery = "SELECT ShopID, ShopName FROM Shops ORDER BY ShopName";
                cmd = new OleDbCommand(shopQuery, myConnection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string shopName = reader["ShopName"].ToString();
                    if (!string.IsNullOrEmpty(shopName))
                        chkListShops.Items.Add(shopName);
                }
                reader.Close();

                // Загружаем память
                string storageQuery = "SELECT DISTINCT Storage FROM Consoles WHERE Storage IS NOT NULL ORDER BY Storage";
                cmd = new OleDbCommand(storageQuery, myConnection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string storage = reader["Storage"].ToString();
                    if (!string.IsNullOrEmpty(storage))
                        chkListStorage.Items.Add(storage);
                }
                reader.Close();

                // Загружаем ОЗУ
                string ramQuery = "SELECT DISTINCT RAM FROM Consoles WHERE RAM IS NOT NULL ORDER BY RAM";
                cmd = new OleDbCommand(ramQuery, myConnection);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string ram = reader["RAM"].ToString();
                    if (!string.IsNullOrEmpty(ram))
                        chkListRAM.Items.Add(ram);
                }
                reader.Close();

                // Загружаем процессоры
                string processorQuery = "SELECT DISTINCT Processor FROM Consoles WHERE Processor IS NOT NULL ORDER BY Processor";
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

                // Загружаем ОС
                string osQuery = "SELECT DISTINCT OS FROM Consoles WHERE OS IS NOT NULL ORDER BY OS";
                cmd = new OleDbCommand(osQuery, myConnection);
                reader = cmd.ExecuteReader();
                cmbOS.Items.Add("Все");
                while (reader.Read())
                {
                    string os = reader["OS"].ToString();
                    if (!string.IsNullOrEmpty(os))
                        cmbOS.Items.Add(os);
                }
                reader.Close();
                cmbOS.SelectedIndex = 0;

                // Загружаем цвета
                string colorQuery = "SELECT DISTINCT Color FROM Consoles WHERE Color IS NOT NULL AND Color <> '' ORDER BY Color";
                cmd = new OleDbCommand(colorQuery, myConnection);
                reader = cmd.ExecuteReader();
                chkListColors.Items.Clear();
                while (reader.Read())
                {
                    string color = reader["Color"].ToString();
                    if (!string.IsNullOrEmpty(color))
                        chkListColors.Items.Add(color);
                }
                reader.Close();
            }
            catch
            {
                // Демо-данные для фильтров
                listBoxBrands.Items.AddRange(new string[] { "Sony", "Microsoft", "Nintendo", "Valve", "ASUS" });
                chkListShops.Items.AddRange(new string[] { "Agroup", "GAMEPARK", "Newton" });
                chkListStorage.Items.AddRange(new string[] { "64", "128", "256", "512", "825", "1024" });
                chkListRAM.Items.AddRange(new string[] { "4", "8", "16", "32" });
                cmbProcessor.Items.AddRange(new string[] { "Все", "AMD Zen 2", "AMD APU", "AMD Z1 Extreme", "NVIDIA Tegra" });
                cmbOS.Items.AddRange(new string[] { "Все", "PlayStation OS", "Xbox OS", "Nintendo OS", "SteamOS", "Windows 11" });
                chkListColors.Items.AddRange(new string[] { "Черный", "Белый", "Серый", "Красный", "Синий", "Фиолетовый" });
                cmbProcessor.SelectedIndex = 0;
                cmbOS.SelectedIndex = 0;
            }
        }

        private void InitializeSearchPanel()
        {
            searchPanel = new Panel();
            searchPanel.Dock = DockStyle.Left;
            searchPanel.Width = 420;
            searchPanel.BackColor = panelBg;
            searchPanel.Padding = new Padding(15);
            searchPanel.ForeColor = textColor;
            searchPanel.AutoScroll = true;

            int yOffset = 10;

            // Кнопка профиля
            btnProfile = new Button();
            btnProfile.Text = "👤 ВОЙТИ";
            btnProfile.Location = new Point(290, 10);
            btnProfile.Size = new Size(110, 35);
            btnProfile.BackColor = accentBlue;
            btnProfile.ForeColor = Color.Black;
            btnProfile.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            btnProfile.FlatStyle = FlatStyle.Flat;
            btnProfile.Cursor = Cursors.Hand;
            btnProfile.Click += BtnProfile_Click;
            searchPanel.Controls.Add(btnProfile);

            Label lblTitle = new Label();
            lblTitle.Text = "⚡ ИГРОВОЙ КАТАЛОГ";
            lblTitle.Location = new Point(15, 10);
            lblTitle.Size = new Size(260, 35);
            lblTitle.Font = new Font("Segoe UI", 14, FontStyle.Bold);
            lblTitle.ForeColor = accentColor;
            lblTitle.TextAlign = ContentAlignment.MiddleCenter;
            searchPanel.Controls.Add(lblTitle);
            yOffset = 55;

            Panel separator = new Panel();
            separator.BackColor = accentColor;
            separator.Location = new Point(15, yOffset);
            separator.Size = new Size(390, 2);
            searchPanel.Controls.Add(separator);
            yOffset += 15;

            // Бренды
            lblBrand = new Label();
            lblBrand.Text = "🔍 БРЕНД (выберите несколько)";
            lblBrand.Location = new Point(15, yOffset);
            lblBrand.Size = new Size(390, 25);
            lblBrand.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblBrand.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblBrand);
            yOffset += 28;

            listBoxBrands = new ListBox();
            listBoxBrands.Location = new Point(15, yOffset);
            listBoxBrands.Size = new Size(390, 80);
            listBoxBrands.BackColor = darkBg;
            listBoxBrands.ForeColor = textColor;
            listBoxBrands.BorderStyle = BorderStyle.FixedSingle;
            listBoxBrands.SelectionMode = SelectionMode.MultiExtended;
            listBoxBrands.Font = new Font("Segoe UI", 10);
            searchPanel.Controls.Add(listBoxBrands);
            yOffset += 90;

            // Модель
            lblModel = new Label();
            lblModel.Text = "🎮 МОДЕЛЬ (поиск по названию)";
            lblModel.Location = new Point(15, yOffset);
            lblModel.Size = new Size(390, 25);
            lblModel.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblModel.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblModel);
            yOffset += 28;

            txtSearchModel = new TextBox();
            txtSearchModel.Location = new Point(15, yOffset);
            txtSearchModel.Size = new Size(390, 30);
            txtSearchModel.Font = new Font("Segoe UI", 11);
            txtSearchModel.BackColor = darkBg;
            txtSearchModel.ForeColor = textColor;
            txtSearchModel.BorderStyle = BorderStyle.FixedSingle;
            txtSearchModel.TextChanged += TxtSearch_TextChanged;
            searchPanel.Controls.Add(txtSearchModel);
            yOffset += 40;

            // Цена
            lblPrice = new Label();
            lblPrice.Text = "💰 ЦЕНА (руб.)";
            lblPrice.Location = new Point(15, yOffset);
            lblPrice.Size = new Size(390, 25);
            lblPrice.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblPrice.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblPrice);
            yOffset += 28;

            Label lblFrom = new Label();
            lblFrom.Text = "от";
            lblFrom.Location = new Point(15, yOffset);
            lblFrom.Size = new Size(30, 25);
            lblFrom.ForeColor = textColor;
            searchPanel.Controls.Add(lblFrom);

            numPriceFrom = new NumericUpDown();
            numPriceFrom.Location = new Point(45, yOffset);
            numPriceFrom.Size = new Size(150, 27);
            numPriceFrom.BackColor = darkBg;
            numPriceFrom.ForeColor = textColor;
            numPriceFrom.Maximum = 1000000;
            numPriceFrom.ThousandsSeparator = true;
            searchPanel.Controls.Add(numPriceFrom);

            Label lblTo = new Label();
            lblTo.Text = "до";
            lblTo.Location = new Point(205, yOffset);
            lblTo.Size = new Size(30, 25);
            lblTo.ForeColor = textColor;
            searchPanel.Controls.Add(lblTo);

            numPriceTo = new NumericUpDown();
            numPriceTo.Location = new Point(235, yOffset);
            numPriceTo.Size = new Size(170, 27);
            numPriceTo.BackColor = darkBg;
            numPriceTo.ForeColor = textColor;
            numPriceTo.Maximum = 1000000;
            numPriceTo.ThousandsSeparator = true;
            searchPanel.Controls.Add(numPriceTo);
            yOffset += 40;

            // Размер экрана
            lblScreenSize = new Label();
            lblScreenSize.Text = "📱 РАЗМЕР ЭКРАНА (дюймы)";
            lblScreenSize.Location = new Point(15, yOffset);
            lblScreenSize.Size = new Size(390, 25);
            lblScreenSize.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblScreenSize.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblScreenSize);
            yOffset += 28;

            Label lblScreenFrom = new Label();
            lblScreenFrom.Text = "от";
            lblScreenFrom.Location = new Point(15, yOffset);
            lblScreenFrom.Size = new Size(30, 25);
            lblScreenFrom.ForeColor = textColor;
            searchPanel.Controls.Add(lblScreenFrom);

            numScreenSizeFrom = new NumericUpDown();
            numScreenSizeFrom.Location = new Point(45, yOffset);
            numScreenSizeFrom.Size = new Size(150, 27);
            numScreenSizeFrom.BackColor = darkBg;
            numScreenSizeFrom.ForeColor = textColor;
            numScreenSizeFrom.Maximum = 100;
            searchPanel.Controls.Add(numScreenSizeFrom);

            Label lblScreenTo = new Label();
            lblScreenTo.Text = "до";
            lblScreenTo.Location = new Point(205, yOffset);
            lblScreenTo.Size = new Size(30, 25);
            lblScreenTo.ForeColor = textColor;
            searchPanel.Controls.Add(lblScreenTo);

            numScreenSizeTo = new NumericUpDown();
            numScreenSizeTo.Location = new Point(235, yOffset);
            numScreenSizeTo.Size = new Size(170, 27);
            numScreenSizeTo.BackColor = darkBg;
            numScreenSizeTo.ForeColor = textColor;
            numScreenSizeTo.Maximum = 100;
            searchPanel.Controls.Add(numScreenSizeTo);
            yOffset += 40;

            // Разрешение
            lblResolution = new Label();
            lblResolution.Text = "🎯 РАЗРЕШЕНИЕ (поиск по тексту)";
            lblResolution.Location = new Point(15, yOffset);
            lblResolution.Size = new Size(390, 25);
            lblResolution.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblResolution.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblResolution);
            yOffset += 28;

            txtResolution = new TextBox();
            txtResolution.Location = new Point(15, yOffset);
            txtResolution.Size = new Size(390, 30);
            txtResolution.Font = new Font("Segoe UI", 11);
            txtResolution.BackColor = darkBg;
            txtResolution.ForeColor = textColor;
            txtResolution.BorderStyle = BorderStyle.FixedSingle;
            txtResolution.PlaceholderText = "например: 1920x1080, 4K, 1280x720";
            searchPanel.Controls.Add(txtResolution);
            yOffset += 40;

            // Память
            lblStorage = new Label();
            lblStorage.Text = "💾 ОБЪЁМ ПАМЯТИ (ГБ)";
            lblStorage.Location = new Point(15, yOffset);
            lblStorage.Size = new Size(390, 25);
            lblStorage.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblStorage.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblStorage);
            yOffset += 28;

            chkListStorage = new CheckedListBox();
            chkListStorage.Location = new Point(15, yOffset);
            chkListStorage.Size = new Size(390, 60);
            chkListStorage.BackColor = darkBg;
            chkListStorage.ForeColor = textColor;
            chkListStorage.BorderStyle = BorderStyle.FixedSingle;
            chkListStorage.CheckOnClick = true;
            searchPanel.Controls.Add(chkListStorage);
            yOffset += 70;

            // ОЗУ
            lblRAM = new Label();
            lblRAM.Text = "🧠 ОЗУ (ГБ)";
            lblRAM.Location = new Point(15, yOffset);
            lblRAM.Size = new Size(390, 25);
            lblRAM.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblRAM.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblRAM);
            yOffset += 28;

            chkListRAM = new CheckedListBox();
            chkListRAM.Location = new Point(15, yOffset);
            chkListRAM.Size = new Size(390, 60);
            chkListRAM.BackColor = darkBg;
            chkListRAM.ForeColor = textColor;
            chkListRAM.BorderStyle = BorderStyle.FixedSingle;
            chkListRAM.CheckOnClick = true;
            searchPanel.Controls.Add(chkListRAM);
            yOffset += 70;

            // Процессор
            lblProcessor = new Label();
            lblProcessor.Text = "⚙ ПРОЦЕССОР";
            lblProcessor.Location = new Point(15, yOffset);
            lblProcessor.Size = new Size(390, 25);
            lblProcessor.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblProcessor.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblProcessor);
            yOffset += 28;

            cmbProcessor = new ComboBox();
            cmbProcessor.Location = new Point(15, yOffset);
            cmbProcessor.Size = new Size(390, 30);
            cmbProcessor.Font = new Font("Segoe UI", 11);
            cmbProcessor.BackColor = darkBg;
            cmbProcessor.ForeColor = textColor;
            cmbProcessor.DropDownStyle = ComboBoxStyle.DropDownList;
            searchPanel.Controls.Add(cmbProcessor);
            yOffset += 40;

            // ОС
            lblOS = new Label();
            lblOS.Text = "💿 ОПЕРАЦИОННАЯ СИСТЕМА";
            lblOS.Location = new Point(15, yOffset);
            lblOS.Size = new Size(390, 25);
            lblOS.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblOS.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblOS);
            yOffset += 28;

            cmbOS = new ComboBox();
            cmbOS.Location = new Point(15, yOffset);
            cmbOS.Size = new Size(390, 30);
            cmbOS.Font = new Font("Segoe UI", 11);
            cmbOS.BackColor = darkBg;
            cmbOS.ForeColor = textColor;
            cmbOS.DropDownStyle = ComboBoxStyle.DropDownList;
            searchPanel.Controls.Add(cmbOS);
            yOffset += 40;

            // Цвет
            lblColor = new Label();
            lblColor.Text = "🎨 ЦВЕТ";
            lblColor.Location = new Point(15, yOffset);
            lblColor.Size = new Size(390, 25);
            lblColor.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblColor.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblColor);
            yOffset += 28;

            chkListColors = new CheckedListBox();
            chkListColors.Location = new Point(15, yOffset);
            chkListColors.Size = new Size(390, 80);
            chkListColors.BackColor = darkBg;
            chkListColors.ForeColor = textColor;
            chkListColors.BorderStyle = BorderStyle.FixedSingle;
            chkListColors.CheckOnClick = true;
            searchPanel.Controls.Add(chkListColors);
            yOffset += 90;

            // МАГАЗИНЫ
            lblShops = new Label();
            lblShops.Text = "🏪 МАГАЗИНЫ (выберите несколько)";
            lblShops.Location = new Point(15, yOffset);
            lblShops.Size = new Size(390, 25);
            lblShops.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblShops.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblShops);
            yOffset += 28;

            chkListShops = new CheckedListBox();
            chkListShops.Location = new Point(15, yOffset);
            chkListShops.Size = new Size(390, 80);
            chkListShops.BackColor = darkBg;
            chkListShops.ForeColor = textColor;
            chkListShops.BorderStyle = BorderStyle.FixedSingle;
            chkListShops.CheckOnClick = true;
            searchPanel.Controls.Add(chkListShops);
            yOffset += 90;

            // Сортировка
            Label lblSort = new Label();
            lblSort.Text = "📊 СОРТИРОВКА (выберите несколько)";
            lblSort.Location = new Point(15, yOffset);
            lblSort.Size = new Size(390, 25);
            lblSort.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblSort.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblSort);
            yOffset += 28;

            chkListSortOptions = new CheckedListBox();
            chkListSortOptions.Location = new Point(15, yOffset);
            chkListSortOptions.Size = new Size(390, 110);
            chkListSortOptions.BackColor = darkBg;
            chkListSortOptions.ForeColor = textColor;
            chkListSortOptions.BorderStyle = BorderStyle.FixedSingle;
            chkListSortOptions.CheckOnClick = true;
            chkListSortOptions.Items.AddRange(new string[] {
                "Бренд", "Модель", "Цена", "Размер экрана", "Разрешение",
                "Объём памяти", "ОЗУ", "Процессор", "Батарея", "ОС", "Цвет"
            });
            searchPanel.Controls.Add(chkListSortOptions);
            yOffset += 120;

            // Направление сортировки
            rbAscending = new RadioButton();
            rbAscending.Text = "⬆ По возрастанию";
            rbAscending.Location = new Point(15, yOffset);
            rbAscending.Size = new Size(120, 25);
            rbAscending.ForeColor = textColor;
            rbAscending.Checked = true;
            searchPanel.Controls.Add(rbAscending);

            rbDescending = new RadioButton();
            rbDescending.Text = "⬇ По убыванию";
            rbDescending.Location = new Point(140, yOffset);
            rbDescending.Size = new Size(120, 25);
            rbDescending.ForeColor = textColor;
            searchPanel.Controls.Add(rbDescending);
            yOffset += 35;

            // Кнопка применения
            btnApply = new Button();
            btnApply.Text = "🔍 ПРИМЕНИТЬ ФИЛЬТРЫ И СОРТИРОВКУ";
            btnApply.Location = new Point(15, yOffset);
            btnApply.Size = new Size(390, 50);
            btnApply.BackColor = accentColor;
            btnApply.ForeColor = Color.Black;
            btnApply.Font = new Font("Segoe UI", 11, FontStyle.Bold);
            btnApply.FlatStyle = FlatStyle.Flat;
            btnApply.Cursor = Cursors.Hand;
            btnApply.Click += BtnApply_Click;
            searchPanel.Controls.Add(btnApply);
            yOffset += 60;

            // Кнопка сброса
            btnReset = new Button();
            btnReset.Text = "🔄 СБРОСИТЬ ВСЕ ФИЛЬТРЫ";
            btnReset.Location = new Point(15, yOffset);
            btnReset.Size = new Size(390, 45);
            btnReset.BackColor = Color.FromArgb(60, 60, 70);
            btnReset.ForeColor = textColor;
            btnReset.Font = new Font("Segoe UI", 11, FontStyle.Bold);
            btnReset.FlatStyle = FlatStyle.Flat;
            btnReset.Cursor = Cursors.Hand;
            btnReset.Click += BtnReset_Click;
            searchPanel.Controls.Add(btnReset);
            yOffset += 55;

            lblResults = new Label();
            lblResults.Text = "📦 Найдено: 0";
            lblResults.Location = new Point(15, yOffset);
            lblResults.Size = new Size(390, 30);
            lblResults.Font = new Font("Segoe UI", 11, FontStyle.Bold);
            lblResults.ForeColor = accentColor;
            lblResults.TextAlign = ContentAlignment.MiddleCenter;
            searchPanel.Controls.Add(lblResults);

            this.Controls.Add(searchPanel);

            listView1.BackColor = darkBg;
            listView1.ForeColor = textColor;
            listView1.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            listView1.BorderStyle = BorderStyle.None;
            listView1.Left = searchPanel.Width;
            listView1.Top = 0;
            listView1.Width = this.ClientSize.Width - searchPanel.Width;
            listView1.Height = this.ClientSize.Height;
            listView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        }

        private void SetupListView()
        {
            imageList1.ImageSize = new Size(320, 200);
            imageList1.ColorDepth = ColorDepth.Depth32Bit;

            listView1.LargeImageList = imageList1;
            listView1.SmallImageList = imageList1;
            listView1.View = View.LargeIcon;
            listView1.LabelWrap = true;
            listView1.OwnerDraw = true;
            listView1.DoubleClick += ListView1_DoubleClick;
            listView1.DrawItem += ListView1_DrawItem;
            listView1.TileSize = new Size(CardWidth, CardHeight);
            listView1.BackColor = darkBg;
            listView1.Dock = DockStyle.Fill;
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
                    // НЕ вызываем LoadData() - карточки не перезагружаются
                }
            }
        }

        private void BtnProfile_Click(object sender, EventArgs e)
        {
            if (CurrentUser != null && CurrentUser.IsLoggedIn)
            {
                FormProfile profileForm = new FormProfile(CurrentUser);
                profileForm.ShowDialog(this);
            }
            else
            {
                FormLogin loginForm = new FormLogin();
                if (loginForm.ShowDialog() == DialogResult.OK)
                {
                    CurrentUser = loginForm.CurrentUser;
                    UpdateProfileButton();

                    if (CurrentUser != null && CurrentUser.IsLoggedIn)
                    {
                        FormProfile profileForm = new FormProfile(CurrentUser);
                        profileForm.ShowDialog(this);
                    }
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
            numScreenSizeFrom.Value = 0;
            numScreenSizeTo.Value = 0;
            txtResolution.Text = "";

            for (int i = 0; i < chkListStorage.Items.Count; i++)
                chkListStorage.SetItemChecked(i, false);

            for (int i = 0; i < chkListRAM.Items.Count; i++)
                chkListRAM.SetItemChecked(i, false);

            for (int i = 0; i < chkListColors.Items.Count; i++)
                chkListColors.SetItemChecked(i, false);

            for (int i = 0; i < chkListShops.Items.Count; i++)
                chkListShops.SetItemChecked(i, false);

            for (int i = 0; i < chkListSortOptions.Items.Count; i++)
                chkListSortOptions.SetItemChecked(i, false);

            cmbProcessor.SelectedIndex = 0;
            cmbOS.SelectedIndex = 0;
            rbAscending.Checked = true;

            LoadData();
        }

        private void LoadData()
        {
            try
            {
                listView1.BeginUpdate();
                listView1.Items.Clear();
                imageList1.Images.Clear();

                string query = @"
                    SELECT DISTINCT
                        c.ConsoleID, c.Brand, c.Model, c.Description, c.ImageURL, 
                        c.OS, c.Weight, c.BatteryLife, c.ScreenSize, c.Resolution,
                        c.Storage, c.RAM, c.Processor, c.Price, c.Color,
                        s.ShopName
                    FROM (Consoles c 
                    LEFT JOIN AvailableInShop a ON c.ConsoleID = a.ConsoleID)
                    LEFT JOIN Shops s ON a.ShopID = s.ShopID
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
                    query += $" AND c.Brand IN ({brands})";
                }

                if (!string.IsNullOrWhiteSpace(txtSearchModel.Text))
                {
                    query += $" AND c.Model LIKE '%{txtSearchModel.Text.Replace("'", "''")}%'";
                }

                if (numPriceFrom.Value > 0)
                {
                    query += $" AND c.Price >= {numPriceFrom.Value}";
                }
                if (numPriceTo.Value > 0)
                {
                    query += $" AND c.Price <= {numPriceTo.Value}";
                }

                if (numScreenSizeFrom.Value > 0)
                {
                    query += $" AND c.ScreenSize >= {numScreenSizeFrom.Value}";
                }
                if (numScreenSizeTo.Value > 0)
                {
                    query += $" AND c.ScreenSize <= {numScreenSizeTo.Value}";
                }

                if (!string.IsNullOrWhiteSpace(txtResolution.Text))
                {
                    query += $" AND c.Resolution LIKE '%{txtResolution.Text.Replace("'", "''")}%'";
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
                    query += $" AND c.Storage IN ({storageList})";
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
                    query += $" AND c.RAM IN ({ramList})";
                }

                if (cmbProcessor.SelectedIndex > 0 && cmbProcessor.SelectedItem.ToString() != "Все")
                {
                    query += $" AND c.Processor LIKE '%{cmbProcessor.SelectedItem.ToString().Replace("'", "''")}%'";
                }

                if (cmbOS.SelectedIndex > 0 && cmbOS.SelectedItem.ToString() != "Все")
                {
                    query += $" AND c.OS LIKE '%{cmbOS.SelectedItem.ToString().Replace("'", "''")}%'";
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
                    query += $" AND c.Color IN ({colorList})";
                }

                // ФИЛЬТР ПО МАГАЗИНАМ
                var selectedShops = new List<string>();
                foreach (var item in chkListShops.CheckedItems)
                {
                    selectedShops.Add(item.ToString());
                }
                if (selectedShops.Count > 0)
                {
                    string shopList = "";
                    foreach (string shop in selectedShops)
                    {
                        shopList += $"'{shop.Replace("'", "''")}',";
                    }
                    shopList = shopList.TrimEnd(',');
                    query += $" AND s.ShopName IN ({shopList})";
                }

                string orderBy = BuildOrderByClause();
                if (!string.IsNullOrEmpty(orderBy))
                {
                    query += $" ORDER BY {orderBy}";
                }
                else
                {
                    query += " ORDER BY c.ConsoleID";
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
                    console.ImageURL = reader["ImageURL"]?.ToString() ?? "";
                    console.OS = reader["OS"] != DBNull.Value ? reader["OS"].ToString() : "";
                    console.BatteryLife = reader["BatteryLife"] != DBNull.Value ? reader["BatteryLife"].ToString() : "";
                    console.ScreenSize = reader["ScreenSize"] != DBNull.Value ? reader["ScreenSize"].ToString() : "";
                    console.Resolution = reader["Resolution"] != DBNull.Value ? reader["Resolution"].ToString() : "";
                    console.Storage = reader["Storage"] != DBNull.Value ? reader["Storage"].ToString() : "";
                    console.RAM = reader["RAM"] != DBNull.Value ? reader["RAM"].ToString() : "";
                    console.Processor = reader["Processor"] != DBNull.Value ? reader["Processor"].ToString() : "";
                    console.Price = reader["Price"] != DBNull.Value ? Convert.ToDecimal(reader["Price"]) : 0;
                    console.Color = reader["Color"] != DBNull.Value ? reader["Color"].ToString() : "";
                    console.ShopName = reader["ShopName"] != DBNull.Value ? reader["ShopName"].ToString() : "Не указан";

                    try
                    {
                        if (!string.IsNullOrEmpty(console.ImageURL) && (console.ImageURL.StartsWith("http") || console.ImageURL.StartsWith("https")))
                        {
                            using (WebClient wc = new WebClient())
                            {
                                byte[] bytes = wc.DownloadData(console.ImageURL);
                                using (MemoryStream ms = new MemoryStream(bytes))
                                {
                                    using (Image image = Image.FromStream(ms))
                                    {
                                        console.Image = new Bitmap(image, imageList1.ImageSize);
                                        imageList1.Images.Add(console.Image);
                                    }
                                }
                            }
                        }
                        else
                        {
                            console.Image = CreateBrandPlaceholderImage(console.Brand);
                            imageList1.Images.Add(console.Image);
                        }
                    }
                    catch
                    {
                        console.Image = CreateBrandPlaceholderImage(console.Brand);
                        imageList1.Images.Add(console.Image);
                    }

                    string priceText = console.Price > 0 ? $"💰 {console.Price:N0} ₽" : "💰 Цена не указана";
                    string colorText = !string.IsNullOrEmpty(console.Color) ? $"🎨 {console.Color}" : "";
                    string shopText = !string.IsNullOrEmpty(console.ShopName) && console.ShopName != "Не указан" ? $"🏪 {console.ShopName}" : "";
                    string specs = $"📱 {console.ScreenSize}\" | 🎯 {console.Resolution} | 💾 {console.Storage}GB | 🧠 {console.RAM}GB | {colorText} | {shopText}";

                    ListViewItem item = new ListViewItem();
                    item.Text = $"{console.Brand} {console.Model}\n{priceText}\n{specs}";
                    item.ImageIndex = imageIndex;
                    item.Tag = console;

                    listView1.Items.Add(item);
                    imageIndex++;
                }

                reader.Close();
                listView1.EndUpdate();
                listView1.Refresh();
                lblResults.Text = $"📦 Найдено: {imageIndex}";
                this.Text = $"GAMER CATALOG - {imageIndex} консолей";
            }
            catch (Exception ex)
            {
                listView1.EndUpdate();
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AddDemoData();
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
                        orderByFields.Add($"c.Brand {sortOrder}");
                        break;
                    case "Модель":
                        orderByFields.Add($"c.Model {sortOrder}");
                        break;
                    case "Цена":
                        orderByFields.Add($"c.Price {sortOrder}");
                        break;
                    case "Размер экрана":
                        orderByFields.Add($"c.ScreenSize {sortOrder}");
                        break;
                    case "Разрешение":
                        orderByFields.Add($"c.Resolution {sortOrder}");
                        break;
                    case "Объём памяти":
                        orderByFields.Add($"c.Storage {sortOrder}");
                        break;
                    case "ОЗУ":
                        orderByFields.Add($"c.RAM {sortOrder}");
                        break;
                    case "Процессор":
                        orderByFields.Add($"c.Processor {sortOrder}");
                        break;
                    case "Батарея":
                        orderByFields.Add($"c.BatteryLife {sortOrder}");
                        break;
                    case "ОС":
                        orderByFields.Add($"c.OS {sortOrder}");
                        break;
                    case "Цвет":
                        orderByFields.Add($"c.Color {sortOrder}");
                        break;
                }
            }

            return orderByFields.Count > 0 ? string.Join(", ", orderByFields) : "";
        }

        private Image CreateBrandPlaceholderImage(string brand)
        {
            Bitmap placeholder = new Bitmap(imageList1.ImageSize.Width, imageList1.ImageSize.Height);
            using (Graphics g = Graphics.FromImage(placeholder))
            {
                g.Clear(cardBg);
                using (LinearGradientBrush gradient = new LinearGradientBrush(
                    new Rectangle(0, 0, placeholder.Width, placeholder.Height),
                    Color.FromArgb(50, accentColor),
                    Color.FromArgb(50, accentBlue), 45f))
                {
                    g.FillRectangle(gradient, 0, 0, placeholder.Width, placeholder.Height);
                }
                using (Font font = new Font("Segoe UI", 48, FontStyle.Bold))
                {
                    string text = brand.Length > 0 ? brand[0].ToString() : "?";
                    SizeF textSize = g.MeasureString(text, font);
                    g.DrawString(text, font, new SolidBrush(accentColor),
                        (placeholder.Width - textSize.Width) / 2,
                        (placeholder.Height - textSize.Height) / 2);
                }
            }
            return placeholder;
        }

        private void AddDemoData()
        {
            listView1.BeginUpdate();
            listView1.Items.Clear();
            imageList1.Images.Clear();

            var demos = new[]
            {
                new { Brand = "Sony", Model = "PlayStation 5", Price = 49999, ScreenSize = "N/A", Resolution = "4K", Storage = "825", RAM = "16", Processor = "AMD Zen 2", BatteryLife = "N/A", OS = "PlayStation OS", Color = "Белый", Shop = "GAMEPARK", Description = "Новейшая консоль Sony" },
                new { Brand = "Microsoft", Model = "Xbox Series X", Price = 45999, ScreenSize = "N/A", Resolution = "4K", Storage = "1024", RAM = "16", Processor = "AMD Zen 2", BatteryLife = "N/A", OS = "Xbox OS", Color = "Черный", Shop = "Agroup", Description = "Самая мощная консоль Xbox" },
                new { Brand = "Nintendo", Model = "Switch OLED", Price = 29999, ScreenSize = "7", Resolution = "1280x720", Storage = "64", RAM = "4", Processor = "NVIDIA Tegra", BatteryLife = "4-9", OS = "Nintendo OS", Color = "Красный", Shop = "Newton", Description = "Гибридная консоль" },
                new { Brand = "Valve", Model = "Steam Deck", Price = 39999, ScreenSize = "7", Resolution = "1280x800", Storage = "512", RAM = "16", Processor = "AMD APU", BatteryLife = "2-8", OS = "SteamOS", Color = "Черный", Shop = "GAMEPARK", Description = "Портативный компьютер" },
                new { Brand = "ASUS", Model = "ROG Ally", Price = 44999, ScreenSize = "7", Resolution = "1920x1080", Storage = "512", RAM = "16", Processor = "AMD Z1 Extreme", BatteryLife = "3-6", OS = "Windows 11", Color = "Белый", Shop = "Agroup", Description = "Игровая портативная консоль" }
            };

            int imageIndex = 0;
            foreach (var demo in demos)
            {
                ConsoleItem console = new ConsoleItem();
                console.ConsoleId = imageIndex + 1;
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
                console.Image = CreateBrandPlaceholderImage(demo.Brand);

                imageList1.Images.Add(console.Image);

                string priceText = console.Price > 0 ? $"💰 {console.Price:N0} ₽" : "💰 Цена не указана";
                string colorText = !string.IsNullOrEmpty(console.Color) ? $"🎨 {console.Color}" : "";
                string shopText = !string.IsNullOrEmpty(console.ShopName) ? $"🏪 {console.ShopName}" : "";
                string specs = $"📱 {console.ScreenSize}\" | 🎯 {console.Resolution} | 💾 {console.Storage}GB | 🧠 {console.RAM}GB | {colorText} | {shopText}";

                ListViewItem item = new ListViewItem();
                item.Text = $"{demo.Brand} {demo.Model}\n{priceText}\n{specs}";
                item.ImageIndex = imageIndex;
                item.Tag = console;

                listView1.Items.Add(item);
                imageIndex++;
            }

            listView1.EndUpdate();
            listView1.Refresh();
            lblResults.Text = $"📦 Найдено: {demos.Length} (Демо)";
            this.Text = $"GAMER CATALOG - {demos.Length} демо-консолей";
        }

        private void ListView1_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
            Rectangle bounds = e.Bounds;

            using (SolidBrush bgBrush = new SolidBrush(e.Item.Selected ? selectedBg : cardBg))
            {
                e.Graphics.FillRectangle(bgBrush, bounds);
            }

            using (Pen borderPen = new Pen(e.Item.Selected ? accentColor : Color.FromArgb(60, 70, 80), 2))
            {
                e.Graphics.DrawRectangle(borderPen, bounds);
            }

            if (e.Item.ImageIndex != -1 && imageList1.Images.Count > e.Item.ImageIndex)
            {
                Image img = imageList1.Images[e.Item.ImageIndex];
                Rectangle imageRect = new Rectangle(bounds.X + 10, bounds.Y + 10, bounds.Width - 20, 160);
                e.Graphics.DrawImage(img, imageRect);
            }

            string[] textLines = e.Item.Text.Split('\n');
            using (Font titleFont = new Font("Segoe UI", 11, FontStyle.Bold))
            using (Font priceFont = new Font("Segoe UI", 10, FontStyle.Bold))
            using (Font specsFont = new Font("Segoe UI", 8))
            using (SolidBrush titleBrush = new SolidBrush(textColor))
            using (SolidBrush priceBrush = new SolidBrush(accentColor))
            using (SolidBrush specsBrush = new SolidBrush(Color.FromArgb(150, 150, 150)))
            {
                if (textLines.Length > 0)
                {
                    e.Graphics.DrawString(textLines[0], titleFont, titleBrush,
                        new RectangleF(bounds.X + 10, bounds.Y + 180, bounds.Width - 20, 30));
                }
                if (textLines.Length > 1)
                {
                    e.Graphics.DrawString(textLines[1], priceFont, priceBrush,
                        new RectangleF(bounds.X + 10, bounds.Y + 205, bounds.Width - 20, 30));
                }
                if (textLines.Length > 2)
                {
                    e.Graphics.DrawString(textLines[2], specsFont, specsBrush,
                        new RectangleF(bounds.X + 10, bounds.Y + 230, bounds.Width - 20, 60));
                }
            }

            e.DrawDefault = false;
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

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (listView1 != null && searchPanel != null)
            {
                listView1.Left = searchPanel.Width;
                listView1.Width = this.ClientSize.Width - searchPanel.Width;
                listView1.Height = this.ClientSize.Height;
            }
        }
    }

    public class ConsoleItem
    {
        public int ConsoleId { get; set; }
        public string Brand { get; set; }
        public string Model { get; set; }
        public string Description { get; set; }
        public string ImageURL { get; set; }
        public decimal Price { get; set; }
        public string ShopName { get; set; }
        public string OS { get; set; }
        public string Weight { get; set; }
        public string BatteryLife { get; set; }
        public string ScreenSize { get; set; }
        public string Resolution { get; set; }
        public string Storage { get; set; }
        public string RAM { get; set; }
        public string Processor { get; set; }
        public string Color { get; set; }
        public Image Image { get; set; }
    }
}