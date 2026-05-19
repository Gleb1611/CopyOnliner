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
        private OleDbConnection myConnection;
        private Panel searchPanel;
        private TextBox txtSearchModel;
        private Button btnSearch;
        private Button btnReset;
        private Label lblResults;
        private CheckedListBox chkListShops;
        private Label lblShops;
        private NumericUpDown numPriceFrom;
        private NumericUpDown numPriceTo;
        private Label lblPrice;
        private System.Windows.Forms.Timer searchTimer;
        private ListBox listBoxBrands;
        private Label lblBrand;

        private Label lblSortBy;
        private CheckedListBox chkListSortOptions;
        private RadioButton rbAscending;
        private RadioButton rbDescending;
        private Button btnApplySort;

        private GroupBox groupBoxSpecs;
        private TextBox txtScreenSize;
        private TextBox txtResolution;
        private TextBox txtStorage;
        private TextBox txtRAM;
        private TextBox txtProcessor;
        private TextBox txtBattery;
        private ComboBox cmbOS;

        private Label lblScreenSize;
        private Label lblResolution;
        private Label lblStorage;
        private Label lblRAM;
        private Label lblProcessor;
        private Label lblBattery;
        private Label lblOS;
        private Label lblModel;

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

            this.BackColor = darkBg;
            this.WindowState = FormWindowState.Maximized;
            this.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            searchTimer = new System.Windows.Forms.Timer();
            searchTimer.Interval = 500;
            searchTimer.Tick += SearchTimer_Tick;

            SetupListView();
            InitializeSearchPanel();

            try
            {
                myConnection = new OleDbConnection(connectionString);
                myConnection.Open();
                LoadFilters();
                LoadOSOptions();
                LoadData();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка подключения: {ex.Message}\n\nДобавляю демо-данные...", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                AddDemoData();
            }
        }

        private void LoadFilters()
        {
            try
            {
                string brandQuery = "SELECT DISTINCT Brand FROM Consoles WHERE Brand IS NOT NULL ORDER BY Brand";
                OleDbCommand cmd = new OleDbCommand(brandQuery, myConnection);
                OleDbDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    listBoxBrands.Items.Add(reader["Brand"].ToString());
                }
                reader.Close();

                string shopQuery = "SELECT ShopID, ShopName FROM Shops ORDER BY ShopName";
                cmd = new OleDbCommand(shopQuery, myConnection);
                reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    chkListShops.Items.Add(new ShopItem
                    {
                        ShopID = Convert.ToInt32(reader["ShopID"]),
                        ShopName = reader["ShopName"].ToString()
                    });
                }
                reader.Close();
            }
            catch
            {
                listBoxBrands.Items.AddRange(new string[] { "Sony", "Microsoft", "Nintendo", "Valve", "ASUS" });
                chkListShops.Items.Add(new ShopItem { ShopID = 1, ShopName = "Agroup" });
                chkListShops.Items.Add(new ShopItem { ShopID = 2, ShopName = "GAMEPARK" });
                chkListShops.Items.Add(new ShopItem { ShopID = 3, ShopName = "Newton" });
            }
        }

        private void LoadOSOptions()
        {
            if (cmbOS != null)
            {
                cmbOS.Items.Add("Все");
                cmbOS.Items.Add("Windows");
                cmbOS.Items.Add("PlayStation OS");
                cmbOS.Items.Add("Xbox OS");
                cmbOS.Items.Add("Nintendo OS");
                cmbOS.Items.Add("SteamOS");
                cmbOS.Items.Add("Android");
                cmbOS.SelectedIndex = 0;
            }
        }

        private void InitializeSearchPanel()
        {
            searchPanel = new Panel();
            searchPanel.Dock = DockStyle.Left;
            searchPanel.Width = 380;
            searchPanel.BackColor = panelBg;
            searchPanel.Padding = new Padding(15);
            searchPanel.ForeColor = textColor;
            searchPanel.AutoScroll = true;

            int yOffset = 10;

            Label lblTitle = new Label();
            lblTitle.Text = "⚡ ИГРОВОЙ КАТАЛОГ";
            lblTitle.Location = new Point(15, yOffset);
            lblTitle.Size = new Size(350, 35);
            lblTitle.Font = new Font("Segoe UI", 14, FontStyle.Bold);
            lblTitle.ForeColor = accentColor;
            lblTitle.TextAlign = ContentAlignment.MiddleCenter;
            searchPanel.Controls.Add(lblTitle);
            yOffset += 45;

            Panel separator = new Panel();
            separator.BackColor = accentColor;
            separator.Location = new Point(15, yOffset);
            separator.Size = new Size(350, 2);
            searchPanel.Controls.Add(separator);
            yOffset += 15;

            lblBrand = new Label();
            lblBrand.Text = "🔍 БРЕНД (выберите несколько)";
            lblBrand.Location = new Point(15, yOffset);
            lblBrand.Size = new Size(350, 25);
            lblBrand.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblBrand.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblBrand);
            yOffset += 28;

            listBoxBrands = new ListBox();
            listBoxBrands.Location = new Point(15, yOffset);
            listBoxBrands.Size = new Size(350, 100);
            listBoxBrands.BackColor = darkBg;
            listBoxBrands.ForeColor = textColor;
            listBoxBrands.BorderStyle = BorderStyle.FixedSingle;
            listBoxBrands.SelectionMode = SelectionMode.MultiExtended;
            listBoxBrands.Font = new Font("Segoe UI", 10);
            listBoxBrands.Click += (s, e) => LoadData();
            searchPanel.Controls.Add(listBoxBrands);
            yOffset += 110;

            lblModel = new Label();
            lblModel.Text = "🎮 МОДЕЛЬ";
            lblModel.Location = new Point(15, yOffset);
            lblModel.Size = new Size(350, 25);
            lblModel.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblModel.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblModel);
            yOffset += 28;

            txtSearchModel = new TextBox();
            txtSearchModel.Location = new Point(15, yOffset);
            txtSearchModel.Size = new Size(350, 30);
            txtSearchModel.Font = new Font("Segoe UI", 11);
            txtSearchModel.BackColor = darkBg;
            txtSearchModel.ForeColor = textColor;
            txtSearchModel.BorderStyle = BorderStyle.FixedSingle;
            txtSearchModel.TextChanged += TxtSearch_TextChanged;
            searchPanel.Controls.Add(txtSearchModel);
            yOffset += 40;

            lblPrice = new Label();
            lblPrice.Text = "💰 ЦЕНА (руб.)";
            lblPrice.Location = new Point(15, yOffset);
            lblPrice.Size = new Size(350, 25);
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
            numPriceFrom.Size = new Size(130, 27);
            numPriceFrom.BackColor = darkBg;
            numPriceFrom.ForeColor = textColor;
            numPriceFrom.Maximum = 1000000;
            numPriceFrom.ThousandsSeparator = true;
            numPriceFrom.ValueChanged += TxtSearch_TextChanged;
            searchPanel.Controls.Add(numPriceFrom);

            Label lblTo = new Label();
            lblTo.Text = "до";
            lblTo.Location = new Point(185, yOffset);
            lblTo.Size = new Size(30, 25);
            lblTo.ForeColor = textColor;
            searchPanel.Controls.Add(lblTo);

            numPriceTo = new NumericUpDown();
            numPriceTo.Location = new Point(215, yOffset);
            numPriceTo.Size = new Size(150, 27);
            numPriceTo.BackColor = darkBg;
            numPriceTo.ForeColor = textColor;
            numPriceTo.Maximum = 1000000;
            numPriceTo.ThousandsSeparator = true;
            numPriceTo.ValueChanged += TxtSearch_TextChanged;
            searchPanel.Controls.Add(numPriceTo);
            yOffset += 40;

            lblShops = new Label();
            lblShops.Text = "🏪 МАГАЗИНЫ";
            lblShops.Location = new Point(15, yOffset);
            lblShops.Size = new Size(350, 25);
            lblShops.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblShops.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblShops);
            yOffset += 28;

            chkListShops = new CheckedListBox();
            chkListShops.Location = new Point(15, yOffset);
            chkListShops.Size = new Size(350, 80);
            chkListShops.BackColor = darkBg;
            chkListShops.ForeColor = textColor;
            chkListShops.BorderStyle = BorderStyle.FixedSingle;
            chkListShops.ItemCheck += (s, e) => TxtSearch_TextChanged(s, e);
            searchPanel.Controls.Add(chkListShops);
            yOffset += 90;

            groupBoxSpecs = new GroupBox();
            groupBoxSpecs.Text = "📊 ХАРАКТЕРИСТИКИ";
            groupBoxSpecs.Location = new Point(15, yOffset);
            groupBoxSpecs.Size = new Size(350, 280);
            groupBoxSpecs.ForeColor = accentBlue;
            groupBoxSpecs.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            searchPanel.Controls.Add(groupBoxSpecs);

            int specYOffset = 25;
            int specHeight = 28;

            lblScreenSize = new Label();
            lblScreenSize.Text = "📱 Размер экрана (дюймы):";
            lblScreenSize.Location = new Point(10, specYOffset);
            lblScreenSize.Size = new Size(180, 25);
            lblScreenSize.Font = new Font("Segoe UI", 9);
            lblScreenSize.ForeColor = textColor;
            groupBoxSpecs.Controls.Add(lblScreenSize);

            txtScreenSize = new TextBox();
            txtScreenSize.Location = new Point(195, specYOffset);
            txtScreenSize.Size = new Size(140, 25);
            txtScreenSize.Font = new Font("Segoe UI", 9);
            txtScreenSize.BackColor = darkBg;
            txtScreenSize.ForeColor = textColor;
            txtScreenSize.BorderStyle = BorderStyle.FixedSingle;
            txtScreenSize.TextChanged += TxtSearch_TextChanged;
            groupBoxSpecs.Controls.Add(txtScreenSize);
            specYOffset += specHeight;

            lblResolution = new Label();
            lblResolution.Text = "🎯 Разрешение:";
            lblResolution.Location = new Point(10, specYOffset);
            lblResolution.Size = new Size(180, 25);
            lblResolution.Font = new Font("Segoe UI", 9);
            lblResolution.ForeColor = textColor;
            groupBoxSpecs.Controls.Add(lblResolution);

            txtResolution = new TextBox();
            txtResolution.Location = new Point(195, specYOffset);
            txtResolution.Size = new Size(140, 25);
            txtResolution.Font = new Font("Segoe UI", 9);
            txtResolution.BackColor = darkBg;
            txtResolution.ForeColor = textColor;
            txtResolution.BorderStyle = BorderStyle.FixedSingle;
            txtResolution.TextChanged += TxtSearch_TextChanged;
            groupBoxSpecs.Controls.Add(txtResolution);
            specYOffset += specHeight;

            lblStorage = new Label();
            lblStorage.Text = "💾 Память (ГБ):";
            lblStorage.Location = new Point(10, specYOffset);
            lblStorage.Size = new Size(180, 25);
            lblStorage.Font = new Font("Segoe UI", 9);
            lblStorage.ForeColor = textColor;
            groupBoxSpecs.Controls.Add(lblStorage);

            txtStorage = new TextBox();
            txtStorage.Location = new Point(195, specYOffset);
            txtStorage.Size = new Size(140, 25);
            txtStorage.Font = new Font("Segoe UI", 9);
            txtStorage.BackColor = darkBg;
            txtStorage.ForeColor = textColor;
            txtStorage.BorderStyle = BorderStyle.FixedSingle;
            txtStorage.TextChanged += TxtSearch_TextChanged;
            groupBoxSpecs.Controls.Add(txtStorage);
            specYOffset += specHeight;

            lblRAM = new Label();
            lblRAM.Text = "🧠 ОЗУ (ГБ):";
            lblRAM.Location = new Point(10, specYOffset);
            lblRAM.Size = new Size(180, 25);
            lblRAM.Font = new Font("Segoe UI", 9);
            lblRAM.ForeColor = textColor;
            groupBoxSpecs.Controls.Add(lblRAM);

            txtRAM = new TextBox();
            txtRAM.Location = new Point(195, specYOffset);
            txtRAM.Size = new Size(140, 25);
            txtRAM.Font = new Font("Segoe UI", 9);
            txtRAM.BackColor = darkBg;
            txtRAM.ForeColor = textColor;
            txtRAM.BorderStyle = BorderStyle.FixedSingle;
            txtRAM.TextChanged += TxtSearch_TextChanged;
            groupBoxSpecs.Controls.Add(txtRAM);
            specYOffset += specHeight;

            lblProcessor = new Label();
            lblProcessor.Text = "⚙ Процессор:";
            lblProcessor.Location = new Point(10, specYOffset);
            lblProcessor.Size = new Size(180, 25);
            lblProcessor.Font = new Font("Segoe UI", 9);
            lblProcessor.ForeColor = textColor;
            groupBoxSpecs.Controls.Add(lblProcessor);

            txtProcessor = new TextBox();
            txtProcessor.Location = new Point(195, specYOffset);
            txtProcessor.Size = new Size(140, 25);
            txtProcessor.Font = new Font("Segoe UI", 9);
            txtProcessor.BackColor = darkBg;
            txtProcessor.ForeColor = textColor;
            txtProcessor.BorderStyle = BorderStyle.FixedSingle;
            txtProcessor.TextChanged += TxtSearch_TextChanged;
            groupBoxSpecs.Controls.Add(txtProcessor);
            specYOffset += specHeight;

            lblBattery = new Label();
            lblBattery.Text = "🔋 Батарея (часы):";
            lblBattery.Location = new Point(10, specYOffset);
            lblBattery.Size = new Size(180, 25);
            lblBattery.Font = new Font("Segoe UI", 9);
            lblBattery.ForeColor = textColor;
            groupBoxSpecs.Controls.Add(lblBattery);

            txtBattery = new TextBox();
            txtBattery.Location = new Point(195, specYOffset);
            txtBattery.Size = new Size(140, 25);
            txtBattery.Font = new Font("Segoe UI", 9);
            txtBattery.BackColor = darkBg;
            txtBattery.ForeColor = textColor;
            txtBattery.BorderStyle = BorderStyle.FixedSingle;
            txtBattery.TextChanged += TxtSearch_TextChanged;
            groupBoxSpecs.Controls.Add(txtBattery);
            specYOffset += specHeight;

            lblOS = new Label();
            lblOS.Text = "💿 Операционная система:";
            lblOS.Location = new Point(10, specYOffset);
            lblOS.Size = new Size(180, 25);
            lblOS.Font = new Font("Segoe UI", 9);
            lblOS.ForeColor = textColor;
            groupBoxSpecs.Controls.Add(lblOS);

            cmbOS = new ComboBox();
            cmbOS.Location = new Point(195, specYOffset);
            cmbOS.Size = new Size(140, 25);
            cmbOS.Font = new Font("Segoe UI", 9);
            cmbOS.BackColor = darkBg;
            cmbOS.ForeColor = textColor;
            cmbOS.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbOS.SelectedIndexChanged += TxtSearch_TextChanged;
            groupBoxSpecs.Controls.Add(cmbOS);

            yOffset += 290;

            lblSortBy = new Label();
            lblSortBy.Text = "📊 СОРТИРОВКА (выберите несколько)";
            lblSortBy.Location = new Point(15, yOffset);
            lblSortBy.Size = new Size(350, 25);
            lblSortBy.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblSortBy.ForeColor = accentBlue;
            searchPanel.Controls.Add(lblSortBy);
            yOffset += 28;

            chkListSortOptions = new CheckedListBox();
            chkListSortOptions.Location = new Point(15, yOffset);
            chkListSortOptions.Size = new Size(350, 100);
            chkListSortOptions.BackColor = darkBg;
            chkListSortOptions.ForeColor = textColor;
            chkListSortOptions.BorderStyle = BorderStyle.FixedSingle;
            chkListSortOptions.Items.AddRange(new string[] {
                "Бренд", "Модель", "Цена", "Размер экрана", "Разрешение",
                "Объём памяти", "ОЗУ", "Процессор", "Батарея", "ОС"
            });
            searchPanel.Controls.Add(chkListSortOptions);
            yOffset += 110;

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

            btnApplySort = new Button();
            btnApplySort.Text = "🔄 Применить сортировку";
            btnApplySort.Location = new Point(15, yOffset);
            btnApplySort.Size = new Size(350, 35);
            btnApplySort.BackColor = accentBlue;
            btnApplySort.ForeColor = Color.Black;
            btnApplySort.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            btnApplySort.FlatStyle = FlatStyle.Flat;
            btnApplySort.Cursor = Cursors.Hand;
            btnApplySort.Click += BtnApplySort_Click;
            searchPanel.Controls.Add(btnApplySort);
            yOffset += 45;

            btnSearch = new Button();
            btnSearch.Text = "🔍 ПОИСК";
            btnSearch.Location = new Point(15, yOffset);
            btnSearch.Size = new Size(170, 45);
            btnSearch.BackColor = accentColor;
            btnSearch.ForeColor = Color.Black;
            btnSearch.Font = new Font("Segoe UI", 11, FontStyle.Bold);
            btnSearch.FlatStyle = FlatStyle.Flat;
            btnSearch.Cursor = Cursors.Hand;
            btnSearch.Click += BtnSearch_Click;
            searchPanel.Controls.Add(btnSearch);

            btnReset = new Button();
            btnReset.Text = "🔄 СБРОС";
            btnReset.Location = new Point(195, yOffset);
            btnReset.Size = new Size(170, 45);
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
            lblResults.Size = new Size(350, 30);
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
                    LoadData();
                }
            }
        }

        private void LoadData()
        {
            try
            {
                listView1.BeginUpdate();
                listView1.Items.Clear();
                imageList1.Images.Clear();

                // Простой запрос без JOIN (только из таблицы Consoles)
                string query = @"
                    SELECT 
                        ConsoleID, 
                        Brand, 
                        Model, 
                        Description, 
                        ImageURL, 
                        OS, 
                        Weight, 
                        BatteryLife, 
                        ScreenSize, 
                        Resolution,
                        Storage, 
                        RAM, 
                        Processor,
                        Price
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

                if (!string.IsNullOrWhiteSpace(txtScreenSize.Text))
                {
                    query += $" AND ScreenSize LIKE '%{txtScreenSize.Text.Replace("'", "''")}%'";
                }
                if (!string.IsNullOrWhiteSpace(txtResolution.Text))
                {
                    query += $" AND Resolution LIKE '%{txtResolution.Text.Replace("'", "''")}%'";
                }
                if (!string.IsNullOrWhiteSpace(txtStorage.Text))
                {
                    query += $" AND Storage LIKE '%{txtStorage.Text.Replace("'", "''")}%'";
                }
                if (!string.IsNullOrWhiteSpace(txtRAM.Text))
                {
                    query += $" AND RAM LIKE '%{txtRAM.Text.Replace("'", "''")}%'";
                }
                if (!string.IsNullOrWhiteSpace(txtProcessor.Text))
                {
                    query += $" AND Processor LIKE '%{txtProcessor.Text.Replace("'", "''")}%'";
                }
                if (!string.IsNullOrWhiteSpace(txtBattery.Text))
                {
                    query += $" AND BatteryLife LIKE '%{txtBattery.Text.Replace("'", "''")}%'";
                }
                if (cmbOS.SelectedIndex > 0 && cmbOS.SelectedItem.ToString() != "Все")
                {
                    query += $" AND OS LIKE '%{cmbOS.SelectedItem.ToString().Replace("'", "''")}%'";
                }

                // Сортировка
                string orderBy = BuildOrderByClause();
                if (!string.IsNullOrEmpty(orderBy))
                {
                    query += $" ORDER BY {orderBy}";
                }
                else
                {
                    query += " ORDER BY Brand, Model";
                }

                OleDbCommand command = new OleDbCommand(query, myConnection);
                OleDbDataReader reader = command.ExecuteReader();

                int imageIndex = 0;

                while (reader.Read())
                {
                    ConsoleItem console = new ConsoleItem();
                    console.ConsoleId = Convert.ToInt32(reader["ConsoleID"]);
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
                    console.ShopName = "Не указан";

                    // Загрузка изображения
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

                    // Формирование строки с ценой и характеристиками
                    string priceText = console.Price > 0 ? $"💰 {console.Price:N0} ₽" : "💰 Цена не указана";
                    string specs = $"📱 {console.ScreenSize}\" | 🎯 {console.Resolution} | 💾 {console.Storage}GB | 🧠 {console.RAM}GB";

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
                    case "Разрешение":
                        orderByFields.Add($"Resolution {sortOrder}");
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
                    case "Батарея":
                        orderByFields.Add($"BatteryLife {sortOrder}");
                        break;
                    case "ОС":
                        orderByFields.Add($"OS {sortOrder}");
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
                new { Brand = "Sony", Model = "PlayStation 5", Price = 49999, ScreenSize = "N/A", Resolution = "4K", Storage = "825", RAM = "16", Processor = "AMD Zen 2", BatteryLife = "N/A", OS = "PlayStation OS", Shop = "GAMEPARK", Description = "Новейшая консоль Sony" },
                new { Brand = "Microsoft", Model = "Xbox Series X", Price = 45999, ScreenSize = "N/A", Resolution = "4K", Storage = "1024", RAM = "16", Processor = "AMD Zen 2", BatteryLife = "N/A", OS = "Xbox OS", Shop = "Agroup", Description = "Самая мощная консоль Xbox" },
                new { Brand = "Nintendo", Model = "Switch OLED", Price = 29999, ScreenSize = "7", Resolution = "1280x720", Storage = "64", RAM = "4", Processor = "NVIDIA Tegra", BatteryLife = "4-9", OS = "Nintendo OS", Shop = "Newton", Description = "Гибридная консоль" },
                new { Brand = "Valve", Model = "Steam Deck", Price = 39999, ScreenSize = "7", Resolution = "1280x800", Storage = "512", RAM = "16", Processor = "AMD APU", BatteryLife = "2-8", OS = "SteamOS", Shop = "GAMEPARK", Description = "Портативный компьютер" },
                new { Brand = "ASUS", Model = "ROG Ally", Price = 44999, ScreenSize = "7", Resolution = "1920x1080", Storage = "512", RAM = "16", Processor = "AMD Z1 Extreme", BatteryLife = "3-6", OS = "Windows 11", Shop = "Agroup", Description = "Игровая портативная консоль" }
            };

            int imageIndex = 0;
            foreach (var demo in demos)
            {
                ConsoleItem console = new ConsoleItem();
                console.ConsoleId = imageIndex + 1;
                console.Brand = demo.Brand;
                console.Model = demo.Model;
                console.Price = demo.Price;
                console.ShopName = demo.Shop;
                console.Description = demo.Description;
                console.OS = demo.OS;
                console.BatteryLife = demo.BatteryLife;
                console.ScreenSize = demo.ScreenSize;
                console.Resolution = demo.Resolution;
                console.Storage = demo.Storage;
                console.RAM = demo.RAM;
                console.Processor = demo.Processor;
                console.Image = CreateBrandPlaceholderImage(demo.Brand);

                imageList1.Images.Add(console.Image);

                string priceText = console.Price > 0 ? $"💰 {console.Price:N0} ₽" : "💰 Цена не указана";
                string specs = $"📱 {console.ScreenSize}\" | 🎯 {console.Resolution} | 💾 {console.Storage}GB | 🧠 {console.RAM}GB";

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
                Rectangle imageRect = new Rectangle(bounds.X + 10, bounds.Y + 10, bounds.Width - 20, 180);
                e.Graphics.DrawImage(img, imageRect);
            }

            string[] textLines = e.Item.Text.Split('\n');
            using (Font titleFont = new Font("Segoe UI", 11, FontStyle.Bold))
            using (Font priceFont = new Font("Segoe UI", 10, FontStyle.Bold))
            using (Font specsFont = new Font("Segoe UI", 8))
            using (SolidBrush titleBrush = new SolidBrush(textColor))
            using (SolidBrush priceBrush = new SolidBrush(accentColor))
            using (SolidBrush specsBrush = new SolidBrush(Color.FromArgb(180, 180, 180)))
            {
                if (textLines.Length > 0)
                {
                    e.Graphics.DrawString(textLines[0], titleFont, titleBrush,
                        new RectangleF(bounds.X + 10, bounds.Y + 200, bounds.Width - 20, 30));
                }
                if (textLines.Length > 1)
                {
                    e.Graphics.DrawString(textLines[1], priceFont, priceBrush,
                        new RectangleF(bounds.X + 10, bounds.Y + 225, bounds.Width - 20, 30));
                }
                if (textLines.Length > 2)
                {
                    e.Graphics.DrawString(textLines[2], specsFont, specsBrush,
                        new RectangleF(bounds.X + 10, bounds.Y + 250, bounds.Width - 20, 40));
                }
            }

            e.DrawDefault = false;
        }

        private void BtnSearch_Click(object sender, EventArgs e) => LoadData();
        private void BtnApplySort_Click(object sender, EventArgs e) => LoadData();

        private void BtnReset_Click(object sender, EventArgs e)
        {
            listBoxBrands.ClearSelected();
            txtSearchModel.Text = "";
            numPriceFrom.Value = 0;
            numPriceTo.Value = 0;
            txtScreenSize.Text = "";
            txtResolution.Text = "";
            txtStorage.Text = "";
            txtRAM.Text = "";
            txtProcessor.Text = "";
            txtBattery.Text = "";
            cmbOS.SelectedIndex = 0;

            for (int i = 0; i < chkListShops.Items.Count; i++)
                chkListShops.SetItemChecked(i, false);

            for (int i = 0; i < chkListSortOptions.Items.Count; i++)
                chkListSortOptions.SetItemChecked(i, false);

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
            if (searchTimer != null)
            {
                searchTimer.Stop();
                LoadData();
            }
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
        public Image Image { get; set; }
    }

    public class ShopItem
    {
        public int ShopID { get; set; }
        public string ShopName { get; set; }
        public override string ToString() => ShopName;
    }
}