using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Net;
using System.Windows.Forms;

namespace CopyOnliner
{
    public partial class Form1 : Form
    {
        public static string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=OnlinerConsoles.accdb;";
        private OleDbConnection myConnection;
        private Panel searchPanel;
        private TextBox txtSearchBrand;
        private TextBox txtSearchModel;
        private ComboBox cmbSortBy;
        private Button btnSearch;
        private Button btnReset;
        private Label lblBrand;
        private Label lblModel;
        private Label lblSort;

        public Form1()
        {
            InitializeComponent();

            // Отключаем автоматическое позиционирование
            this.SuspendLayout();

            // Открываем на весь экран
            this.WindowState = FormWindowState.Maximized;

            // Сначала настраиваем ListView
            SetupListView();

            // Затем создаем панель поиска
            InitializeSearchPanel();

            // Возобновляем позиционирование
            this.ResumeLayout(false);

            myConnection = new OleDbConnection(connectionString);
            myConnection.Open();
        }

        private void InitializeSearchPanel()
        {
            // Создаем панель поиска с фиксированной шириной
            searchPanel = new Panel();
            searchPanel.Dock = DockStyle.Left;
            searchPanel.Width = 280;
            searchPanel.BackColor = Color.FromArgb(240, 240, 240);
            searchPanel.Padding = new Padding(15);

            // Создаем элементы управления для поиска
            lblBrand = new Label();
            lblBrand.Text = "Бренд:";
            lblBrand.Location = new Point(15, 15);
            lblBrand.Size = new Size(250, 25);
            lblBrand.Font = new Font("Arial", 11, FontStyle.Bold);

            txtSearchBrand = new TextBox();
            txtSearchBrand.Location = new Point(15, 45);
            txtSearchBrand.Size = new Size(250, 27);
            txtSearchBrand.Font = new Font("Arial", 11);

            lblModel = new Label();
            lblModel.Text = "Модель:";
            lblModel.Location = new Point(15, 85);
            lblModel.Size = new Size(250, 25);
            lblModel.Font = new Font("Arial", 11, FontStyle.Bold);

            txtSearchModel = new TextBox();
            txtSearchModel.Location = new Point(15, 115);
            txtSearchModel.Size = new Size(250, 27);
            txtSearchModel.Font = new Font("Arial", 11);

            lblSort = new Label();
            lblSort.Text = "Сортировка:";
            lblSort.Location = new Point(15, 155);
            lblSort.Size = new Size(250, 25);
            lblSort.Font = new Font("Arial", 11, FontStyle.Bold);

            cmbSortBy = new ComboBox();
            cmbSortBy.Location = new Point(15, 185);
            cmbSortBy.Size = new Size(250, 29);
            cmbSortBy.Font = new Font("Arial", 11);
            cmbSortBy.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbSortBy.Items.AddRange(new string[] { "По бренду и модели", "По бренду", "По модели" });
            cmbSortBy.SelectedIndex = 0;

            btnSearch = new Button();
            btnSearch.Text = "Найти";
            btnSearch.Location = new Point(15, 235);
            btnSearch.Size = new Size(115, 40);
            btnSearch.BackColor = Color.LightGreen;
            btnSearch.Font = new Font("Arial", 11, FontStyle.Bold);
            btnSearch.FlatStyle = FlatStyle.Flat;
            btnSearch.Click += BtnSearch_Click;

            btnReset = new Button();
            btnReset.Text = "Сбросить";
            btnReset.Location = new Point(150, 235);
            btnReset.Size = new Size(115, 40);
            btnReset.BackColor = Color.LightCoral;
            btnReset.Font = new Font("Arial", 11, FontStyle.Bold);
            btnReset.FlatStyle = FlatStyle.Flat;
            btnReset.Click += BtnReset_Click;

            // Добавляем элементы на панель
            searchPanel.Controls.AddRange(new Control[] {
                lblBrand, txtSearchBrand,
                lblModel, txtSearchModel,
                lblSort, cmbSortBy,
                btnSearch, btnReset
            });

            // Сначала добавляем панель поиска на форму
            this.Controls.Add(searchPanel);

            // Затем убеждаемся, что ListView правильно настроен
            if (listView1 != null)
            {
                // Убираем DockStyle.Fill временно
                listView1.Dock = DockStyle.None;
                // Устанавливаем правильные границы
                listView1.Left = searchPanel.Width;
                listView1.Top = 0;
                listView1.Width = this.ClientSize.Width - searchPanel.Width;
                listView1.Height = this.ClientSize.Height;
                listView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            }
        }

        private void SetupListView()
        {
            // Настраиваем существующий listView1 из дизайнера
            imageList1.ImageSize = new Size(320, 240);
            imageList1.ColorDepth = ColorDepth.Depth32Bit;

            listView1.LargeImageList = imageList1;
            listView1.SmallImageList = imageList1;

            // Устанавливаем режим отображения карточек
            listView1.View = View.LargeIcon;
            listView1.LabelWrap = true;
            listView1.OwnerDraw = true;

            // Подписываемся на события
            listView1.DrawItem += ListView1_DrawItem;
            listView1.DrawSubItem += ListView1_DrawSubItem;
            listView1.DrawColumnHeader += ListView1_DrawColumnHeader;
            listView1.SelectedIndexChanged += listView1_SelectedIndexChanged;

            // Включаем сортировку
            listView1.Sorting = SortOrder.Ascending;

            // Увеличиваем шрифт для текста карточек
            listView1.Font = new Font(listView1.Font.FontFamily, listView1.Font.Size + 2, FontStyle.Bold);

            // Убираем Dock (будем управлять вручную)
            listView1.Dock = DockStyle.None;
        }

        private void LoadData()
        {
            try
            {
                // Очищаем текущие данные
                listView1.Items.Clear();
                imageList1.Images.Clear();

                // Строим запрос с условиями поиска
                string query = "SELECT Brand, Model, Description, ImageURL FROM Consoles WHERE 1=1";

                if (!string.IsNullOrWhiteSpace(txtSearchBrand.Text))
                {
                    query += $" AND Brand LIKE '%{txtSearchBrand.Text.Replace("'", "''")}%'";
                }

                if (!string.IsNullOrWhiteSpace(txtSearchModel.Text))
                {
                    query += $" AND Model LIKE '%{txtSearchModel.Text.Replace("'", "''")}%'";
                }

                // Добавляем сортировку
                switch (cmbSortBy.SelectedIndex)
                {
                    case 0:
                        query += " ORDER BY Brand, Model";
                        break;
                    case 1:
                        query += " ORDER BY Brand";
                        break;
                    case 2:
                        query += " ORDER BY Model";
                        break;
                }

                OleDbCommand command = new OleDbCommand(query, myConnection);
                OleDbDataReader reader = command.ExecuteReader();

                int imageIndex = 0;

                while (reader.Read())
                {
                    string brand = reader["Brand"].ToString();
                    string model = reader["Model"].ToString();
                    string description = reader["Description"].ToString();
                    string imageUrl = reader["ImageURL"].ToString();

                    using (WebClient wc = new WebClient())
                    {
                        byte[] bytes = wc.DownloadData(imageUrl);
                        using (MemoryStream ms = new MemoryStream(bytes))
                        {
                            using (Image image = Image.FromStream(ms))
                            {
                                Image resizedImage = new Bitmap(image, imageList1.ImageSize);
                                imageList1.Images.Add(resizedImage);
                            }
                        }
                    }

                    ListViewItem item = new ListViewItem();
                    item.Text = brand + " " + model;
                    item.ImageIndex = imageIndex;
                    item.ToolTipText = description;

                    listView1.Items.Add(item);
                    imageIndex++;
                }

                reader.Close();
                listView1.Refresh();

                this.Text = $"CopyOnliner - Найдено товаров: {imageIndex}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}");
            }
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            txtSearchBrand.Text = "";
            txtSearchModel.Text = "";
            cmbSortBy.SelectedIndex = 0;
            LoadData();
        }

        private void ListView1_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            e.DrawDefault = false;
        }

        private void ListView1_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            e.DrawDefault = false;
        }

        private void ListView1_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
            Rectangle bounds = new Rectangle(e.Bounds.X, e.Bounds.Y, 340, 280);

            if (e.Item.Selected)
            {
                e.Graphics.FillRectangle(Brushes.LightBlue, bounds);
                ControlPaint.DrawBorder(e.Graphics, bounds, Color.Blue, ButtonBorderStyle.Solid);
            }
            else
            {
                e.Graphics.FillRectangle(Brushes.White, bounds);
                ControlPaint.DrawBorder(e.Graphics, bounds, Color.Gray, ButtonBorderStyle.Solid);
            }

            if (e.Item.ImageIndex != -1 && imageList1.Images.Count > e.Item.ImageIndex)
            {
                Image img = imageList1.Images[e.Item.ImageIndex];
                Rectangle imageRect = new Rectangle(bounds.X + 10, bounds.Y + 10, 300, 220);
                e.Graphics.DrawImage(img, imageRect);
            }

            string text = e.Item.Text;
            using (StringFormat sf = new StringFormat())
            {
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;
                RectangleF textRect = new RectangleF(bounds.X, bounds.Y + 235, bounds.Width, 35);
                e.Graphics.DrawString(text, listView1.Font, Brushes.Black, textRect, sf);
            }

            e.DrawDefault = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (myConnection != null && myConnection.State == ConnectionState.Open)
            {
                myConnection.Close();
            }

            if (imageList1 != null)
            {
                foreach (Image img in imageList1.Images)
                {
                    img.Dispose();
                }
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                ListViewItem selectedItem = listView1.SelectedItems[0];
                // Можно добавить дополнительную логику при выборе товара
            }
        }

        // Обработчик изменения размера формы
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
}