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

        public Form1()
        {
            InitializeComponent();

            myConnection = new OleDbConnection(connectionString);
            myConnection.Open();

            SetupListView();
        }

        private void SetupListView()
        {
            imageList1.ImageSize = new Size(150, 100);
            imageList1.ColorDepth = ColorDepth.Depth32Bit;

            listView1.LargeImageList = imageList1;
            listView1.SmallImageList = imageList1;

            // Устанавливаем режим отображения
            listView1.View = View.LargeIcon;
            // Дополнительные настройки
            listView1.LabelWrap = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                string query = "SELECT Brand, Model, Description, ImageURL FROM Consoles";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                OleDbDataReader reader = command.ExecuteReader();

                int imageIndex = 0;

                while (reader.Read())
                {
                    string brand = reader["Brand"].ToString();
                    string model = reader["Model"].ToString();
                    string description = reader["Description"].ToString();
                    string imageUrl = reader["ImageURL"].ToString();

                    WebClient wc = new WebClient();
                    byte[] bytes = wc.DownloadData(imageUrl);
                    MemoryStream ms = new MemoryStream(bytes);
                    Image image = Image.FromStream(ms);

                    Image resizedImage = new Bitmap(image, imageList1.ImageSize);

                    imageList1.Images.Add(resizedImage);

                    image.Dispose();
                    ms.Dispose();
                    wc.Dispose();

                    ListViewItem item = new ListViewItem();
                    item.Text = brand + model;
                    item.ImageIndex = imageIndex;

                    listView1.Items.Add(item);

                    imageIndex++;
                }

                reader.Close();

                listView1.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (myConnection != null && myConnection.State == ConnectionState.Open)
            {
                myConnection.Close();
            }

            foreach (Image img in imageList1.Images)
            {
                img.Dispose();
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}