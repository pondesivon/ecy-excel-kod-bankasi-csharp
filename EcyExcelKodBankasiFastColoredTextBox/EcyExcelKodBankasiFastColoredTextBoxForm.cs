using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EcyExcelKodBankasiFastColoredTextBox
{
    public partial class EcyExcelKodBankasiFastColoredTextBoxForm : Form
    {
        public EcyExcelKodBankasiFastColoredTextBoxForm()
        {
            InitializeComponent();
        }

        #region Metotlar

        void BaslangicAyarlari()
        {
            aranacakYerToolStripComboBox.SelectedIndex = 0;
        }

        void VeriYukleDataGridView(DataGridView dataGridView)
        {
            //--------------------------------------------------
            //Veri Tabanı Tanımlamaları
            //--------------------------------------------------
            OleDbConnection oleDbConnection = new OleDbConnection(
                @"Provider=Microsoft.ACE.OLEDB.16.0;Data Source="
                + Path.Combine(Application.StartupPath, "EcyExcelKodBankasi.accdb"));

            DataSet dataSet = new DataSet();
            OleDbCommand oleDbCommand = new OleDbCommand();
            OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter();

            //--------------------------------------------------
            //DataGridView Temizlik İşlemi
            //--------------------------------------------------
            dataGridView.Columns.Clear();
            dataSet.Tables.Clear();
            dataGridView.Refresh();

            //--------------------------------------------------
            //DataGridView Atama İşlemleri
            //--------------------------------------------------
            dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            dataGridView.Columns.Add("idDataGridViewColumn", "ID");
            dataGridView.Columns.Add("baslikDataGridViewColumn", "Başlık");
            dataGridView.Columns.Add("icerikDataGridViewColumn", "İçerik");
            dataGridView.Columns.Add("kategoriDataGridViewColumn", "Kategori");
            dataGridView.Columns.Add("tarihDataGridViewColumn", "Tarih");

            dataGridView.Columns[0].Width = 0;
            dataGridView.Columns[1].Width = dataGridView.Width - 65;
            dataGridView.Columns[2].Width = 0;
            dataGridView.Columns[3].Width = 0;
            dataGridView.Columns[4].Width = 0;

            //--------------------------------------------------
            //Listeleme İşlemi
            //--------------------------------------------------
            oleDbConnection.Open();
            oleDbCommand.Connection = oleDbConnection;
            oleDbCommand.CommandText = "SELECT * From Kodlar";

            OleDbDataReader reader = oleDbCommand.ExecuteReader();

            int sayac = 0;

            while (reader.Read())
            {
                dataGridView.Rows.Add();

                dataGridView.Rows[sayac].Cells["idDataGridViewColumn"].Value = reader[0].ToString();
                dataGridView.Rows[sayac].Cells["baslikDataGridViewColumn"].Value = reader[1].ToString();
                dataGridView.Rows[sayac].Cells["icerikDataGridViewColumn"].Value = reader[2].ToString();
                dataGridView.Rows[sayac].Cells["kategoriDataGridViewColumn"].Value = reader[3].ToString();
                dataGridView.Rows[sayac].Cells["tarihDataGridViewColumn"].Value = reader[4].ToString();
                sayac++;
            }

            oleDbConnection.Close();
        }

        void VeriFiltreleDataGridView(DataGridView dataGridView, string filtre = "*")
        {
            //--------------------------------------------------
            //Veri Tabanı Tanımlamaları
            //--------------------------------------------------
            OleDbConnection oleDbConnection = new OleDbConnection(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
                + Path.Combine(Application.StartupPath, "EcyExcelKodBankasi.accdb"));

            DataSet dataSet = new DataSet();
            OleDbCommand oleDbCommand = new OleDbCommand();
            OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter();

            //--------------------------------------------------
            //DataGridView Temizlik İşlemi
            //--------------------------------------------------
            dataGridView.Columns.Clear();
            dataSet.Tables.Clear();
            dataGridView.Refresh();

            //--------------------------------------------------
            //DataGridView Atama İşlemleri
            //--------------------------------------------------
            dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            dataGridView.Columns.Add("idFiltreDataGridViewColumn", "ID");
            dataGridView.Columns.Add("baslikFiltreDataGridViewColumn", "Başlık");
            dataGridView.Columns.Add("icerikFiltreDataGridViewColumn", "İçerik");
            dataGridView.Columns.Add("kategoriFiltreDataGridViewColumn", "Kategori");
            dataGridView.Columns.Add("tarihFiltreDataGridViewColumn", "Tarih");

            dataGridView.Columns[0].Width = 0;
            dataGridView.Columns[1].Width = dataGridView.Width - 65;
            dataGridView.Columns[2].Width = 0;
            dataGridView.Columns[3].Width = 0;
            dataGridView.Columns[4].Width = 0;

            //--------------------------------------------------
            //Listeleme İşlemi
            //--------------------------------------------------
            oleDbConnection.Open();
            oleDbCommand.Connection = oleDbConnection;

            if (filtre == "*" || filtre == "")
            {
                oleDbCommand.CommandText = "SELECT * From Kodlar";
            }
            else
            {
                if (aranacakYerToolStripComboBox.Text == "Başlık")
                {
                    oleDbCommand.CommandText = "SELECT * From Kodlar WHERE LCase(baslik) LIKE LCase(@baslik)";
                    oleDbCommand.Parameters.AddWithValue("@baslik", "%" + araToolStripComboBox.Text + "%");
                }
                else if (aranacakYerToolStripComboBox.Text == "İçerik")
                {
                    oleDbCommand.CommandText = "SELECT * From Kodlar WHERE LCase(icerik) LIKE LCase(@icerik)";
                    oleDbCommand.Parameters.AddWithValue("@icerik", "%" + araToolStripComboBox.Text + "%");
                }
            }

            OleDbDataReader reader = oleDbCommand.ExecuteReader();

            int sayac = 0;

            while (reader.Read())
            {
                dataGridView.Rows.Add();

                dataGridView.Rows[sayac].Cells["idFiltreDataGridViewColumn"].Value = reader[0].ToString();
                dataGridView.Rows[sayac].Cells["baslikFiltreDataGridViewColumn"].Value = reader[1].ToString();
                dataGridView.Rows[sayac].Cells["icerikFiltreDataGridViewColumn"].Value = reader[2].ToString();
                dataGridView.Rows[sayac].Cells["kategoriFiltreDataGridViewColumn"].Value = reader[3].ToString();
                dataGridView.Rows[sayac].Cells["tarihFiltreDataGridViewColumn"].Value = reader[4].ToString();
                sayac++;
            }

            oleDbConnection.Close();
        }

        void VeriEkle(string baslik, string icerik, string kategori)
        {
            try
            {
                //--------------------------------------------------
                //Değişken Tanımlama İşlemi
                //--------------------------------------------------
                OleDbConnection oleDbConnection = new OleDbConnection(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
                    + Path.Combine(Application.StartupPath, "EcyExcelKodBankasi.accdb"));

                string sorgu = "INSERT INTO Kodlar(baslik, icerik, kategori) "
                    + "VALUES(@baslik, @icerik, @kategori)";

                OleDbCommand oleDbCommand = new OleDbCommand(sorgu, oleDbConnection);
                oleDbCommand.Parameters.AddWithValue("@baslik", baslik);
                oleDbCommand.Parameters.AddWithValue("@icerik", icerik);
                oleDbCommand.Parameters.AddWithValue("@kategori", kategori != "" ? kategori : "GENEL");

                //--------------------------------------------------
                //Ekleme İşlemi
                //--------------------------------------------------
                oleDbConnection.Open();
                oleDbCommand.ExecuteNonQuery();
                oleDbConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        void VeriGuncelle(string baslik, string icerik, string kategori, int id)
        {
            try
            {
                //--------------------------------------------------
                //Değişken Tanımlama İşlemi
                //--------------------------------------------------
                OleDbConnection oleDbConnection = new OleDbConnection(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
                    + Path.Combine(Application.StartupPath, "EcyExcelKodBankasi.accdb"));

                string sorgu = "UPDATE Kodlar SET "
                    + "baslik=@baslik, icerik=@icerik, kategori=@kategori "
                    + "WHERE id = @id";

                OleDbCommand oleDbCommand = new OleDbCommand(sorgu, oleDbConnection);
                oleDbCommand.Parameters.AddWithValue("@baslik", baslik);
                oleDbCommand.Parameters.AddWithValue("@icerik", icerik);
                oleDbCommand.Parameters.AddWithValue("@kategori", kategori != "" ? kategori : "GENEL");
                oleDbCommand.Parameters.AddWithValue("@id", id);

                //--------------------------------------------------
                //Ekleme İşlemi
                //--------------------------------------------------
                oleDbConnection.Open();
                oleDbCommand.ExecuteNonQuery();
                oleDbConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        void VeriSil(int id)
        {
            try
            {
                //--------------------------------------------------
                //Değişken Tanımlama İşlemi
                //--------------------------------------------------
                OleDbConnection oleDbConnection = new OleDbConnection(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
                    + Path.Combine(Application.StartupPath, "EcyExcelKodBankasi.accdb"));

                string sorgu = "DELETE FROM Kodlar "
                    + "WHERE id = @id";

                OleDbCommand oleDbCommand = new OleDbCommand(sorgu, oleDbConnection);
                oleDbCommand.Parameters.AddWithValue("@id", id);

                //--------------------------------------------------
                //Ekleme İşlemi
                //--------------------------------------------------
                oleDbConnection.Open();
                oleDbCommand.ExecuteNonQuery();
                oleDbConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        void IcerigiFormAlaninaAktar(DataGridView dataGridView)
        {
            idToolStripTextBox.Text = dataGridView.SelectedRows[0].Cells[0].Value.ToString();
            baslikRichTextBox.Text = dataGridView.SelectedRows[0].Cells[1].Value.ToString();
            icerikFastColoredTextBox.Text = dataGridView.SelectedRows[0].Cells[2].Value.ToString();
            kategoriToolStripTextBox.Text = dataGridView.SelectedRows[0].Cells[3].Value.ToString();
            eklenmeTarihiToolStripTextBox.Text = dataGridView.SelectedRows[0].Cells[4].Value.ToString();
        }

        #endregion

        #region Olaylar

        void Ctrl_F_AramaKutusunaGec(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.F)
            {
                araToolStripComboBox.Focus();
                araToolStripComboBox.SelectAll();
            }
        }

        #endregion

        private void EcyExcelKodBankasiFctbForm_Load(object sender, EventArgs e)
        {
            BaslangicAyarlari();
            VeriYukleDataGridView(listeDataGridView);
        }

        private void araToolStripComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                e.SuppressKeyPress = (e.KeyData == Keys.Enter);

                sidebarSolTabControl.SelectedTab = sidebarSolTabControl.TabPages["aramaSonuclariTabPage"];
                VeriFiltreleDataGridView(filtreListeDataGridView, araToolStripComboBox.Text);
                araToolStripComboBox.Focus();
            }
        }

        private void listeDataGridView_Click(object sender, EventArgs e)
        {
            IcerigiFormAlaninaAktar(listeDataGridView);
        }

        private void seciliMetniKopyalaToolStripButton_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(icerikFastColoredTextBox.SelectedText);
        }

        private void tumunuKopyalaToolStripButton_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(icerikFastColoredTextBox.Text);
        }

        private void yapistirToolStripButton_Click(object sender, EventArgs e)
        {
            icerikFastColoredTextBox.Text = Clipboard.GetText();
        }

        private void alaniTemizleToolStripButton_Click(object sender, EventArgs e)
        {
            idToolStripTextBox.Text = String.Empty;
            baslikRichTextBox.Text = String.Empty;
            icerikFastColoredTextBox.Text = String.Empty;
            kategoriToolStripTextBox.Text = String.Empty;
            eklenmeTarihiToolStripTextBox.Text = String.Empty;
        }

        private void icerikFastColoredTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            Ctrl_F_AramaKutusunaGec(sender, e);
        }

        private void ekleToolStripButton_Click(object sender, EventArgs e)
        {
            VeriEkle(
                baslikRichTextBox.Text,
                icerikFastColoredTextBox.Text,
                kategoriToolStripTextBox.Text);
        }

        private void guncelleToolStripButton_Click(object sender, EventArgs e)
        {
            VeriGuncelle(
                baslikRichTextBox.Text,
                icerikFastColoredTextBox.Text,
                kategoriToolStripTextBox.Text,
                Convert.ToInt32(idToolStripTextBox.Text));
        }

        private void silToolStripButton_Click(object sender, EventArgs e)
        {
            VeriSil(Convert.ToInt32(idToolStripTextBox.Text));
        }

        private void yenileToolStripButton_Click(object sender, EventArgs e)
        {
            VeriYukleDataGridView(listeDataGridView);
        }

        private void filtreListeDataGridView_Click(object sender, EventArgs e)
        {
            IcerigiFormAlaninaAktar(filtreListeDataGridView);
        }
    }
}
