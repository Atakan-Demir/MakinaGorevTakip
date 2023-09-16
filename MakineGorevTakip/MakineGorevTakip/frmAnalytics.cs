using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Runtime.InteropServices;

namespace MakineGorevTakip
{
    public partial class frmAnalytics : Form
    {
        public frmAnalytics()
        {
            
            InitializeComponent();
        }

        

        void Veriler()
        {
            string filePath = Form1.filePath;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Excel.Range usedRange = worksheet.UsedRange;

            // Verileri DataTable nesnesine aktar
            System.Data.DataTable dt = new System.Data.DataTable();
            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;

            // Sütun başlıklarını ekleyin
            for (int j = 1; j <= usedRange.Columns.Count; j++)
            {
                if (worksheet.Cells[1, j] != null && worksheet.Cells[1, j].Value != null)
                {
                    dt.Columns.Add((worksheet.Cells[1, j].Value).ToString());
                }
            }

            // Verileri satır satır DataTable'a ekleyin
            for (int i = 2; i <= usedRange.Rows.Count; i++)
            {
                DataRow dr = dt.NewRow();

                bool rowHasData = false;

                for (int j = 1; j <= usedRange.Columns.Count; j++)
                {
                    if (worksheet.Cells[i, j] != null && worksheet.Cells[i, j].Value != null)
                    {
                        dr[j - 1] = (worksheet.Cells[i, j].Value).ToString();
                        rowHasData = true;
                    }
                }

                if (rowHasData)
                {
                    dt.Rows.Add(dr);
                }
            }

            // DataGridView kontrolüne DataTable nesnesini atayın
            dataGridView1.DataSource = dt;
            // Tüm sütunları gizle
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                col.Visible = false;
            }

            // Sadece ikinci sütunu görünür yap
            dataGridView1.Columns[0].Visible = true;

            // Excel uygulamasını kapatın
            //workbook.Close();           
            //excelApp.Quit();
            // Workbook ve Worksheet nesnelerini kapat
            workbook.Close(false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);

            // Excel uygulamasını kapat
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


        }
        private void frmAnalytics_Load(object sender, EventArgs e)
        {
            Veriler();
        }

        private void btnKaydet_Click(object sender, EventArgs e)
        {/*
            // textbox1'deki metni al
            string metin = txtMakAd.Text;

            // Excel dosyasını aç
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(@"C:\Users\DmR\Desktop\db.xlsx");
            Excel._Worksheet worksheet = workbook.Sheets[1];
            Excel.Range usedRange = worksheet.UsedRange;

            // Son satırı bul
            int lastRow = usedRange.Rows.Count + 1;

            // Metni 1. sütuna ekle
            worksheet.Cells[lastRow, 1] = metin;

            // Excel dosyasını kaydet ve kapat
            workbook.Save();
            workbook.Close();
            MessageBox.Show("Makine Eklendi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Veriler();
        */
            // Excel dosyasını aç
            string filePath = Form1.filePath;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel._Worksheet worksheet = workbook.Sheets[1];
            Excel.Range usedRange = worksheet.UsedRange;

            // 1. sütundaki tüm verileri bir List<string> nesnesine aktar
            List<string> firstColumnValues = new List<string>();
            for (int i = 2; i <= usedRange.Rows.Count; i++)
            {
                if (worksheet.Cells[i, 1] != null && worksheet.Cells[i, 1].Value != null)
                {
                    firstColumnValues.Add((worksheet.Cells[i, 1].Value).ToString());

                }
            }

            // Ekleme yapılacak metni al
            string newValue = txtMakAd.Text.Trim();

            // Eğer metin daha önce eklenmemişse, 1. sütuna ekle
            if (!firstColumnValues.Contains(newValue))
            {
                // Yeni satır ekle
                int newRow = usedRange.Rows.Count + 1;

                // 1. sütuna yeni veriyi ekle
                worksheet.Cells[newRow, 1] = newValue;

                // Excel dosyasını kaydet
                workbook.Save();
                MessageBox.Show("Makine Eklendi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Makine Zaten Ekli! Lütfen eklemeye çalıştığınız makinanın mevcut olmadığından emin olun.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Excel dosyasını kapat
            workbook.Close();
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(worksheet);
            Veriler();

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int secilen = dataGridView1.SelectedCells[0].RowIndex;
            txtMakAd.Text = dataGridView1.Rows[secilen].Cells[0].Value.ToString();

        }

        private void btnGuncelle_Click(object sender, EventArgs e)
        {
            string filePath = Form1.filePath;
            // Seçili hücrenin satır ve sütun indekslerini al
            int row = dataGridView1.CurrentCell.RowIndex +1;
            int col = dataGridView1.CurrentCell.ColumnIndex +1;

            // Düzenlenecek veriyi textbox'lardan al
            string yeniDeger = txtMakAd.Text;

            // Seçili hücrenin değerini güncelle
            dataGridView1.Rows[row].Cells[col].Value = yeniDeger;

            // Excel dosyasını aç
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            

            // Veriyi Excel dosyasına kaydet
            var excelWorksheet = workbook.Worksheets[1];
            excelWorksheet.Cells[row + 1, col] = yeniDeger;
            workbook.Save();

            // Excel dosyasını kapat
            workbook.Close();
            excelApp.Quit();
            Veriler();

        }

        private void btnSil_Click(object sender, EventArgs e)
        {
            string filePath = Form1.filePath;
            // Seçili hücrenin satırını ve sütununu al

            int rowIndex = dataGridView1.CurrentCell.RowIndex + 1; // DataGridView sıfırdan değil 1'den başlıyor
            int columnIndex = dataGridView1.CurrentCell.ColumnIndex + 1; // DataGridView sıfırdan değil 1'den başlıyor

            // Excel dosyasını aç
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel._Worksheet worksheet = workbook.Sheets[1];

            // Hücreyi sil
            Excel.Range range = worksheet.Rows[rowIndex + 1];
            range.Delete();
            
            

            // Değişiklikleri kaydet ve Excel'i kapat
            workbook.Save();
            workbook.Close();
            MessageBox.Show("Makine Silindi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            Veriler();
            


        }

        private void btnListele_Click(object sender, EventArgs e)
        {
            Veriler();
        }
    }
}
