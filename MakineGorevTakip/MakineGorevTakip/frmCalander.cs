using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace MakineGorevTakip
{
    public partial class frmCalander : Form
    {
        public frmCalander()
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
            
            for (int i = 2; i <= worksheet.UsedRange.Rows.Count; i++)
            {
                string makineAdi = worksheet.Cells[i, 1].Value;
                comboBox1.Items.Add(makineAdi);
            }
            
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
        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void PictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void frmCalander_Load(object sender, EventArgs e)
        {

            Veriler();
        }

        private void btnGuncelle_Click(object sender, EventArgs e)
        {
            // Seçilen makine adını al
            string selectedMachine = comboBox1.SelectedItem.ToString();
            
            // dataGridView'deki tüm satırları döngüye alarak, seçilen makine adına göre filtrele
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[0].Value.ToString() == selectedMachine)
                {
                    if (comboBox2.SelectedIndex == comboBox2.Items.Count - 1)
                    {
                        textBox1.Text = "***Görev Atanmadı***";
                    }
                    // Makinenin olduğu satırdaki görev ve görev durumunu güncelle
                    row.Cells[1].Value = textBox1.Text;
                    row.Cells[2].Value = comboBox2.Text;

                    // Excel dosyasını aç
                    string filePath = Form1.filePath;
                    Excel.Application excel = new Excel.Application();
                    Excel.Workbook workbook = excel.Workbooks.Open(filePath); // Excel dosyanızın yolunu belirtin

                    // Excel dosyasındaki ilgili hücreleri güncelle
                    Excel.Worksheet worksheet = workbook.Sheets[1]; // İlk sayfa
                    worksheet.Cells[row.Index + 2, 2] = textBox1.Text; // Görev hücresi
                    worksheet.Cells[row.Index + 2, 3] = comboBox2.Text; // Görev durumu hücresi

                    // Excel dosyasını kaydet ve kapat
                    workbook.Save();
                    workbook.Close();
                    excel.Quit();

                    break;
                }
            }
            Veriler();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            // DataGridView'de bir satır seçildiğinde tetiklenen olay

            // Seçilen satırın indeksini al
            int rowIndex = dataGridView1.CurrentCell.RowIndex;

            // Seçilen satırın hücrelerindeki verileri al
            string column1Value = dataGridView1.Rows[rowIndex].Cells[0].Value.ToString();
            string column2Value = dataGridView1.Rows[rowIndex].Cells[1].Value.ToString();
            string column3Value = dataGridView1.Rows[rowIndex].Cells[2].Value.ToString();

            // ComboBox1'in içeriğini güncelle
            comboBox1.Text = column1Value;

            // TextBox1'in içeriğini güncelle
            textBox1.Text = column2Value;

            // ComboBox2'in içeriğini güncelle
            comboBox2.Text = column3Value;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Seçilen makine adını al
            string selectedMachine = comboBox1.SelectedItem.ToString();

            // dataGridView'deki tüm satırları döngüye alarak, seçilen makine adına göre filtrele
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[0].Value.ToString() == selectedMachine)
                {
                    // Makinenin olduğu satırdaki verileri al ve textBox1 ve comboBox2'ye ata
                    textBox1.Text = row.Cells[1].Value.ToString();
                    comboBox2.Text = row.Cells[2].Value.ToString();
                    break;
                }
            }
        }

        private void btnSil_Click(object sender, EventArgs e)
        {
            comboBox2.SelectedIndex = comboBox2.Items.Count - 1;
            textBox1.Text = "** Görev Atanmamış **";
            // Seçilen makine adını al
            string selectedMachine = comboBox1.SelectedItem.ToString();

            // dataGridView'deki tüm satırları döngüye alarak, seçilen makine adına göre filtrele
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[0].Value.ToString() == selectedMachine)
                {
                    // Makinenin olduğu satırdaki görev ve görev durumunu güncelle
                    row.Cells[1].Value = textBox1.Text;
                    row.Cells[2].Value = comboBox2.Text;

                    // Excel dosyasını aç
                    string filePath = Form1.filePath;
                    Excel.Application excel = new Excel.Application();
                    Excel.Workbook workbook = excel.Workbooks.Open(filePath); // Excel dosyanızın yolunu belirtin

                    // Excel dosyasındaki ilgili hücreleri güncelle
                    Excel.Worksheet worksheet = workbook.Sheets[1]; // İlk sayfa
                    worksheet.Cells[row.Index + 2, 2] = textBox1.Text; // Görev hücresi
                    worksheet.Cells[row.Index + 2, 3] = comboBox2.Text; // Görev durumu hücresi

                    // Excel dosyasını kaydet ve kapat
                    workbook.Save();
                    workbook.Close();
                    excel.Quit();

                    break;
                }
            }

        }
    }
}
