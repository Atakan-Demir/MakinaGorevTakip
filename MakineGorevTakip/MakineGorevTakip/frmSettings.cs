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

namespace MakineGorevTakip
{
    public partial class frmSettings : Form
    {
        public frmSettings()
        {
            InitializeComponent();
            
    }
        public static string DosyaYolu;
        string DosyaAdi;
        private void btnGozat_Click(object sender, EventArgs e)
        {
            
            
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Excel Dosyası | *.xls; *.xlsx; *.xlsm";
            if (file.ShowDialog() == DialogResult.OK)
            {
                DosyaYolu = file.FileName;// seçilen dosyanın tüm yolunu verir
                txtPath.Text = DosyaYolu;
                DosyaAdi = file.SafeFileName;// seçilen dosyanın adını verir.
                Excel.Application excelApp = new Excel.Application();
                if (excelApp == null)
                { //Excel Yüklümü Kontrolü Yapılmaktadır.
                    MessageBox.Show("Excel yüklü değil.");
                    return;
                }

                

            }
            else
            {
                MessageBox.Show("Dosya Seçilemedi.");
            }
        }

        private void btnTamam_Click(object sender, EventArgs e)
        {
            Form1.filePath = txtPath.Text;
            MessageBox.Show("Dosya Seçildi.");
        }

        private void frmSettings_Load(object sender, EventArgs e)
        {

        }
    }
}
