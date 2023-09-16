using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace MakineGorevTakip
{
    public partial class Form1 : Form
    {

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]

        private static extern IntPtr CreateRoundRectRgn
     (
          int nLeftRect,
          int nTopRect,
          int nRightRect,
          int nBottomRect,
          int nWidthEllipse,
         int nHeightEllipse

      );
        public static string filePath = "C:/Users/DmR/Desktop/db.xlsx";
        public Form1()
        {
            InitializeComponent();
            
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 25, 25));
            pnlNav.Height = btnDashboard.Height;
            pnlNav.Top = btnDashboard.Top;
            pnlNav.Left = btnDashboard.Left;
            btnDashboard.BackColor = Color.FromArgb(46, 51, 73);

            lblTitle.Text = "Anasayfa";
            this.PnlFormLoader.Controls.Clear();
            frmDashboard FrmDashboard_Vrb = new frmDashboard() { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
            FrmDashboard_Vrb.FormBorderStyle = FormBorderStyle.None;
            this.PnlFormLoader.Controls.Add(FrmDashboard_Vrb);
            FrmDashboard_Vrb.Show();
            
        }
        

        private void Form1_Load(object sender, EventArgs e)
        {
            if (frmSettings.DosyaYolu != "")
            {
                filePath = frmSettings.DosyaYolu;
                Veriler();
                

            }
            else
            {
                MessageBox.Show("Lütfen ayarlar sekmesine giderek excel (.xlsx) çalışma tablosunun yolunu seçin.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        void Veriler()
        {
            try
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
            catch (Exception)
            {
                MessageBox.Show("Lütfen ayarları Kontrol EDiniz","Bir Sorun Yaşandı!",MessageBoxButtons.OK,MessageBoxIcon.Error);

                BtnSettings_Click(this,new EventArgs());
            }
            


        }
        private void BtnDashboard_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnDashboard.Height;
            pnlNav.Top = btnDashboard.Top;
            pnlNav.Left = btnDashboard.Left;
            btnDashboard.BackColor = Color.FromArgb(46, 51, 73);

            lblTitle.Text = "Anasayfa";
            this.PnlFormLoader.Controls.Clear();
            frmDashboard FrmDashboard_Vrb = new frmDashboard() { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
            FrmDashboard_Vrb.FormBorderStyle = FormBorderStyle.None;
            this.PnlFormLoader.Controls.Add(FrmDashboard_Vrb);
            FrmDashboard_Vrb.Show();
        }

        private void BtnAnalytics_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnAnalytics.Height;
            pnlNav.Top = btnAnalytics.Top;
            btnAnalytics.BackColor = Color.FromArgb(46, 51, 73);

            lblTitle.Text = "Makineler";
            this.PnlFormLoader.Controls.Clear();
            frmAnalytics FrmDashboard_Vrb = new frmAnalytics() { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
            FrmDashboard_Vrb.FormBorderStyle = FormBorderStyle.None;
            this.PnlFormLoader.Controls.Add(FrmDashboard_Vrb);
            FrmDashboard_Vrb.Show();
        }

        private void BtnCalander_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnCalander.Height;
            pnlNav.Top = btnCalander.Top;
            btnCalander.BackColor = Color.FromArgb(46, 51, 73);

            lblTitle.Text = "Görevler";
            this.PnlFormLoader.Controls.Clear();
            frmCalander FrmDashboard_Vrb = new frmCalander() { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
            FrmDashboard_Vrb.FormBorderStyle = FormBorderStyle.None;
            this.PnlFormLoader.Controls.Add(FrmDashboard_Vrb);
            FrmDashboard_Vrb.Show();
        }

        

        

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnSettings.Height;
            pnlNav.Top = btnSettings.Top;
            btnSettings.BackColor = Color.FromArgb(46, 51, 73);

            lblTitle.Text = "Ayarlar";
            this.PnlFormLoader.Controls.Clear();
            frmSettings FrmDashboard_Vrb = new frmSettings() { Dock = DockStyle.Fill, TopLevel = false, TopMost = true };
            FrmDashboard_Vrb.FormBorderStyle = FormBorderStyle.None;
            this.PnlFormLoader.Controls.Add(FrmDashboard_Vrb);
            FrmDashboard_Vrb.Show();
        }

        private void BtnDashboard_Leave(object sender, EventArgs e)
        {
            btnDashboard.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void BtnAnalytics_Leave(object sender, EventArgs e)
        {
            btnAnalytics.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void BtnCalander_Leave(object sender, EventArgs e)
        {
            btnCalander.BackColor = Color.FromArgb(24, 30, 54);
        }

        

       

        private void BtnSettings_Leave(object sender, EventArgs e)
        {
            btnSettings.BackColor = Color.FromArgb(24, 30, 54);
        }

        


        private void Button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

       

        private void Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void PnlFormLoader_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
