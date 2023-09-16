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
    public partial class frmDashboard : Form
    {
        public frmDashboard()
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
        private void frmDashboard_Load(object sender, EventArgs e)
        {
            Veriler();
        }
    }
}
