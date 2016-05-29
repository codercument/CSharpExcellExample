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


namespace MyExcel
{
    public partial class Form1 : Form
    {
        string dosya = Application.StartupPath + "\\" + "rapor.xlsx";
        Excel.Application ExcelApp = new Excel.Application();

        Excel.Range rng;
        Excel.Range rngnum;
        Excel.Range rngdate;
        Excel.Range rnguser;
        Excel.Range rngres;
        Excel.Range rngid;
        Excel.Range rnginfo;

        int count =1;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            


            ExcelApp.Visible = true;

        }

        private void button1_Click(object sender, EventArgs e)
        {

            ExcelApp.Visible = false;
            ExcelApp.Workbooks.Open(dosya);
            ExcelApp.Worksheets[1].Activate();
            Excel.Worksheet sheet = ExcelApp.ActiveSheet;
            rng = (Excel.Range)sheet.get_Range("Testler", Type.Missing);
            rngnum = (Excel.Range)sheet.get_Range("NUMBER", Type.Missing);
            rngdate = (Excel.Range)sheet.get_Range("DATE", Type.Missing);
            rnguser = (Excel.Range)sheet.get_Range("USER", Type.Missing);
            rngres = (Excel.Range)sheet.get_Range("RESULTS", Type.Missing);
            rngid = (Excel.Range)sheet.get_Range("PRODUCT_ID", Type.Missing);
            rnginfo = (Excel.Range)sheet.get_Range("Device_Info", Type.Missing);

            for (count =1; count <= rngnum.Count; count++)
            {

                rngnum.Rows[count] = count;
                
            }

            ExcelApp.Visible = true;




        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            
           
            ExcelApp.Quit();
            ExcelApp = null;
            GC.Collect();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //ExcelApp.Workbooks.Close();
            // ExcelApp.Application.Quit();
            ExcelApp.GetSaveAsFilename();
          
        }
    }
}
