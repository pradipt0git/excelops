using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace ExcelOps
{
    public partial class EasyCompare : Form
    {
        public EasyCompare()
        {
            InitializeComponent();

        }

        private void EasyCompare_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelSelection es = new ExcelSelection();
            es.selectedNo = 1;
            es.ShowDialog();

            lbl_selection1.Text = static_repos.select1_file_path + "\n" + static_repos.select1_excel_sheet + "\n" + static_repos.select1_start_col + static_repos.select1_start_row + "-" + static_repos.select1_end_col + static_repos.select1_end_row;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelSelection es = new ExcelSelection();
            es.selectedNo = 2;
            es.ShowDialog();
            lbl_selection2.Text = static_repos.select2_file_path + "\n" + static_repos.select2_excel_sheet + "\n" + static_repos.select2_start_col + static_repos.select2_start_row + "-" + static_repos.select2_end_col + static_repos.select2_end_row;
        }

        private void btnCompare_Click(object sender, EventArgs e)
        {
            DataTable dt1 = ReadExcel(static_repos.select1_file_path,".xlsx",static_repos.select1_excel_sheet, static_repos.select1_start_col + static_repos.select1_start_row, static_repos.select1_end_col + static_repos.select1_end_row);
            DataTable dt2 = ReadExcel(static_repos.select2_file_path, ".xlsx", static_repos.select2_excel_sheet, static_repos.select2_start_col + static_repos.select2_start_row, static_repos.select2_end_col + static_repos.select2_end_row);

        }

        public DataTable ReadExcel(string fileName, string fileExt, string excel_sheet, string startLocation,string endLocation)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=NO;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from ["+ excel_sheet + "$"+startLocation+":"+endLocation+"]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch { }
            }
            return dtexcel;
        }
    }
}
