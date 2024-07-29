
using Excel = Microsoft.Office.Interop.Excel;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExcelOps
{
    public partial class LoadExcelAndManage : Form
    {
        Dictionary<string, int> dictSheets = new Dictionary<string, int>();
        string filename = "";
        string delimeter = "";
        string workbookPath = "";
        System.Data.DataTable dtexcel = new System.Data.DataTable();
        System.Data.DataTable finaldtexcel = new System.Data.DataTable();
        List<string> selectedColumns = new List<string>();
        System.Data.DataTable dt = new System.Data.DataTable();
        public static int val = 0;

        public LoadExcelAndManage()
        {
            InitializeComponent();
        }

        //private void button1_Click(object sender, EventArgs e)
        //{

        //}

        private void button1_Click_1(object sender, EventArgs e)
        {
            string filte_txt = "";
            string default_filter = "xlsx";
            delimeter = "";


            switch (comboBox2.Text)
            {
                case "Excel":
                    filte_txt = "excel files (*.xlsx)|*.xlsx";
                    default_filter = "xlsx";
                    break;
                case "Comma(,) Separated":
                case "Semicolon(;) Separated":
                case "Space() Separated":
                    default_filter = "csv";
                    filte_txt = "csv files (*.csv)|*.csv|structured txt files (*.txt)|*.txt";

                    delimeter = ",";
                    if (comboBox2.Text == "Semicolon(;) Separated")
                        delimeter = ";";
                    else if (comboBox2.Text == "Space() Separated")
                        delimeter = " ";
                    break;
            }

            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                Title = "Browse Target Excel",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = default_filter,
                Filter = filte_txt,
                //FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filename = openFileDialog1.FileName;
                ReadSheets(filename, delimeter);
                foreach (var single in dictSheets.Keys)
                {
                    listBox1.Items.Add(single);
                }
            }
        }

        private void LoadExcelAndManage_Load(object sender, EventArgs e)
        {

        }

        private void ReadSheets(string path, string delimeter)
        {
            workbookPath = path;

            try
            {
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false;  // Makes Excel visible to the user.           
                                           // The following code opens an existing workbook
                var fileformat = Excel.XlFileFormat.xlWK1;
                if (delimeter == "," || delimeter == ";" || delimeter == " ")
                    fileformat = Excel.XlFileFormat.xlCSV;

                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0,
                false, fileformat, "", "", false, Excel.XlPlatform.xlWindows, delimeter, true,
                false, 0, true, false, false);



                // The following gets the Worksheets collection
                Excel.Sheets excelSheets = excelWorkbook.Worksheets;

                foreach (Worksheet worksheet in excelWorkbook.Worksheets)
                {
                    dictSheets.Add(worksheet.Name, worksheet.Index);
                }

                excelWorkbook.Close();
                excelApp.Quit();
            }
            catch
            {
                //Create a new workbook if the existing workbook failed to open.
                //excelWorkbook = excelApp.Workbooks.Add();
                MessageBox.Show("Error in excel reading", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        public static System.Data.DataTable GetDataTableExcel(string strFileName, string Table)
        {
            System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + strFileName + "; Extended Properties = \"Excel 12.0;HDR=Yes\";");
            conn.Open();
            string strQuery = "SELECT * FROM [" + Table + "$]";
            System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter(strQuery, conn);
            System.Data.DataSet ds = new System.Data.DataSet();
            adapter.Fill(ds);
            return ds.Tables[0];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                //ReadExcel1(filename, listBox1.SelectedItem.ToString());
                System.Data.DataTable dt = GetDataTableExcel(filename, listBox1.SelectedItem.ToString());

                dataGridView1.DataSource = dt;
            }
        }
    }
}
