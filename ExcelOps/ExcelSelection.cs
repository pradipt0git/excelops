using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOps
{
    public partial class ExcelSelection : Form
    {
        public int selectedNo = 0;
        string selectedFileName = "";
        public ExcelSelection()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (selectedNo == 1)
            {
                static_repos.select1_end_row = textBox4.Text;
                static_repos.select1_start_row = textBox3.Text;
                static_repos.select1_end_col = textBox2.Text;
                static_repos.select1_start_col = textBox1.Text;
                static_repos.select1_excel_sheet = listBox1.SelectedItem.ToString();
                static_repos.select1_file_path = selectedFileName;
            }
            else if (selectedNo == 2)
            {
                static_repos.select2_end_row = textBox4.Text;
                static_repos.select2_start_row = textBox3.Text;
                static_repos.select2_end_col = textBox2.Text;
                static_repos.select2_start_col = textBox1.Text;
                static_repos.select2_excel_sheet = listBox1.SelectedItem.ToString();
                static_repos.select2_file_path = selectedFileName;
            }
                this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Loader l = new Loader();


            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx|All Files|*.*"; // Set filter for XLSX files
            openFileDialog.Title = "Select an XLSX File";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //l.Show();

                selectedFileName = openFileDialog.FileName;
                label6.Text = selectedFileName;

                //read the excel and populate sheet names
                Excel.Application excelApp = null;
                Excel.Workbook workbook = null;

                try
                {
                    excelApp = new Excel.Application();
                    excelApp.DisplayAlerts = false; // Suppress alerts

                    string filePath = selectedFileName; // Replace with the actual file path
                    workbook = excelApp.Workbooks.Open(filePath, ReadOnly: true);

                    List<string> sheetNames = new List<string>();
                    foreach (Excel.Worksheet sheet in workbook.Sheets)
                    {
                        listBox1.Items.Add(sheet.Name);
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("An error occurred: " + ex.Message);
                }
                finally
                {
                    // Close and release resources
                    if (workbook != null)
                    {
                        Marshal.ReleaseComObject(workbook);
                    }
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
            }
            //l.Close();
        }

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Loader l = new Loader();
            //l.Show();
            //read the excel and populate sheet names
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false; // Suppress alerts

                workbook = excelApp.Workbooks.Open(selectedFileName, ReadOnly: true);

                List<string> sheetNames = new List<string>();
                foreach (Excel.Worksheet sheet in workbook.Sheets)
                {
                    //listBox1.Items.Add(sheet.Name);
                    if (sheet.Name == listBox1.SelectedItem.ToString())
                    {
                        int cols = sheet.UsedRange.Columns.Count;
                        string colChar = GetExcelColumnName(cols);
                        textBox2.Text = colChar;

                        int rows = sheet.UsedRange.Rows.Count;
                        textBox4.Text = rows.ToString();

                        break;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            finally
            {
                // Close and release resources
                if (workbook != null)
                {
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }
            //l.Close();
        }
    }
}
