//using Microsoft.Office.Interop.Excel;
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
using ExcelDataReader;
using System.IO;

namespace ExcelOps
{
    public partial class DataManage : Form
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
        public DataManage()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
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
                //textBox1.Text = openFileDialog1.FileName;
            }
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

        private void button2_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                //ReadExcel1(filename, listBox1.SelectedItem.ToString());
                ReadExcel(filename, listBox1.SelectedItem.ToString());

                //dataGridView1.DataSource = dtexcel;
            }
        }

        private void ReadExcel(string filename, string sheetName)
        {
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
                Worksheet ws = (Worksheet)excelWorkbook.Worksheets[dictSheets[sheetName]];

                if (ws != null)
                {

                    int cols = ws.UsedRange.Columns.Count;
                    int rows = ws.UsedRange.Rows.Count;
                    int noofrow = 1;
                    int headerrowNo = 1;

                    if (numericUpDown1.Value > 0)
                        headerrowNo = Convert.ToInt32(numericUpDown1.Value);
                    if (numericUpDown2.Value > 0)
                        rows = Convert.ToInt32(numericUpDown2.Value);

                    for (int c = 1; c <= cols; c++)
                    {
                        //get column names from header row
                        string colname = ws.Cells[headerrowNo, c].Text;
                        selectedColumns.Add(colname);
                        chklst1.Items.Add(colname);

                        checkedListBox1.Items.Add(colname);
                        //checkedListBox2.Items.Add(colname);

                        dt.Columns.Add(colname);
                        finaldtexcel.Columns.Add(colname);
                    }
                    noofrow = headerrowNo + 1;

                    Excel.Range c1 = ws.Cells[noofrow, 1];
                    Excel.Range c2 = ws.Cells[rows, cols];

                    Excel.Range oRange = (Excel.Range)ws.get_Range(c1, c2);

                    int totalNoOfRows = rows - noofrow;
                    int noofItemsPerLoop = 50;
                    int loopCount = totalNoOfRows / noofItemsPerLoop;


                    if (loopCount == 0)
                    {
                        loopCount = 1;
                    }

                    int startingfrom = noofrow;
                    int endingRow = noofItemsPerLoop;

                    //for each 50 rows call this loop once
                    int loopStartFrom = 0;
                    int loopEndJustBefore = loopCount;
                    //{
                    //for (int j = 0; j < loopCount; j++)
                    Parallel.For(loopStartFrom, loopEndJustBefore, async j =>
                    {
                        if (endingRow >= totalNoOfRows)
                        {
                            endingRow = totalNoOfRows;
                        }

                        Excel.Range c3 = ws.Cells[startingfrom, 1];
                        Excel.Range c4 = ws.Cells[endingRow, cols];

                        Excel.Range childRange = (Excel.Range)ws.get_Range(c3, c4);

                        //for each 50 records a new thread will initiate
                        //Thread myNewThread = new Thread(() => ReadSections(childRange, cols));
                        //myNewThread.Start();

                        await ReadSections(childRange, cols);
                        //Console.WriteLine(j.ToString());

                        startingfrom = startingfrom + noofItemsPerLoop;

                        endingRow = startingfrom + noofItemsPerLoop;
                        //}
                    });

                    dtexcel = dt;
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

        private async
        Task
ReadSections(Excel.Range childRange, int cols)
        {
            try
            {
                int loopStartFrom = 1;
                int loopEndJustBefore = 50 + 1;
                if (childRange.Rows.Count <= loopEndJustBefore)
                    loopEndJustBefore = childRange.Rows.Count;

                //Parallel.For(loopStartFrom, loopEndJustBefore, i =>
                for (int i = 1; i < loopEndJustBefore; i++)
                {
                    DataRow dr = dt.NewRow();
                    //Excel.Range r = oRange.Rows[i];

                    for (int c = 1; c <= cols; c++)
                    {
                        dr[c - 1] = (childRange.Cells[i, c]).Text;
                    }
                    dt.Rows.Add(dr);
                    val++;

                }
                //});
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void DataManage_Load(object sender, EventArgs e)
        {

        }
    }
}
