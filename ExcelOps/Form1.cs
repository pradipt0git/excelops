
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;
using ExcelDataReader;

namespace ExcelOps
{
    public partial class Form1 : Form
    {
        Dictionary<string, Worksheet> dict = new Dictionary<string, Worksheet>();
        Excel.Application excelApp = new Excel.Application();  // Creates a new Excel Application
        Excel.Workbook excelWorkbook = null;
        string filename = "";
        List<string> selectedColumns = new List<string>();

        System.Data.DataTable dtexcel = new System.Data.DataTable();
        System.Data.DataTable finaldtexcel = new System.Data.DataTable();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            resetGridAndData(1);
            string filte_txt = "";
            string default_filter = "xlsx";
            string delimeter = "";


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

                foreach (var single in dict.Keys)
                {
                    listBox1.Items.Add(single);
                }
                //textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void ReadSheets(string path, string delimeter)
        {
            excelApp.Visible = false;  // Makes Excel visible to the user.           
                                       // The following code opens an existing workbook
            string workbookPath = path;

            try
            {
                var fileformat = Excel.XlFileFormat.xlWK1;
                if (delimeter == "," || delimeter == ";" || delimeter == " ")
                    fileformat = Excel.XlFileFormat.xlCSV;

                excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0,
                true, fileformat, "", "", false, Excel.XlPlatform.xlWindows, delimeter, false,
                false, 0, true, false, false);
            }
            catch
            {
                //Create a new workbook if the existing workbook failed to open.
                //excelWorkbook = excelApp.Workbooks.Add();
                MessageBox.Show("Error in excel reading", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            // The following gets the Worksheets collection            

            foreach (Worksheet worksheet in excelWorkbook.Worksheets)
            {
                dict.Add(worksheet.Name, worksheet);
            }

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            resetGridAndData(2);
            txtHeaderRow.Text = "";
        }

        private void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                if (excelWorkbook != null)
                {
                    excelWorkbook.Close();
                    excelApp.Quit();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            resetGridAndData(3);
            if (listBox1.SelectedItem != null)
            {
                //ReadExcel1(filename, listBox1.SelectedItem.ToString());
                GetDataTableExcel(filename, listBox1.SelectedItem.ToString());

                dataGridView1.DataSource = dtexcel;
                //as dataset is ready now we can remove excel object

            }
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

        public void GetDataTableExcel(string strFileName, string Table)
        {
            Worksheet ws = dict[Table];
            if (ws != null)
            {

                int cols = ws.UsedRange.Columns.Count;
                string colChar = GetExcelColumnName(cols);
                int rows = ws.UsedRange.Rows.Count;
                System.Data.DataTable dt = new System.Data.DataTable();
                int noofrow = 1;
                int headerrowNo = 1;

                if (!string.IsNullOrEmpty(txtHeaderRow.Text))
                    headerrowNo = Convert.ToInt32(txtHeaderRow.Text);
                if (!string.IsNullOrEmpty(txtLastRow.Text))
                    rows = Convert.ToInt32(txtLastRow.Text);

                for (int c = 1; c <= cols; c++)
                {
                    //get column names from header row
                    string colname = ws.Cells[headerrowNo, c].Text;
                    selectedColumns.Add(colname);
                    chklst1.Items.Add(colname);

                    checkedListBox1.Items.Add(colname);
                    checkedListBox2.Items.Add(colname);

                    dt.Columns.Add(colname);
                    finaldtexcel.Columns.Add(colname);
                }

                using (System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + strFileName + "; Extended Properties = \"Excel 12.0;HDR=Yes\";"))
                {
                    conn.Open();
                    string strQuery = "SELECT * FROM [" + Table + "$A" + headerrowNo + ":" + colChar + "" + rows + "]";
                    using (System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter(strQuery, conn))
                    {
                        System.Data.DataSet ds = new System.Data.DataSet();
                        adapter.Fill(ds);
                        dtexcel = ds.Tables[0];
                    }
                }


            }

        }

        private void ReadExcel(string fileName, string sheetName)
        {

            Worksheet ws = dict[sheetName];
            if (ws != null)
            {

                int cols = ws.UsedRange.Columns.Count;
                int rows = ws.UsedRange.Rows.Count;
                System.Data.DataTable dt = new System.Data.DataTable();
                int noofrow = 1;
                int headerrowNo = 1;

                if (!string.IsNullOrEmpty(txtHeaderRow.Text))
                    headerrowNo = Convert.ToInt32(txtHeaderRow.Text);
                if (!string.IsNullOrEmpty(txtLastRow.Text))
                    rows = Convert.ToInt32(txtLastRow.Text);

                for (int c = 1; c <= cols; c++)
                {
                    //get column names from header row
                    string colname = ws.Cells[headerrowNo, c].Text;
                    selectedColumns.Add(colname);
                    chklst1.Items.Add(colname);

                    checkedListBox1.Items.Add(colname);
                    checkedListBox2.Items.Add(colname);

                    dt.Columns.Add(colname);
                    finaldtexcel.Columns.Add(colname);
                }
                noofrow = headerrowNo + 1;

                for (int r = noofrow; r <= rows; r++)
                {
                    DataRow dr = dt.NewRow();
                    for (int c = 1; c <= cols; c++)
                    {
                        dr[c - 1] = ws.Cells[r, c].Text;
                    }

                    dt.Rows.Add(dr);
                }

                dtexcel = dt;
            }


        }

        Dictionary<string, DataRow> dtUniqueCollection = new Dictionary<string, DataRow>();
        Dictionary<string, int> unique_count = new Dictionary<string, int>();


        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                resetGridAndData(4);
                System.Data.DataTable dtexcel_clone = dtexcel.Copy();

                //dtexcel.Columns.Add("unique");
                List<string> finalUniqueColumnSet = getSelectedChekboxList(checkedListBox1);

                for (int i = 0; i < dtexcel_clone.Rows.Count; i++)
                {
                    string uniqueVal = "";
                    foreach (string selected_colname in finalUniqueColumnSet)
                    {
                        uniqueVal += (Convert.ToString(dtexcel_clone.Rows[i][selected_colname]) ?? "") + "|";
                    }

                    //check if the unique value already existin the lst

                    if (dtUniqueCollection.ContainsKey(uniqueVal) == true)
                    {
                        int existing_count = unique_count[uniqueVal];
                        existing_count++;
                        unique_count[uniqueVal] = existing_count;

                        if (chklst1.SelectedItem != null)
                        {
                            if (comboBox1.SelectedItem == "Sum")
                            {
                                DataRow drnew = dtexcel_clone.NewRow();
                                drnew = dtUniqueCollection[uniqueVal];
                                string aggregatedColName = chklst1.SelectedItem.ToString() ?? "";

                                var val1 = getValue(drnew[aggregatedColName]);
                                var val2 = getValue(dtexcel_clone.Rows[i][aggregatedColName]);
                                val1 += val2;
                                drnew[aggregatedColName] = val1.ToString();

                                //replace existing row with the new row 
                                dtUniqueCollection[uniqueVal] = drnew;
                            }
                            else if (comboBox1.SelectedItem == "Max")
                            {
                                DataRow drnew = dtexcel_clone.NewRow();
                                drnew = dtUniqueCollection[uniqueVal];
                                string aggregatedColName = chklst1.SelectedItem.ToString() ?? "";
                                var val1 = getValue(drnew[aggregatedColName]);
                                var val2 = getValue(dtexcel_clone.Rows[i][aggregatedColName]);

                                drnew[aggregatedColName] = Math.Max(val1, val2).ToString();
                                //replace existing row with the new row 
                                dtUniqueCollection[uniqueVal] = drnew;

                            }
                            else if (comboBox1.SelectedItem == "Min")
                            {
                                DataRow drnew = dtexcel_clone.NewRow();
                                drnew = dtUniqueCollection[uniqueVal];
                                string aggregatedColName = chklst1.SelectedItem.ToString() ?? "";
                                var val1 = getValue(drnew[aggregatedColName]);
                                var val2 = getValue(dtexcel_clone.Rows[i][aggregatedColName]);

                                drnew[aggregatedColName] = Math.Min(val1, val2).ToString(); ;
                                //replace existing row with the new row 
                                dtUniqueCollection[uniqueVal] = drnew;

                            }
                            else if (comboBox1.SelectedItem == "Count")
                            {
                                DataRow drnew = dtexcel_clone.NewRow();
                                drnew = dtUniqueCollection[uniqueVal];

                                string aggregatedColName = chklst1.SelectedItem.ToString() ?? "";
                                drnew[aggregatedColName] = existing_count.ToString();
                                dtUniqueCollection[uniqueVal] = drnew;

                            }

                        }
                    }
                    else
                    {
                        unique_count.Add(uniqueVal, 1);

                        //for count set count as 1 for non matching unique records
                        if (chklst1.SelectedItem != null && comboBox1.SelectedItem == "Count")
                        {
                            DataRow drnew = dtexcel_clone.Rows[i];

                            string aggregatedColName = chklst1.SelectedItem.ToString() ?? "";
                            drnew[aggregatedColName] = 1.ToString();
                            dtUniqueCollection[uniqueVal] = drnew;
                        }
                        else
                        {
                            dtUniqueCollection.Add(uniqueVal, dtexcel_clone.Rows[i]);
                        }
                    }
                }

                if (dtUniqueCollection.Count > 0)
                {
                    foreach (string single in dtUniqueCollection.Keys)
                    {
                        finaldtexcel.Rows.Add(dtUniqueCollection[single].ItemArray);
                    }
                }

                dataGridView1.DataSource = finaldtexcel;
                hideColumns(dataGridView1);
            }
            catch (Exception ex)
            {

            }
        }

        private void hideColumns(DataGridView dataGridView1)
        {
            if (chkSelectedColumns.Checked)
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    if (!checkedListBox1.GetItemChecked(i))
                    {
                        string str = (string)checkedListBox1.Items[i];
                        dataGridView1.Columns[str].Visible = false;
                    }
                }
                for (int i = 0; i < chklst1.Items.Count; i++)
                {
                    if (chklst1.GetItemChecked(i))
                    {
                        string str = (string)chklst1.Items[i];
                        dataGridView1.Columns[str].Visible = true;
                    }
                }
            }
        }

        private dynamic getValue(object v)
        {
            int x = 0;
            float x1 = 0;
            double x2 = 0;
            DateTime x3 = new DateTime();
            if (Int32.TryParse(Convert.ToString(v), out x))
            {
                return x;
            }
            if (float.TryParse(Convert.ToString(v), out x1))
            {
                return x1;
            }
            if (double.TryParse(Convert.ToString(v), out x2))
            {
                return x2;
            }
            if (DateTime.TryParse(Convert.ToString(v), out x3))
            {
                return x3;
            }
            return Convert.ToString(v);
        }

        private void chklst1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = chklst1.SelectedIndex;
            int count = chklst1.Items.Count;
            for (int i = 0; i < count; i++)
            {
                if (i != index)
                {
                    chklst1.SetItemCheckState(i, CheckState.Unchecked);
                }
            }
        }

        private List<string> getSelectedChekboxList(CheckedListBox passedChkbxList)
        {
            List<string> finalUnique = new List<string>();

            for (int i = 0; i < passedChkbxList.Items.Count; i++)
            {
                if (passedChkbxList.GetItemChecked(i))
                {
                    string str = (string)passedChkbxList.Items[i];
                    finalUnique.Add(str);
                }
            }
            return finalUnique;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            resetGridAndData(5);
            System.Data.DataTable dtexcel_clone = dtexcel.Copy();

            //dtexcel.Columns.Add("unique");
            List<string> finalUniqueColumnSet = getSelectedChekboxList(checkedListBox2);

            for (int i = 0; i < dtexcel_clone.Rows.Count; i++)
            {
                string uniqueVal = "";
                foreach (string selected_colname in finalUniqueColumnSet)
                {
                    uniqueVal += (Convert.ToString(dtexcel_clone.Rows[i][selected_colname]) ?? "") + "|";
                }

                //check if the unique value already existin the lst

                if (dtUniqueCollection.ContainsKey(uniqueVal) == true)
                {

                }
                else
                {
                    dtUniqueCollection.Add(uniqueVal, dtexcel_clone.Rows[i]);
                }
            }

            foreach (string single in dtUniqueCollection.Keys)
            {
                finaldtexcel.Rows.Add(dtUniqueCollection[single].ItemArray);
            }

            dataGridView1.DataSource = finaldtexcel;
        }

        private void resetGridAndData(int stageno)
        {
            switch (stageno)
            {
                case 1:
                    //full reset
                    listBox1.Items.Clear();
                    txtHeaderRow.Text = "";
                    txtLastRow.Text = "";

                    checkedListBox1.Items.Clear();
                    chklst1.Items.Clear();

                    checkedListBox2.Items.Clear();
                    dataGridView1.DataSource = new System.Data.DataTable();
                    dtUniqueCollection = new Dictionary<string, DataRow>();
                    unique_count = new Dictionary<string, int>();
                    finaldtexcel.Clear();
                    finaldtexcel = new System.Data.DataTable();

                    dict.Clear();

                    break;
                case 2:
                    //excel will be selected 
                    //sheet should be selected
                    //but rest will be reset

                    txtHeaderRow.Text = "";
                    txtLastRow.Text = "";

                    checkedListBox1.Items.Clear();
                    chklst1.Items.Clear();

                    checkedListBox2.Items.Clear();
                    dataGridView1.DataSource = new System.Data.DataTable();
                    dtUniqueCollection = new Dictionary<string, DataRow>();
                    unique_count = new Dictionary<string, int>();
                    finaldtexcel.Clear();
                    finaldtexcel = new System.Data.DataTable();

                    //dict.Clear();
                    break;
                case 3:
                    //excel will be selected 
                    //sheet should be selected
                    //row header and last row will be selected
                    //but rest will be reset

                    checkedListBox1.Items.Clear();
                    chklst1.Items.Clear();

                    checkedListBox2.Items.Clear();
                    dataGridView1.DataSource = new System.Data.DataTable();
                    dtUniqueCollection = new Dictionary<string, DataRow>();
                    unique_count = new Dictionary<string, int>();
                    finaldtexcel.Clear();
                    finaldtexcel = new System.Data.DataTable();

                    //dict.Clear();
                    break;
                case 4:
                    //excel will be selected 
                    //sheet should be selected
                    //row header and last row will be selected
                    //column checkbox and aggreegated checkbox will remain same 
                    //remove data from finaldtexcel
                    //but rest will be reset

                    clearSelection(checkedListBox2);

                    dataGridView1.DataSource = new System.Data.DataTable();
                    dtUniqueCollection = new Dictionary<string, DataRow>();
                    unique_count = new Dictionary<string, int>();
                    finaldtexcel.Clear();
                    //finaldtexcel = new System.Data.DataTable();


                    //dict.Clear();
                    break;
                case 5:
                    //excel will be selected 
                    //sheet should be selected
                    //row header and last row will be selected
                    //column checkbox and aggreegated checkbox selection will gone
                    //remove data from finaldtexcel
                    //but rest will be reset

                    clearSelection(checkedListBox1);
                    clearSelection(chklst1);


                    dataGridView1.DataSource = new System.Data.DataTable();
                    dtUniqueCollection = new Dictionary<string, DataRow>();
                    unique_count = new Dictionary<string, int>();
                    finaldtexcel.Clear();
                    //finaldtexcel = new System.Data.DataTable();

                    //dict.Clear();
                    break;
            }

        }

        private void clearSelection(CheckedListBox checkedListBoxPassed)
        {
            int count = checkedListBoxPassed.Items.Count;
            for (int i = 0; i < count; i++)
            {
                checkedListBoxPassed.SetItemCheckState(i, CheckState.Unchecked);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            dataGridView1.Width = this.Width;
            dataGridView1.Height = this.Height - 250;

            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            System.Data.DataTable copy_finaldtexcel = new System.Data.DataTable();
            //get the datasource of the gridview
            //then apply filter based on written query
            if (finaldtexcel != null && finaldtexcel.Rows.Count > 0)
            {
                copy_finaldtexcel = finaldtexcel.Copy();
            }
            else
            {
                copy_finaldtexcel = dtexcel.Copy();
            }

            if (copy_finaldtexcel != null && copy_finaldtexcel.Rows.Count > 0)
            {
                //var copy_finaldtexcel = ((System.Data.DataTable)dataGridView1.DataSource).Copy();
                string filterExpression = txtSearch.Text;
                copy_finaldtexcel.DefaultView.RowFilter = filterExpression;
                dataGridView1.DataSource = copy_finaldtexcel.DefaultView;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                SaveDataGridViewToCSV(saveFileDialog1.FileName);
            }
        }

        void SaveDataGridViewToCSV(string filename)
        {
            try
            {
                // Choose whether to write header. Use EnableWithoutHeaderText instead to omit header.
                dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                // Select all the cells
                dataGridView1.SelectAll();
                // Copy selected cells to DataObject
                DataObject dataObject = dataGridView1.GetClipboardContent();
                // Get the text of the DataObject, and serialize it to a file
                File.WriteAllText(filename, dataObject.GetText(TextDataFormat.CommaSeparatedValue));
            }
            catch (Exception ex)
            {

            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void easyAggregationToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void easyCompareToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EasyCompare ec = new EasyCompare();
            ec.Show();
        }

        //private System.Data.DataTable ReadExcel1(string fileName, string fileExt)
        //{
        //    string conn = string.Empty;
        //    System.Data.DataTable dtexcel = new System.Data.DataTable();
        //    if (fileExt.CompareTo(".xls") == 0) conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007
        //    else conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007
        //    using (OleDbConnection con = new System.Data.OleDb.OleDbConnection(conn))
        //    {
        //        try
        //        {
        //            OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1
        //            oleAdpt.Fill(dtexcel); //fill excel data into dataTable
        //        }
        //        catch { }
        //    }
        //    return dtexcel;
        //}

        //private string GetNumberOfRows(string filename, string sheetName)
        //{
        //    string connectionString = "";
        //    string count = "";

        //    if (filename.EndsWith(".xlsx"))
        //    {
        //        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Mode=ReadWrite;Extended Properties=\"Excel 12.0;HDR=NO\"";
        //    }
        //    else if (filename.EndsWith(".xls"))
        //    {
        //        connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Mode=ReadWrite;Extended Properties=\"Excel 8.0;HDR=NO;\"";
        //    }

        //    string SQL = "SELECT COUNT (*) FROM [" + sheetName + "$]";

        //    using (OleDbConnection conn = new OleDbConnection(connectionString))
        //    {
        //        conn.Open();

        //        try
        //        {
        //            OleDbDataAdapter oleAdpt = new OleDbDataAdapter(SQL, conn); //here we read data from sheet1
        //            oleAdpt.Fill(dtexcel); //fill excel data into dataTable
        //        }
        //        catch { }


        //        conn.Close();
        //    }

        //    return count;
        //}

    }
}