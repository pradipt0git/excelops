using System.Runtime.InteropServices;

namespace ExcelOps
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            button1 = new Button();
            listBox1 = new ListBox();
            openFileDialog1 = new OpenFileDialog();
            txtHeaderRow = new TextBox();
            label1 = new Label();
            button2 = new Button();
            chklst1 = new CheckedListBox();
            label2 = new Label();
            txtLastRow = new TextBox();
            button3 = new Button();
            dataGridView1 = new DataGridView();
            checkedListBox1 = new CheckedListBox();
            label4 = new Label();
            checkedListBox2 = new CheckedListBox();
            button4 = new Button();
            groupBox1 = new GroupBox();
            chkSelectedColumns = new CheckBox();
            comboBox1 = new ComboBox();
            groupBox2 = new GroupBox();
            comboBox2 = new ComboBox();
            txtSearch = new RichTextBox();
            button5 = new Button();
            label6 = new Label();
            label7 = new Label();
            groupBox4 = new GroupBox();
            label5 = new Label();
            groupBox3 = new GroupBox();
            button6 = new Button();
            saveFileDialog1 = new SaveFileDialog();
            groupBox5 = new GroupBox();
            groupBox6 = new GroupBox();
            menuStrip1 = new MenuStrip();
            easyAggregationToolStripMenuItem = new ToolStripMenuItem();
            easyCompareToolStripMenuItem = new ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            groupBox4.SuspendLayout();
            groupBox3.SuspendLayout();
            groupBox5.SuspendLayout();
            groupBox6.SuspendLayout();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // button1
            // 
            button1.BackColor = Color.CornflowerBlue;
            button1.Location = new Point(24, 118);
            button1.Name = "button1";
            button1.Size = new Size(101, 51);
            button1.TabIndex = 1;
            button1.Text = "Browse";
            button1.UseVisualStyleBackColor = false;
            button1.Click += button1_Click;
            // 
            // listBox1
            // 
            listBox1.FormattingEnabled = true;
            listBox1.ItemHeight = 20;
            listBox1.Location = new Point(3, 53);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(150, 104);
            listBox1.TabIndex = 2;
            listBox1.SelectedIndexChanged += listBox1_SelectedIndexChanged;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // txtHeaderRow
            // 
            txtHeaderRow.Location = new Point(159, 65);
            txtHeaderRow.Name = "txtHeaderRow";
            txtHeaderRow.Size = new Size(125, 27);
            txtHeaderRow.TabIndex = 3;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(159, 42);
            label1.Name = "label1";
            label1.Size = new Size(118, 20);
            label1.TabIndex = 3;
            label1.Text = "Header Row No.";
            // 
            // button2
            // 
            button2.BackColor = Color.CornflowerBlue;
            button2.Location = new Point(298, 117);
            button2.Name = "button2";
            button2.Size = new Size(101, 54);
            button2.TabIndex = 4;
            button2.Text = "Initiate";
            button2.UseVisualStyleBackColor = false;
            button2.Click += button2_Click;
            // 
            // chklst1
            // 
            chklst1.CheckOnClick = true;
            chklst1.FormattingEnabled = true;
            chklst1.Location = new Point(189, 57);
            chklst1.Name = "chklst1";
            chklst1.Size = new Size(150, 114);
            chklst1.TabIndex = 7;
            chklst1.SelectedIndexChanged += chklst1_SelectedIndexChanged;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(159, 104);
            label2.Name = "label2";
            label2.Size = new Size(131, 20);
            label2.TabIndex = 7;
            label2.Text = "Data Last Row No.";
            // 
            // txtLastRow
            // 
            txtLastRow.Location = new Point(159, 127);
            txtLastRow.Name = "txtLastRow";
            txtLastRow.Size = new Size(125, 27);
            txtLastRow.TabIndex = 4;
            // 
            // button3
            // 
            button3.BackColor = Color.CornflowerBlue;
            button3.Location = new Point(345, 117);
            button3.Name = "button3";
            button3.Size = new Size(101, 55);
            button3.TabIndex = 9;
            button3.Text = "Submit";
            button3.UseVisualStyleBackColor = false;
            button3.Click += button3_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.AllowUserToOrderColumns = true;
            dataGridView1.BackgroundColor = SystemColors.ButtonHighlight;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(0, 250);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.RowTemplate.Height = 29;
            dataGridView1.Size = new Size(692, 181);
            dataGridView1.TabIndex = 9;
            dataGridView1.CellContentClick += dataGridView1_CellContentClick;
            // 
            // checkedListBox1
            // 
            checkedListBox1.CheckOnClick = true;
            checkedListBox1.FormattingEnabled = true;
            checkedListBox1.Location = new Point(6, 57);
            checkedListBox1.Name = "checkedListBox1";
            checkedListBox1.Size = new Size(150, 114);
            checkedListBox1.TabIndex = 6;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(191, 34);
            label4.Name = "label4";
            label4.Size = new Size(148, 20);
            label4.TabIndex = 12;
            label4.Text = "Aggregating Column";
            // 
            // checkedListBox2
            // 
            checkedListBox2.CheckOnClick = true;
            checkedListBox2.FormattingEnabled = true;
            checkedListBox2.Location = new Point(6, 58);
            checkedListBox2.Name = "checkedListBox2";
            checkedListBox2.Size = new Size(150, 114);
            checkedListBox2.TabIndex = 10;
            // 
            // button4
            // 
            button4.BackColor = Color.CornflowerBlue;
            button4.Location = new Point(156, 116);
            button4.Name = "button4";
            button4.Size = new Size(101, 55);
            button4.TabIndex = 11;
            button4.Text = "Get distinct";
            button4.UseVisualStyleBackColor = false;
            button4.Click += button4_Click;
            // 
            // groupBox1
            // 
            groupBox1.BackColor = SystemColors.GradientActiveCaption;
            groupBox1.Controls.Add(chkSelectedColumns);
            groupBox1.Controls.Add(comboBox1);
            groupBox1.Controls.Add(checkedListBox1);
            groupBox1.Controls.Add(chklst1);
            groupBox1.Controls.Add(button3);
            groupBox1.Controls.Add(label4);
            groupBox1.FlatStyle = FlatStyle.Popup;
            groupBox1.Location = new Point(601, 44);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(457, 182);
            groupBox1.TabIndex = 6;
            groupBox1.TabStop = false;
            groupBox1.Text = "Group by";
            // 
            // chkSelectedColumns
            // 
            chkSelectedColumns.AutoSize = true;
            chkSelectedColumns.Checked = true;
            chkSelectedColumns.CheckState = CheckState.Checked;
            chkSelectedColumns.Location = new Point(6, 33);
            chkSelectedColumns.Name = "chkSelectedColumns";
            chkSelectedColumns.Size = new Size(147, 24);
            chkSelectedColumns.TabIndex = 5;
            chkSelectedColumns.Text = "Selected columns";
            chkSelectedColumns.UseVisualStyleBackColor = true;
            // 
            // comboBox1
            // 
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox1.FormattingEnabled = true;
            comboBox1.Items.AddRange(new object[] { "Sum", "Count", "Max", "Min" });
            comboBox1.Location = new Point(355, 57);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(77, 28);
            comboBox1.TabIndex = 8;
            // 
            // groupBox2
            // 
            groupBox2.BackColor = SystemColors.GradientActiveCaption;
            groupBox2.Controls.Add(checkedListBox2);
            groupBox2.Controls.Add(button4);
            groupBox2.FlatStyle = FlatStyle.Popup;
            groupBox2.Location = new Point(1080, 44);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(267, 183);
            groupBox2.TabIndex = 10;
            groupBox2.TabStop = false;
            groupBox2.Text = "Distinct Columns";
            // 
            // comboBox2
            // 
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.FormattingEnabled = true;
            comboBox2.Items.AddRange(new object[] { "Excel", "Comma(,) Separated", "Semicolon(;) Separated", "Space( ) Separated" });
            comboBox2.Location = new Point(5, 60);
            comboBox2.Name = "comboBox2";
            comboBox2.Size = new Size(137, 28);
            comboBox2.TabIndex = 0;
            // 
            // txtSearch
            // 
            txtSearch.Location = new Point(11, 66);
            txtSearch.Name = "txtSearch";
            txtSearch.Size = new Size(379, 48);
            txtSearch.TabIndex = 19;
            txtSearch.Text = "";
            // 
            // button5
            // 
            button5.BackColor = Color.CornflowerBlue;
            button5.Location = new Point(176, 120);
            button5.Name = "button5";
            button5.Size = new Size(101, 51);
            button5.TabIndex = 20;
            button5.Text = "Search";
            button5.UseVisualStyleBackColor = false;
            button5.Click += button5_Click;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(11, 23);
            label6.Name = "label6";
            label6.Size = new Size(358, 40);
            label6.TabIndex = 21;
            label6.Text = "ColumnaName='value'\r\nColumnName1='value1' and ColumnName2='value2'\r\n";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(5, 37);
            label7.Name = "label7";
            label7.Size = new Size(78, 20);
            label7.TabIndex = 22;
            label7.Text = "Input Type";
            // 
            // groupBox4
            // 
            groupBox4.BackColor = SystemColors.GradientActiveCaption;
            groupBox4.Controls.Add(label5);
            groupBox4.Controls.Add(txtLastRow);
            groupBox4.Controls.Add(listBox1);
            groupBox4.Controls.Add(txtHeaderRow);
            groupBox4.Controls.Add(label1);
            groupBox4.Controls.Add(button2);
            groupBox4.Controls.Add(label2);
            groupBox4.FlatStyle = FlatStyle.Popup;
            groupBox4.Location = new Point(176, 44);
            groupBox4.Name = "groupBox4";
            groupBox4.Size = new Size(405, 182);
            groupBox4.TabIndex = 2;
            groupBox4.TabStop = false;
            groupBox4.Text = "Data load";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(3, 30);
            label5.Name = "label5";
            label5.Size = new Size(99, 20);
            label5.TabIndex = 8;
            label5.Text = "Choose Sheet";
            // 
            // groupBox3
            // 
            groupBox3.BackColor = SystemColors.GradientActiveCaption;
            groupBox3.Controls.Add(comboBox2);
            groupBox3.Controls.Add(button1);
            groupBox3.Controls.Add(label7);
            groupBox3.FlatStyle = FlatStyle.Popup;
            groupBox3.Location = new Point(6, 45);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(154, 182);
            groupBox3.TabIndex = 0;
            groupBox3.TabStop = false;
            groupBox3.Text = "File Choose";
            // 
            // button6
            // 
            button6.BackColor = Color.CornflowerBlue;
            button6.Location = new Point(297, 120);
            button6.Name = "button6";
            button6.Size = new Size(93, 49);
            button6.TabIndex = 25;
            button6.Text = "Download";
            button6.UseVisualStyleBackColor = false;
            button6.Click += button6_Click;
            // 
            // groupBox5
            // 
            groupBox5.Controls.Add(groupBox6);
            groupBox5.Controls.Add(groupBox3);
            groupBox5.Controls.Add(groupBox1);
            groupBox5.Controls.Add(groupBox2);
            groupBox5.Controls.Add(groupBox4);
            groupBox5.Dock = DockStyle.Top;
            groupBox5.Location = new Point(0, 28);
            groupBox5.Name = "groupBox5";
            groupBox5.Size = new Size(1776, 244);
            groupBox5.TabIndex = 26;
            groupBox5.TabStop = false;
            groupBox5.Text = "Logic";
            // 
            // groupBox6
            // 
            groupBox6.BackColor = SystemColors.GradientActiveCaption;
            groupBox6.Controls.Add(label6);
            groupBox6.Controls.Add(button6);
            groupBox6.Controls.Add(button5);
            groupBox6.Controls.Add(txtSearch);
            groupBox6.Location = new Point(1367, 44);
            groupBox6.Name = "groupBox6";
            groupBox6.Size = new Size(422, 183);
            groupBox6.TabIndex = 27;
            groupBox6.TabStop = false;
            groupBox6.Text = "Manual Search";
            // 
            // menuStrip1
            // 
            menuStrip1.ImageScalingSize = new Size(20, 20);
            menuStrip1.Items.AddRange(new ToolStripItem[] { easyAggregationToolStripMenuItem, easyCompareToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(1776, 28);
            menuStrip1.TabIndex = 27;
            menuStrip1.Text = "menuStrip1";
            // 
            // easyAggregationToolStripMenuItem
            // 
            easyAggregationToolStripMenuItem.Name = "easyAggregationToolStripMenuItem";
            easyAggregationToolStripMenuItem.Size = new Size(140, 24);
            easyAggregationToolStripMenuItem.Text = "Easy Aggregation";
            easyAggregationToolStripMenuItem.Click += easyAggregationToolStripMenuItem_Click;
            // 
            // easyCompareToolStripMenuItem
            // 
            easyCompareToolStripMenuItem.Name = "easyCompareToolStripMenuItem";
            easyCompareToolStripMenuItem.Size = new Size(117, 24);
            easyCompareToolStripMenuItem.Text = "Easy Compare";
            easyCompareToolStripMenuItem.Click += easyCompareToolStripMenuItem_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.ButtonFace;
            ClientSize = new Size(1776, 764);
            Controls.Add(groupBox5);
            Controls.Add(dataGridView1);
            Controls.Add(menuStrip1);
            MainMenuStrip = menuStrip1;
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "XLOps";
            WindowState = FormWindowState.Maximized;
            Closing += Form1_Closing;
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            groupBox3.ResumeLayout(false);
            groupBox3.PerformLayout();
            groupBox5.ResumeLayout(false);
            groupBox6.ResumeLayout(false);
            groupBox6.PerformLayout();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private ListBox listBox1;
        private OpenFileDialog openFileDialog1;
        private TextBox txtHeaderRow;
        private Label label1;
        private Button button2;
        private CheckedListBox chklst1;
        private Label label2;
        private TextBox txtLastRow;
        private Button button3;
        private DataGridView dataGridView1;
        private CheckedListBox checkedListBox1;
        private Label label4;
        private CheckedListBox checkedListBox2;
        private Button button4;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
        private ComboBox comboBox1;
        private ComboBox comboBox2;
        private RichTextBox txtSearch;
        private Button button5;
        private Label label6;
        private Label label7;
        private GroupBox groupBox4;
        private Label label5;
        private GroupBox groupBox3;
        private Button button6;
        private SaveFileDialog saveFileDialog1;
        private CheckBox chkSelectedColumns;
        private GroupBox groupBox5;
        private GroupBox groupBox6;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem easyAggregationToolStripMenuItem;
        private ToolStripMenuItem easyCompareToolStripMenuItem;
    }
}