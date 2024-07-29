namespace ExcelOps
{
    partial class LoadExcelAndManage
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            groupBox4 = new GroupBox();
            numericUpDown2 = new NumericUpDown();
            numericUpDown1 = new NumericUpDown();
            label3 = new Label();
            button1 = new Button();
            comboBox2 = new ComboBox();
            label7 = new Label();
            listBox1 = new ListBox();
            label1 = new Label();
            button2 = new Button();
            label2 = new Label();
            dataGridView1 = new DataGridView();
            groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)numericUpDown2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // groupBox4
            // 
            groupBox4.BackColor = SystemColors.GradientActiveCaption;
            groupBox4.Controls.Add(numericUpDown2);
            groupBox4.Controls.Add(numericUpDown1);
            groupBox4.Controls.Add(label3);
            groupBox4.Controls.Add(button1);
            groupBox4.Controls.Add(comboBox2);
            groupBox4.Controls.Add(label7);
            groupBox4.Controls.Add(listBox1);
            groupBox4.Controls.Add(label1);
            groupBox4.Controls.Add(button2);
            groupBox4.Controls.Add(label2);
            groupBox4.FlatStyle = FlatStyle.Popup;
            groupBox4.Location = new Point(12, 12);
            groupBox4.Name = "groupBox4";
            groupBox4.Size = new Size(387, 238);
            groupBox4.TabIndex = 27;
            groupBox4.TabStop = false;
            groupBox4.Text = "Data load";
            // 
            // numericUpDown2
            // 
            numericUpDown2.Location = new Point(244, 145);
            numericUpDown2.Name = "numericUpDown2";
            numericUpDown2.Size = new Size(69, 27);
            numericUpDown2.TabIndex = 28;
            // 
            // numericUpDown1
            // 
            numericUpDown1.Location = new Point(244, 83);
            numericUpDown1.Name = "numericUpDown1";
            numericUpDown1.Size = new Size(69, 27);
            numericUpDown1.TabIndex = 27;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(6, 60);
            label3.Name = "label3";
            label3.Size = new Size(52, 20);
            label3.TabIndex = 23;
            label3.Text = "Sheets";
            // 
            // button1
            // 
            button1.BackColor = Color.CornflowerBlue;
            button1.Location = new Point(244, 17);
            button1.Name = "button1";
            button1.Size = new Size(101, 37);
            button1.TabIndex = 0;
            button1.Text = "Browse";
            button1.UseVisualStyleBackColor = false;
            button1.Click += button1_Click_1;
            // 
            // comboBox2
            // 
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.FormattingEnabled = true;
            comboBox2.Items.AddRange(new object[] { "Excel", "Comma(,) Separated", "Semicolon(;) Separated", "Space( ) Separated" });
            comboBox2.Location = new Point(90, 26);
            comboBox2.Name = "comboBox2";
            comboBox2.Size = new Size(137, 28);
            comboBox2.TabIndex = 14;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(6, 32);
            label7.Name = "label7";
            label7.Size = new Size(78, 20);
            label7.TabIndex = 22;
            label7.Text = "Input Type";
            // 
            // listBox1
            // 
            listBox1.FormattingEnabled = true;
            listBox1.ItemHeight = 20;
            listBox1.Location = new Point(90, 60);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(137, 164);
            listBox1.TabIndex = 1;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(244, 60);
            label1.Name = "label1";
            label1.Size = new Size(118, 20);
            label1.TabIndex = 3;
            label1.Text = "Header Row No.";
            // 
            // button2
            // 
            button2.BackColor = Color.CornflowerBlue;
            button2.Location = new Point(244, 185);
            button2.Name = "button2";
            button2.Size = new Size(101, 39);
            button2.TabIndex = 4;
            button2.Text = "Initiate";
            button2.UseVisualStyleBackColor = false;
            button2.Click += button2_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(244, 122);
            label2.Name = "label2";
            label2.Size = new Size(131, 20);
            label2.TabIndex = 7;
            label2.Text = "Data Last Row No.";
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(12, 256);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.RowTemplate.Height = 29;
            dataGridView1.Size = new Size(1353, 466);
            dataGridView1.TabIndex = 28;
            // 
            // LoadExcelAndManage
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1377, 724);
            Controls.Add(dataGridView1);
            Controls.Add(groupBox4);
            Name = "LoadExcelAndManage";
            Text = "LoadExcelAndManage";
            Load += LoadExcelAndManage_Load;
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)numericUpDown2).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown1).EndInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private GroupBox groupBox4;
        private NumericUpDown numericUpDown2;
        private NumericUpDown numericUpDown1;
        private Label label3;
        private Button button1;
        private ComboBox comboBox2;
        private Label label7;
        private ListBox listBox1;
        private Label label1;
        private Button button2;
        private Label label2;
        private DataGridView dataGridView1;
    }
}