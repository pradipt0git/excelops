namespace ExcelOps
{
    partial class DataManage
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
            comboBox2 = new ComboBox();
            button1 = new Button();
            label7 = new Label();
            groupBox4 = new GroupBox();
            numericUpDown2 = new NumericUpDown();
            numericUpDown1 = new NumericUpDown();
            label3 = new Label();
            listBox1 = new ListBox();
            label1 = new Label();
            button2 = new Button();
            label2 = new Label();
            groupBox1 = new GroupBox();
            chkSelectedColumns = new CheckBox();
            comboBox1 = new ComboBox();
            checkedListBox1 = new CheckedListBox();
            chklst1 = new CheckedListBox();
            button3 = new Button();
            label4 = new Label();
            progressBar1 = new ProgressBar();
            label5 = new Label();
            groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)numericUpDown2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown1).BeginInit();
            groupBox1.SuspendLayout();
            SuspendLayout();
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
            // button1
            // 
            button1.BackColor = Color.CornflowerBlue;
            button1.Location = new Point(244, 17);
            button1.Name = "button1";
            button1.Size = new Size(101, 37);
            button1.TabIndex = 0;
            button1.Text = "Browse";
            button1.UseVisualStyleBackColor = false;
            button1.Click += button1_Click;
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
            groupBox4.TabIndex = 26;
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
            groupBox1.Location = new Point(405, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(457, 238);
            groupBox1.TabIndex = 27;
            groupBox1.TabStop = false;
            groupBox1.Text = "Group by";
            // 
            // chkSelectedColumns
            // 
            chkSelectedColumns.AutoSize = true;
            chkSelectedColumns.Checked = true;
            chkSelectedColumns.CheckState = CheckState.Checked;
            chkSelectedColumns.Location = new Point(319, 178);
            chkSelectedColumns.Name = "chkSelectedColumns";
            chkSelectedColumns.Size = new Size(147, 24);
            chkSelectedColumns.TabIndex = 14;
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
            comboBox1.TabIndex = 13;
            // 
            // checkedListBox1
            // 
            checkedListBox1.CheckOnClick = true;
            checkedListBox1.FormattingEnabled = true;
            checkedListBox1.Location = new Point(6, 26);
            checkedListBox1.Name = "checkedListBox1";
            checkedListBox1.Size = new Size(146, 202);
            checkedListBox1.TabIndex = 10;
            checkedListBox1.SelectedIndexChanged += checkedListBox1_SelectedIndexChanged;
            // 
            // chklst1
            // 
            chklst1.CheckOnClick = true;
            chklst1.FormattingEnabled = true;
            chklst1.Location = new Point(163, 26);
            chklst1.Name = "chklst1";
            chklst1.Size = new Size(150, 202);
            chklst1.TabIndex = 5;
            // 
            // button3
            // 
            button3.BackColor = Color.CornflowerBlue;
            button3.Location = new Point(345, 117);
            button3.Name = "button3";
            button3.Size = new Size(101, 55);
            button3.TabIndex = 8;
            button3.Text = "Submit";
            button3.UseVisualStyleBackColor = false;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(165, 0);
            label4.Name = "label4";
            label4.Size = new Size(148, 20);
            label4.TabIndex = 12;
            label4.Text = "Aggregating Column";
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(12, 256);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(850, 29);
            progressBar1.Step = 1;
            progressBar1.TabIndex = 28;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(914, 93);
            label5.Name = "label5";
            label5.Size = new Size(50, 20);
            label5.TabIndex = 29;
            label5.Text = "label5";
            // 
            // DataManage
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1257, 450);
            Controls.Add(label5);
            Controls.Add(progressBar1);
            Controls.Add(groupBox1);
            Controls.Add(groupBox4);
            Name = "DataManage";
            Text = "DataManage";
            Load += DataManage_Load;
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)numericUpDown2).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown1).EndInit();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private ComboBox comboBox2;
        private Button button1;
        private Label label7;
        private GroupBox groupBox4;
        private ListBox listBox1;
        private Label label1;
        private Button button2;
        private Label label2;
        private NumericUpDown numericUpDown2;
        private NumericUpDown numericUpDown1;
        private Label label3;
        private GroupBox groupBox1;
        private CheckBox chkSelectedColumns;
        private ComboBox comboBox1;
        private CheckedListBox checkedListBox1;
        private CheckedListBox chklst1;
        private Button button3;
        private Label label4;
        private ProgressBar progressBar1;
        private Label label5;
    }
}