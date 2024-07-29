namespace ExcelOps
{
    partial class ExcelSelection
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
            button1 = new Button();
            label2 = new Label();
            label3 = new Label();
            label4 = new Label();
            label5 = new Label();
            textBox1 = new TextBox();
            textBox2 = new TextBox();
            button2 = new Button();
            label6 = new Label();
            listBox1 = new ListBox();
            textBox3 = new TextBox();
            textBox4 = new TextBox();
            backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(56, 12);
            button1.Name = "button1";
            button1.Size = new Size(439, 36);
            button1.TabIndex = 3;
            button1.Text = "Browse";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(359, 129);
            label2.Name = "label2";
            label2.Size = new Size(160, 20);
            label2.TabIndex = 5;
            label2.Text = "Start Column Alphabet";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(671, 129);
            label3.Name = "label3";
            label3.Size = new Size(131, 20);
            label3.TabIndex = 6;
            label3.Text = "Start Row Number";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(363, 194);
            label4.Name = "label4";
            label4.Size = new Size(154, 20);
            label4.TabIndex = 7;
            label4.Text = "End Column Alphabet";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(677, 198);
            label5.Name = "label5";
            label5.Size = new Size(125, 20);
            label5.TabIndex = 8;
            label5.Text = "End Row Number";
            // 
            // textBox1
            // 
            textBox1.Location = new Point(523, 128);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(125, 27);
            textBox1.TabIndex = 9;
            textBox1.Text = "A";
            // 
            // textBox2
            // 
            textBox2.Location = new Point(523, 191);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(125, 27);
            textBox2.TabIndex = 11;
            // 
            // button2
            // 
            button2.Location = new Point(56, 273);
            button2.Name = "button2";
            button2.Size = new Size(439, 42);
            button2.TabIndex = 13;
            button2.Text = "Submit";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(61, 55);
            label6.Name = "label6";
            label6.Size = new Size(0, 20);
            label6.TabIndex = 14;
            // 
            // listBox1
            // 
            listBox1.FormattingEnabled = true;
            listBox1.ItemHeight = 20;
            listBox1.Location = new Point(56, 88);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(237, 164);
            listBox1.TabIndex = 15;
            listBox1.SelectedIndexChanged += listBox1_SelectedIndexChanged;
            // 
            // textBox3
            // 
            textBox3.Location = new Point(836, 126);
            textBox3.Name = "textBox3";
            textBox3.Size = new Size(125, 27);
            textBox3.TabIndex = 16;
            textBox3.Text = "1";
            // 
            // textBox4
            // 
            textBox4.Location = new Point(836, 191);
            textBox4.Name = "textBox4";
            textBox4.Size = new Size(125, 27);
            textBox4.TabIndex = 17;
            // 
            // ExcelSelection
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1027, 343);
            Controls.Add(textBox4);
            Controls.Add(textBox3);
            Controls.Add(listBox1);
            Controls.Add(label6);
            Controls.Add(button2);
            Controls.Add(textBox2);
            Controls.Add(textBox1);
            Controls.Add(label5);
            Controls.Add(label4);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(button1);
            Name = "ExcelSelection";
            Text = "ExcelSelection";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private Button button1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label5;
        private TextBox textBox1;
        private TextBox textBox2;
        private Button button2;
        private Label label6;
        private ListBox listBox1;
        private TextBox textBox3;
        private TextBox textBox4;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
    }
}