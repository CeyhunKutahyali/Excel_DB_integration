namespace Excel_DB_integration
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
            richTextBox1 = new RichTextBox();
            btnExport = new Button();
            btnImport = new Button();
            richTextBox2 = new RichTextBox();
            SuspendLayout();
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(56, 41);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(513, 192);
            richTextBox1.TabIndex = 0;
            richTextBox1.Text = "";
            // 
            // btnExport
            // 
            btnExport.Location = new Point(575, 41);
            btnExport.Name = "btnExport";
            btnExport.Size = new Size(153, 54);
            btnExport.TabIndex = 1;
            btnExport.Text = "Excel'e Aktar";
            btnExport.UseVisualStyleBackColor = true;
            btnExport.Click += btnGetData_Click;
            // 
            // btnImport
            // 
            btnImport.Location = new Point(575, 277);
            btnImport.Name = "btnImport";
            btnImport.Size = new Size(153, 54);
            btnImport.TabIndex = 3;
            btnImport.Text = "Excel'den Al";
            btnImport.UseVisualStyleBackColor = true;
            btnImport.Click += btnImport_Click;
            // 
            // richTextBox2
            // 
            richTextBox2.Location = new Point(56, 277);
            richTextBox2.Name = "richTextBox2";
            richTextBox2.Size = new Size(513, 192);
            richTextBox2.TabIndex = 2;
            richTextBox2.Text = "";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            BackColor = Color.Silver;
            ClientSize = new Size(797, 512);
            Controls.Add(btnImport);
            Controls.Add(richTextBox2);
            Controls.Add(btnExport);
            Controls.Add(richTextBox1);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Veri Transferi";
            ResumeLayout(false);
        }

        #endregion

        private RichTextBox richTextBox1;
        private Button btnExport;
        private Button btnImport;
        private RichTextBox richTextBox2;
    }
}