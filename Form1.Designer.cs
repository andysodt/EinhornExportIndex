namespace EinhornExportIndex
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be
        ///     disposed; otherwise, false.</param>
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
        /// Required method for Designer support - do not modify the contents of
        /// this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            CreateIndex = new Button();
            label1 = new Label();
            textBox1 = new TextBox();
            SuspendLayout();
            // 
            // CreateIndex
            // 
            CreateIndex.Location = new Point(19, 20);
            CreateIndex.Margin = new Padding(6, 5, 6, 5);
            CreateIndex.Name = "CreateIndex";
            CreateIndex.Size = new Size(397, 85);
            CreateIndex.TabIndex = 2;
            CreateIndex.Text = "Click to choose project folder.";
            CreateIndex.UseVisualStyleBackColor = true;
            CreateIndex.Click += EinhornExportIndex_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(20, 160);
            label1.Margin = new Padding(6, 0, 6, 0);
            label1.Name = "label1";
            label1.Size = new Size(85, 25);
            label1.TabIndex = 3;
            label1.Text = "Progress:";
            // 
            // textBox1
            // 
            textBox1.Location = new Point(22, 190);
            textBox1.Margin = new Padding(6, 5, 6, 5);
            textBox1.Multiline = true;
            textBox1.Name = "textBox1";
            textBox1.ScrollBars = ScrollBars.Both;
            textBox1.Size = new Size(1083, 654);
            textBox1.TabIndex = 6;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1135, 871);
            Controls.Add(textBox1);
            Controls.Add(label1);
            Controls.Add(CreateIndex);
            Margin = new Padding(6, 5, 6, 5);
            Name = "Form1";
            Text = "Update Index Spreadsheet";
            Load += EinhornExportIndex_Load;
            ResumeLayout(false);
            PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button CreateIndex;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
    }
}