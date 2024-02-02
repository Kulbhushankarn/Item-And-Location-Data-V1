namespace Item_And_Location_Data_V1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.textBox_inputFile = new System.Windows.Forms.TextBox();
            this.textBox_outputFile = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(28, 94);
            this.button1.Name = "btn_selectExcelfile";
            this.button1.Size = new System.Drawing.Size(154, 33);
            this.button1.TabIndex = 0;
            this.button1.Text = "Select Excel File";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.btn_selectExcelfile);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(28, 180);
            this.button2.Name = "btn_selectOutputfile";
            this.button2.Size = new System.Drawing.Size(154, 33);
            this.button2.TabIndex = 1;
            this.button2.Text = "Select Output Path";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.btn_selectOutputfile);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(381, 310);
            this.button3.Name = "btn_process";
            this.button3.Size = new System.Drawing.Size(115, 30);
            this.button3.TabIndex = 2;
            this.button3.Text = "Process";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.btn_process);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(28, 380);
            this.button4.Name = "btn_exit";
            this.button4.Size = new System.Drawing.Size(154, 36);
            this.button4.TabIndex = 3;
            this.button4.Text = "Exit";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.btn_exit);
            // 
            // textBox_inputFile
            // 
            this.textBox_inputFile.Location = new System.Drawing.Point(214, 94);
            this.textBox_inputFile.Multiline = true;
            this.textBox_inputFile.Name = "textBox_inputFile";
            this.textBox_inputFile.Size = new System.Drawing.Size(494, 33);
            this.textBox_inputFile.TabIndex = 4;
            // 
            // textBox_outputFile
            // 
            this.textBox_outputFile.Location = new System.Drawing.Point(214, 180);
            this.textBox_outputFile.Multiline = true;
            this.textBox_outputFile.Name = "textBox_outputFile";
            this.textBox_outputFile.Size = new System.Drawing.Size(494, 33);
            this.textBox_outputFile.TabIndex = 5;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Item_And_Location_Data_V1.Properties.Resources.logo;
            this.pictureBox1.Location = new System.Drawing.Point(612, 310);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(145, 87);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 6;
            this.pictureBox1.TabStop = false;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(214, 247);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(494, 23);
            this.progressBar1.TabIndex = 7;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.textBox_outputFile);
            this.Controls.Add(this.textBox_inputFile);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Item And Location Data V1";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TextBox textBox_inputFile;
        private System.Windows.Forms.TextBox textBox_outputFile;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ProgressBar progressBar1;
    }
}

