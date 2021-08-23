namespace ExcelBelegger
{
    partial class FileChooser
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FileChooser));
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.decisionLabel = new System.Windows.Forms.Label();
            this.OkButton = new System.Windows.Forms.Button();
            this.cButton = new System.Windows.Forms.Button();
            this.bButton = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(12, 38);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(131, 21);
            this.comboBox1.TabIndex = 0;
            // 
            // decisionLabel
            // 
            this.decisionLabel.AutoSize = true;
            this.decisionLabel.Location = new System.Drawing.Point(13, 13);
            this.decisionLabel.Name = "decisionLabel";
            this.decisionLabel.Size = new System.Drawing.Size(221, 13);
            this.decisionLabel.TabIndex = 1;
            this.decisionLabel.Text = "Choose which type of CSV you want to open:";
            // 
            // OkButton
            // 
            this.OkButton.Location = new System.Drawing.Point(12, 177);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(75, 23);
            this.OkButton.TabIndex = 2;
            this.OkButton.Text = "OK";
            this.OkButton.UseVisualStyleBackColor = true;
            this.OkButton.Click += new System.EventHandler(this.OkBtn_Click);
            // 
            // cButton
            // 
            this.cButton.Location = new System.Drawing.Point(158, 177);
            this.cButton.Name = "cButton";
            this.cButton.Size = new System.Drawing.Size(75, 23);
            this.cButton.TabIndex = 3;
            this.cButton.Text = "Cancel";
            this.cButton.UseVisualStyleBackColor = true;
            this.cButton.Click += new System.EventHandler(this.cancelBtn_Click);
            // 
            // bButton
            // 
            this.bButton.Location = new System.Drawing.Point(159, 38);
            this.bButton.Name = "bButton";
            this.bButton.Size = new System.Drawing.Size(75, 23);
            this.bButton.TabIndex = 4;
            this.bButton.Text = "Browse";
            this.bButton.UseVisualStyleBackColor = true;
            this.bButton.Click += new System.EventHandler(this.browseBtn_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.InitialImage = null;
            this.pictureBox1.Location = new System.Drawing.Point(12, 66);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(222, 105);
            this.pictureBox1.TabIndex = 5;
            this.pictureBox1.TabStop = false;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // FileChooser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(245, 212);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.bButton);
            this.Controls.Add(this.cButton);
            this.Controls.Add(this.OkButton);
            this.Controls.Add(this.decisionLabel);
            this.Controls.Add(this.comboBox1);
            this.Name = "FileChooser";
            this.Text = "CSV Importer";
            this.Load += new System.EventHandler(this.FileChooser_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label decisionLabel;
        private System.Windows.Forms.Button OkButton;
        private System.Windows.Forms.Button cButton;
        private System.Windows.Forms.Button bButton;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}