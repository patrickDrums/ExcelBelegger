namespace ExcelBelegger
{
    partial class SelectColumnForm
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

        public void setCollectionComboBoxes(string[] collectionArray)
        {
            foreach(string item in collectionArray)
            {
                cBDate.Items.Add(item);
                cBDescription.Items.Add(item);
                cBProduct.Items.Add(item);
                cBSaldo.Items.Add(item);
            }
        }

        public string getSelectedDateColumn()
        {
            return (string) cBDate.SelectedItem;
        }

        public string getSelectedProductColumn()
        {
            return (string)cBProduct.SelectedItem;
        }

        public string getSelectedDescriptionColumn()
        {
            return (string)cBDescription.SelectedItem;
        }

        public string getSelectedSaldoColumn()
        {
            return (string)cBSaldo.SelectedItem;
        }

        public int getSelectedIndexDateColumn()
        {
            return cBDate.SelectedIndex;
        }

        public int getSelectedIndexProductColumn()
        {
            return cBProduct.SelectedIndex;
        }

        public int getSelectedIndexDescriptionColumn()
        {
            return cBDescription.SelectedIndex;
        }

        public int getSelectedIndexSaldoColumn()
        {
            return cBSaldo.SelectedIndex;
        }


        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.cBDate = new System.Windows.Forms.ComboBox();
            this.cBProduct = new System.Windows.Forms.ComboBox();
            this.cBDescription = new System.Windows.Forms.ComboBox();
            this.cBSaldo = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cBDate
            // 
            this.cBDate.FormattingEnabled = true;
            this.cBDate.Location = new System.Drawing.Point(141, 12);
            this.cBDate.Name = "cBDate";
            this.cBDate.Size = new System.Drawing.Size(121, 21);
            this.cBDate.TabIndex = 0;
            // 
            // cBProduct
            // 
            this.cBProduct.FormattingEnabled = true;
            this.cBProduct.Location = new System.Drawing.Point(141, 40);
            this.cBProduct.Name = "cBProduct";
            this.cBProduct.Size = new System.Drawing.Size(121, 21);
            this.cBProduct.TabIndex = 1;
            // 
            // cBDescription
            // 
            this.cBDescription.FormattingEnabled = true;
            this.cBDescription.Location = new System.Drawing.Point(141, 68);
            this.cBDescription.Name = "cBDescription";
            this.cBDescription.Size = new System.Drawing.Size(121, 21);
            this.cBDescription.TabIndex = 2;
            // 
            // cBSaldo
            // 
            this.cBSaldo.FormattingEnabled = true;
            this.cBSaldo.Location = new System.Drawing.Point(141, 96);
            this.cBSaldo.Name = "cBSaldo";
            this.cBSaldo.Size = new System.Drawing.Size(121, 21);
            this.cBSaldo.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Datum";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Product";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 68);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(67, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Omschrijving";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 96);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(34, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Saldo";
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button1.Location = new System.Drawing.Point(16, 138);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(246, 23);
            this.button1.TabIndex = 8;
            this.button1.Text = "Gereed";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // SelectColumnForm
            // 
            this.AcceptButton = this.button1;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(285, 173);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cBSaldo);
            this.Controls.Add(this.cBDescription);
            this.Controls.Add(this.cBProduct);
            this.Controls.Add(this.cBDate);
            this.Name = "SelectColumnForm";
            this.Text = "Selecteer  kolommen";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cBDate;
        private System.Windows.Forms.ComboBox cBProduct;
        private System.Windows.Forms.ComboBox cBDescription;
        private System.Windows.Forms.ComboBox cBSaldo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button1;
    }
}