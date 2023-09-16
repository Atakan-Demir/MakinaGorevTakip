namespace MakineGorevTakip
{
    partial class frmSettings
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
            this.label1 = new System.Windows.Forms.Label();
            this.btnGozat = new System.Windows.Forms.Button();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.btnTamam = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F);
            this.label1.ForeColor = System.Drawing.Color.DarkGray;
            this.label1.Location = new System.Drawing.Point(3, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(708, 50);
            this.label1.TabIndex = 6;
            this.label1.Text = "Lütfen veri tabanı olarak tutulan Excel (.xlsx) dosyasının yolunu aşağıdan seçini" +
    "z.\r\nBütün program parçacıklarında bu yol baz alınarak işlem gerçekleştirilmekted" +
    "ir.";
            // 
            // btnGozat
            // 
            this.btnGozat.FlatAppearance.BorderSize = 0;
            this.btnGozat.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGozat.Font = new System.Drawing.Font("Nirmala UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGozat.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(126)))), ((int)(((byte)(249)))));
            this.btnGozat.Location = new System.Drawing.Point(8, 179);
            this.btnGozat.Name = "btnGozat";
            this.btnGozat.Size = new System.Drawing.Size(139, 42);
            this.btnGozat.TabIndex = 7;
            this.btnGozat.Text = "Gözat   :";
            this.btnGozat.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
            this.btnGozat.UseVisualStyleBackColor = true;
            this.btnGozat.Click += new System.EventHandler(this.btnGozat_Click);
            // 
            // txtPath
            // 
            this.txtPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.txtPath.Location = new System.Drawing.Point(153, 184);
            this.txtPath.Name = "txtPath";
            this.txtPath.ReadOnly = true;
            this.txtPath.Size = new System.Drawing.Size(558, 35);
            this.txtPath.TabIndex = 8;
            // 
            // btnTamam
            // 
            this.btnTamam.FlatAppearance.BorderSize = 0;
            this.btnTamam.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnTamam.Font = new System.Drawing.Font("Nirmala UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTamam.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(126)))), ((int)(((byte)(249)))));
            this.btnTamam.Location = new System.Drawing.Point(12, 246);
            this.btnTamam.Name = "btnTamam";
            this.btnTamam.Size = new System.Drawing.Size(139, 42);
            this.btnTamam.TabIndex = 9;
            this.btnTamam.Text = "Tamam";
            this.btnTamam.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
            this.btnTamam.UseVisualStyleBackColor = true;
            this.btnTamam.Click += new System.EventHandler(this.btnTamam_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F);
            this.label2.ForeColor = System.Drawing.Color.DarkGray;
            this.label2.Location = new System.Drawing.Point(157, 257);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(511, 25);
            this.label2.TabIndex = 10;
            this.label2.Text = "Dosya yolunu seçtikten sonra \"Tamam\" butonuna basınız.";
            // 
            // frmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.ClientSize = new System.Drawing.Size(733, 477);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnTamam);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.btnGozat);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "frmSettings";
            this.Text = "frmSettings";
            this.Load += new System.EventHandler(this.frmSettings_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnGozat;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Button btnTamam;
        private System.Windows.Forms.Label label2;
    }
}