
using System.Windows.Forms;

namespace one034
{
    partial class FormWybor
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormWybor));
            this.buttonZapisy = new System.Windows.Forms.Button();
            this.buttonExcel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonZapisy
            // 
            this.buttonZapisy.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.5F);
            this.buttonZapisy.Location = new System.Drawing.Point(15, 44);
            this.buttonZapisy.Name = "buttonZapisy";
            this.buttonZapisy.Size = new System.Drawing.Size(395, 45);
            this.buttonZapisy.TabIndex = 4;
            this.buttonZapisy.Text = "Rozlicz płatności za najem";
            this.buttonZapisy.UseVisualStyleBackColor = true;
            this.buttonZapisy.Click += new System.EventHandler(this.buttonZapisy_Click);
            // 
            // buttonExcel
            // 
            this.buttonExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.5F);
            this.buttonExcel.Location = new System.Drawing.Point(15, 117);
            this.buttonExcel.Name = "buttonExcel";
            this.buttonExcel.Size = new System.Drawing.Size(395, 45);
            this.buttonExcel.TabIndex = 4;
            this.buttonExcel.Text = "Rozlicz płatności wybierając arkusz kalkulacyjny";
            this.buttonExcel.UseVisualStyleBackColor = true;
            this.buttonExcel.Click += new System.EventHandler(this.buttonExcel_Click);
            // 
            // FormWybor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(427, 192);
            this.Controls.Add(this.buttonZapisy);
            this.Controls.Add(this.buttonExcel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormWybor";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Opcje działania dodatku";
            this.ResumeLayout(false);

        }

        private Button buttonZapisy;
        private Button buttonExcel;
        #endregion
    }
}