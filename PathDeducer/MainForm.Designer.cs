namespace PathDeducer
{
    partial class MainForm
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
            this.button_SelectExcel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button_SelectExcel
            // 
            this.button_SelectExcel.Location = new System.Drawing.Point(12, 12);
            this.button_SelectExcel.Name = "button_SelectExcel";
            this.button_SelectExcel.Size = new System.Drawing.Size(75, 23);
            this.button_SelectExcel.TabIndex = 0;
            this.button_SelectExcel.Text = "Select";
            this.button_SelectExcel.UseVisualStyleBackColor = true;
            this.button_SelectExcel.Click += new System.EventHandler(this.button_SelectExcel_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.button_SelectExcel);
            this.Name = "MainForm";
            this.Text = "Path Deducer";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button_SelectExcel;
    }
}

