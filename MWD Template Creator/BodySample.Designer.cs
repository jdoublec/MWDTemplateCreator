namespace MWD_Template_Creator
{
    partial class frmBodySample
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
            this.rtxtTemplate = new System.Windows.Forms.RichTextBox();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // rtxtTemplate
            // 
            this.rtxtTemplate.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rtxtTemplate.BackColor = System.Drawing.Color.Gainsboro;
            this.rtxtTemplate.Font = new System.Drawing.Font("Courier New", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rtxtTemplate.ForeColor = System.Drawing.Color.Black;
            this.rtxtTemplate.Location = new System.Drawing.Point(0, 37);
            this.rtxtTemplate.Name = "rtxtTemplate";
            this.rtxtTemplate.Size = new System.Drawing.Size(752, 638);
            this.rtxtTemplate.TabIndex = 201;
            this.rtxtTemplate.TabStop = false;
            this.rtxtTemplate.Text = "";
            this.rtxtTemplate.WordWrap = false;
            // 
            // txtSubject
            // 
            this.txtSubject.BackColor = System.Drawing.Color.Gainsboro;
            this.txtSubject.Location = new System.Drawing.Point(0, 11);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(752, 20);
            this.txtSubject.TabIndex = 202;
            // 
            // frmBodySample
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SlateGray;
            this.ClientSize = new System.Drawing.Size(752, 675);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.rtxtTemplate);
            this.Name = "frmBodySample";
            this.Text = "Body Sample";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox rtxtTemplate;
        private System.Windows.Forms.TextBox txtSubject;
    }
}