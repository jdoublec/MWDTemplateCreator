namespace MWD_Template_Creator
{
    partial class frmTemplateName
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
            this.txtTemplateName = new System.Windows.Forms.TextBox();
            this.btnOKDuplicate = new System.Windows.Forms.Button();
            this.lblTemplateName = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.cboTemplates = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // txtTemplateName
            // 
            this.txtTemplateName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTemplateName.Location = new System.Drawing.Point(36, 63);
            this.txtTemplateName.Name = "txtTemplateName";
            this.txtTemplateName.Size = new System.Drawing.Size(587, 22);
            this.txtTemplateName.TabIndex = 0;
            this.txtTemplateName.TextChanged += new System.EventHandler(this.txtTemplateName_TextChanged);
            // 
            // btnOKDuplicate
            // 
            this.btnOKDuplicate.Location = new System.Drawing.Point(36, 114);
            this.btnOKDuplicate.Name = "btnOKDuplicate";
            this.btnOKDuplicate.Size = new System.Drawing.Size(253, 31);
            this.btnOKDuplicate.TabIndex = 1;
            this.btnOKDuplicate.Text = "Ok";
            this.btnOKDuplicate.UseVisualStyleBackColor = true;
            this.btnOKDuplicate.Click += new System.EventHandler(this.btnOkDuplicate_Click);
            // 
            // lblTemplateName
            // 
            this.lblTemplateName.AutoSize = true;
            this.lblTemplateName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTemplateName.Location = new System.Drawing.Point(32, 27);
            this.lblTemplateName.Name = "lblTemplateName";
            this.lblTemplateName.Size = new System.Drawing.Size(121, 20);
            this.lblTemplateName.TabIndex = 2;
            this.lblTemplateName.Text = "Template Name";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(370, 114);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(253, 31);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // cboTemplates
            // 
            this.cboTemplates.FormattingEnabled = true;
            this.cboTemplates.Location = new System.Drawing.Point(36, 63);
            this.cboTemplates.Name = "cboTemplates";
            this.cboTemplates.Size = new System.Drawing.Size(587, 21);
            this.cboTemplates.TabIndex = 4;
            this.cboTemplates.TextChanged += new System.EventHandler(this.cboTemplates_TextChanged);
            // 
            // frmTemplateName
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(635, 162);
            this.Controls.Add(this.cboTemplates);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.lblTemplateName);
            this.Controls.Add(this.btnOKDuplicate);
            this.Controls.Add(this.txtTemplateName);
            this.Name = "frmTemplateName";
            this.Text = "Template Name";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtTemplateName;
        private System.Windows.Forms.Button btnOKDuplicate;
        private System.Windows.Forms.Label lblTemplateName;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ComboBox cboTemplates;
    }
}