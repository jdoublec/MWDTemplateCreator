namespace MWD_Template_Creator
{
    partial class frmTemplates
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmTemplates));
            this.dgvBodyTemplate = new System.Windows.Forms.DataGridView();
            this.dgvSubjectTemplate = new System.Windows.Forms.DataGridView();
            this.cboTemplates = new System.Windows.Forms.ComboBox();
            this.lblTemplate = new System.Windows.Forms.Label();
            this.btnNewTemplate = new System.Windows.Forms.Button();
            this.btnSaveTemplate = new System.Windows.Forms.Button();
            this.lblSubjectBuilder = new System.Windows.Forms.Label();
            this.lblBodyBuilder = new System.Windows.Forms.Label();
            this.btnAddSubjectRow = new System.Windows.Forms.Button();
            this.btnAddBodyRow = new System.Windows.Forms.Button();
            this.chkBodyDuplicateSettings = new System.Windows.Forms.CheckBox();
            this.lblBodySample = new System.Windows.Forms.Label();
            this.lblSubjectSample = new System.Windows.Forms.Label();
            this.rtxtBodySample = new System.Windows.Forms.RichTextBox();
            this.txtSubjectSample = new System.Windows.Forms.TextBox();
            this.lblStaticIndentInfo = new System.Windows.Forms.Label();
            this.btnBodyMoveDown = new System.Windows.Forms.Button();
            this.btnBodyMoveUp = new System.Windows.Forms.Button();
            this.btnSubjectMoveDown = new System.Windows.Forms.Button();
            this.btnSubjectMoveUp = new System.Windows.Forms.Button();
            this.btnDeleteTemplate = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvBodyTemplate)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSubjectTemplate)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvBodyTemplate
            // 
            this.dgvBodyTemplate.AllowUserToAddRows = false;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgvBodyTemplate.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvBodyTemplate.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvBodyTemplate.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvBodyTemplate.BackgroundColor = System.Drawing.Color.Gainsboro;
            this.dgvBodyTemplate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvBodyTemplate.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgvBodyTemplate.Location = new System.Drawing.Point(448, 162);
            this.dgvBodyTemplate.Name = "dgvBodyTemplate";
            this.dgvBodyTemplate.Size = new System.Drawing.Size(890, 510);
            this.dgvBodyTemplate.TabIndex = 115;
            this.dgvBodyTemplate.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvBodyTemplate_CellClick);
            this.dgvBodyTemplate.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvBodyTemplate_CellValueChanged);
            // 
            // dgvSubjectTemplate
            // 
            this.dgvSubjectTemplate.AllowUserToAddRows = false;
            this.dgvSubjectTemplate.BackgroundColor = System.Drawing.Color.Gainsboro;
            this.dgvSubjectTemplate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSubjectTemplate.Location = new System.Drawing.Point(45, 89);
            this.dgvSubjectTemplate.Name = "dgvSubjectTemplate";
            this.dgvSubjectTemplate.Size = new System.Drawing.Size(397, 203);
            this.dgvSubjectTemplate.TabIndex = 116;
            this.dgvSubjectTemplate.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvSubjectTemplate_CellValueChanged);
            this.dgvSubjectTemplate.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dgvSubjectTemplate_KeyDown);
            // 
            // cboTemplates
            // 
            this.cboTemplates.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cboTemplates.FormattingEnabled = true;
            this.cboTemplates.Location = new System.Drawing.Point(153, 22);
            this.cboTemplates.Name = "cboTemplates";
            this.cboTemplates.Size = new System.Drawing.Size(809, 28);
            this.cboTemplates.TabIndex = 117;
            this.cboTemplates.TextChanged += new System.EventHandler(this.cboTemplates_TextChanged);
            // 
            // lblTemplate
            // 
            this.lblTemplate.AutoSize = true;
            this.lblTemplate.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTemplate.Location = new System.Drawing.Point(8, 25);
            this.lblTemplate.Name = "lblTemplate";
            this.lblTemplate.Size = new System.Drawing.Size(121, 20);
            this.lblTemplate.TabIndex = 118;
            this.lblTemplate.Text = "Template Name";
            // 
            // btnNewTemplate
            // 
            this.btnNewTemplate.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNewTemplate.Location = new System.Drawing.Point(979, 22);
            this.btnNewTemplate.Name = "btnNewTemplate";
            this.btnNewTemplate.Size = new System.Drawing.Size(116, 28);
            this.btnNewTemplate.TabIndex = 119;
            this.btnNewTemplate.Text = "New Template";
            this.btnNewTemplate.UseVisualStyleBackColor = true;
            this.btnNewTemplate.Click += new System.EventHandler(this.btnNewTemplate_Click);
            // 
            // btnSaveTemplate
            // 
            this.btnSaveTemplate.Location = new System.Drawing.Point(1101, 22);
            this.btnSaveTemplate.Name = "btnSaveTemplate";
            this.btnSaveTemplate.Size = new System.Drawing.Size(116, 28);
            this.btnSaveTemplate.TabIndex = 120;
            this.btnSaveTemplate.Text = "Save Template";
            this.btnSaveTemplate.UseVisualStyleBackColor = true;
            this.btnSaveTemplate.Click += new System.EventHandler(this.btnSaveTemplate_Click);
            // 
            // lblSubjectBuilder
            // 
            this.lblSubjectBuilder.AutoSize = true;
            this.lblSubjectBuilder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSubjectBuilder.Location = new System.Drawing.Point(49, 61);
            this.lblSubjectBuilder.Name = "lblSubjectBuilder";
            this.lblSubjectBuilder.Size = new System.Drawing.Size(98, 16);
            this.lblSubjectBuilder.TabIndex = 121;
            this.lblSubjectBuilder.Text = "Subject Builder";
            // 
            // lblBodyBuilder
            // 
            this.lblBodyBuilder.AutoSize = true;
            this.lblBodyBuilder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblBodyBuilder.Location = new System.Drawing.Point(444, 140);
            this.lblBodyBuilder.Name = "lblBodyBuilder";
            this.lblBodyBuilder.Size = new System.Drawing.Size(85, 16);
            this.lblBodyBuilder.TabIndex = 122;
            this.lblBodyBuilder.Text = "Body Builder";
            // 
            // btnAddSubjectRow
            // 
            this.btnAddSubjectRow.Location = new System.Drawing.Point(153, 56);
            this.btnAddSubjectRow.Name = "btnAddSubjectRow";
            this.btnAddSubjectRow.Size = new System.Drawing.Size(98, 27);
            this.btnAddSubjectRow.TabIndex = 123;
            this.btnAddSubjectRow.Text = "Add Subject Row";
            this.btnAddSubjectRow.UseVisualStyleBackColor = true;
            this.btnAddSubjectRow.Click += new System.EventHandler(this.btnAddSubjectRow_Click);
            // 
            // btnAddBodyRow
            // 
            this.btnAddBodyRow.Location = new System.Drawing.Point(535, 130);
            this.btnAddBodyRow.Name = "btnAddBodyRow";
            this.btnAddBodyRow.Size = new System.Drawing.Size(98, 27);
            this.btnAddBodyRow.TabIndex = 124;
            this.btnAddBodyRow.Text = "Add Body Row";
            this.btnAddBodyRow.UseVisualStyleBackColor = true;
            this.btnAddBodyRow.Click += new System.EventHandler(this.btnAddBodyRow_Click);
            // 
            // chkBodyDuplicateSettings
            // 
            this.chkBodyDuplicateSettings.AutoSize = true;
            this.chkBodyDuplicateSettings.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkBodyDuplicateSettings.Location = new System.Drawing.Point(649, 135);
            this.chkBodyDuplicateSettings.Name = "chkBodyDuplicateSettings";
            this.chkBodyDuplicateSettings.Size = new System.Drawing.Size(135, 20);
            this.chkBodyDuplicateSettings.TabIndex = 126;
            this.chkBodyDuplicateSettings.Text = "Duplicate Settings";
            this.chkBodyDuplicateSettings.UseVisualStyleBackColor = true;
            // 
            // lblBodySample
            // 
            this.lblBodySample.AutoSize = true;
            this.lblBodySample.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblBodySample.Location = new System.Drawing.Point(12, 302);
            this.lblBodySample.Name = "lblBodySample";
            this.lblBodySample.Size = new System.Drawing.Size(90, 16);
            this.lblBodySample.TabIndex = 130;
            this.lblBodySample.Text = "Body Sample";
            // 
            // lblSubjectSample
            // 
            this.lblSubjectSample.AutoSize = true;
            this.lblSubjectSample.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSubjectSample.Location = new System.Drawing.Point(443, 76);
            this.lblSubjectSample.Name = "lblSubjectSample";
            this.lblSubjectSample.Size = new System.Drawing.Size(103, 16);
            this.lblSubjectSample.TabIndex = 132;
            this.lblSubjectSample.Text = "Subject Sample";
            // 
            // rtxtBodySample
            // 
            this.rtxtBodySample.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.rtxtBodySample.BackColor = System.Drawing.SystemColors.ControlDark;
            this.rtxtBodySample.Font = new System.Drawing.Font("Courier New", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rtxtBodySample.ForeColor = System.Drawing.Color.Black;
            this.rtxtBodySample.Location = new System.Drawing.Point(12, 321);
            this.rtxtBodySample.Name = "rtxtBodySample";
            this.rtxtBodySample.Size = new System.Drawing.Size(428, 351);
            this.rtxtBodySample.TabIndex = 136;
            this.rtxtBodySample.TabStop = false;
            this.rtxtBodySample.Text = "";
            this.rtxtBodySample.WordWrap = false;
            this.rtxtBodySample.DoubleClick += new System.EventHandler(this.rtxtBodySample_DoubleClick);
            // 
            // txtSubjectSample
            // 
            this.txtSubjectSample.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSubjectSample.BackColor = System.Drawing.SystemColors.ControlDark;
            this.txtSubjectSample.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSubjectSample.Location = new System.Drawing.Point(552, 71);
            this.txtSubjectSample.Name = "txtSubjectSample";
            this.txtSubjectSample.Size = new System.Drawing.Size(786, 24);
            this.txtSubjectSample.TabIndex = 137;
            this.txtSubjectSample.DoubleClick += new System.EventHandler(this.txtSubjectSample_DoubleClick);
            // 
            // lblStaticIndentInfo
            // 
            this.lblStaticIndentInfo.AutoSize = true;
            this.lblStaticIndentInfo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblStaticIndentInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStaticIndentInfo.Location = new System.Drawing.Point(905, 105);
            this.lblStaticIndentInfo.Name = "lblStaticIndentInfo";
            this.lblStaticIndentInfo.Size = new System.Drawing.Size(433, 50);
            this.lblStaticIndentInfo.TabIndex = 138;
            this.lblStaticIndentInfo.Text = resources.GetString("lblStaticIndentInfo.Text");
            this.lblStaticIndentInfo.DoubleClick += new System.EventHandler(this.lblStaticIndentInfo_DoubleClick);
            // 
            // btnBodyMoveDown
            // 
            this.btnBodyMoveDown.Image = global::MWD_Template_Creator.Properties.Resources.Down_Arrow;
            this.btnBodyMoveDown.Location = new System.Drawing.Point(844, 110);
            this.btnBodyMoveDown.Name = "btnBodyMoveDown";
            this.btnBodyMoveDown.Size = new System.Drawing.Size(55, 47);
            this.btnBodyMoveDown.TabIndex = 140;
            this.btnBodyMoveDown.UseVisualStyleBackColor = true;
            this.btnBodyMoveDown.Click += new System.EventHandler(this.btnBodyMoveDown_Click);
            // 
            // btnBodyMoveUp
            // 
            this.btnBodyMoveUp.Image = global::MWD_Template_Creator.Properties.Resources.Up_Arrow;
            this.btnBodyMoveUp.Location = new System.Drawing.Point(780, 110);
            this.btnBodyMoveUp.Name = "btnBodyMoveUp";
            this.btnBodyMoveUp.Size = new System.Drawing.Size(58, 47);
            this.btnBodyMoveUp.TabIndex = 139;
            this.btnBodyMoveUp.UseVisualStyleBackColor = true;
            this.btnBodyMoveUp.Click += new System.EventHandler(this.btnBodyMoveUp_Click);
            // 
            // btnSubjectMoveDown
            // 
            this.btnSubjectMoveDown.Image = global::MWD_Template_Creator.Properties.Resources.Down_Arrow;
            this.btnSubjectMoveDown.Location = new System.Drawing.Point(4, 214);
            this.btnSubjectMoveDown.Name = "btnSubjectMoveDown";
            this.btnSubjectMoveDown.Size = new System.Drawing.Size(35, 49);
            this.btnSubjectMoveDown.TabIndex = 142;
            this.btnSubjectMoveDown.UseVisualStyleBackColor = true;
            this.btnSubjectMoveDown.Click += new System.EventHandler(this.btnSubjectMoveDown_Click);
            // 
            // btnSubjectMoveUp
            // 
            this.btnSubjectMoveUp.Image = global::MWD_Template_Creator.Properties.Resources.Up_Arrow;
            this.btnSubjectMoveUp.Location = new System.Drawing.Point(4, 110);
            this.btnSubjectMoveUp.Name = "btnSubjectMoveUp";
            this.btnSubjectMoveUp.Size = new System.Drawing.Size(35, 49);
            this.btnSubjectMoveUp.TabIndex = 141;
            this.btnSubjectMoveUp.UseVisualStyleBackColor = true;
            this.btnSubjectMoveUp.Click += new System.EventHandler(this.btnSubjectMoveUp_Click);
            // 
            // btnDeleteTemplate
            // 
            this.btnDeleteTemplate.Location = new System.Drawing.Point(1223, 22);
            this.btnDeleteTemplate.Name = "btnDeleteTemplate";
            this.btnDeleteTemplate.Size = new System.Drawing.Size(116, 28);
            this.btnDeleteTemplate.TabIndex = 143;
            this.btnDeleteTemplate.Text = "Delete Template";
            this.btnDeleteTemplate.UseVisualStyleBackColor = true;
            this.btnDeleteTemplate.Click += new System.EventHandler(this.btnDeleteTemplate_Click);
            // 
            // frmTemplates
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SlateGray;
            this.ClientSize = new System.Drawing.Size(1350, 675);
            this.Controls.Add(this.btnDeleteTemplate);
            this.Controls.Add(this.btnSubjectMoveDown);
            this.Controls.Add(this.btnSubjectMoveUp);
            this.Controls.Add(this.btnBodyMoveDown);
            this.Controls.Add(this.btnBodyMoveUp);
            this.Controls.Add(this.lblStaticIndentInfo);
            this.Controls.Add(this.txtSubjectSample);
            this.Controls.Add(this.rtxtBodySample);
            this.Controls.Add(this.lblSubjectSample);
            this.Controls.Add(this.lblBodySample);
            this.Controls.Add(this.chkBodyDuplicateSettings);
            this.Controls.Add(this.btnAddBodyRow);
            this.Controls.Add(this.btnAddSubjectRow);
            this.Controls.Add(this.lblBodyBuilder);
            this.Controls.Add(this.lblSubjectBuilder);
            this.Controls.Add(this.btnSaveTemplate);
            this.Controls.Add(this.btnNewTemplate);
            this.Controls.Add(this.lblTemplate);
            this.Controls.Add(this.cboTemplates);
            this.Controls.Add(this.dgvSubjectTemplate);
            this.Controls.Add(this.dgvBodyTemplate);
            this.Name = "frmTemplates";
            this.Text = "Templates";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmTemplates_FormClosing);
            this.Shown += new System.EventHandler(this.frmTemplates_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.dgvBodyTemplate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSubjectTemplate)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvBodyTemplate;
        private System.Windows.Forms.DataGridView dgvSubjectTemplate;
        private System.Windows.Forms.ComboBox cboTemplates;
        private System.Windows.Forms.Label lblTemplate;
        private System.Windows.Forms.Button btnNewTemplate;
        private System.Windows.Forms.Button btnSaveTemplate;
        private System.Windows.Forms.Label lblSubjectBuilder;
        private System.Windows.Forms.Label lblBodyBuilder;
        private System.Windows.Forms.Button btnAddSubjectRow;
        private System.Windows.Forms.Button btnAddBodyRow;
        private System.Windows.Forms.CheckBox chkBodyDuplicateSettings;
        private System.Windows.Forms.Label lblBodySample;
        private System.Windows.Forms.Label lblSubjectSample;
        private System.Windows.Forms.RichTextBox rtxtBodySample;
        private System.Windows.Forms.TextBox txtSubjectSample;
        private System.Windows.Forms.Label lblStaticIndentInfo;
        private System.Windows.Forms.Button btnBodyMoveUp;
        private System.Windows.Forms.Button btnBodyMoveDown;
        private System.Windows.Forms.Button btnSubjectMoveDown;
        private System.Windows.Forms.Button btnSubjectMoveUp;
        private System.Windows.Forms.Button btnDeleteTemplate;
    }
}