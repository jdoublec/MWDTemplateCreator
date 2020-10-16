using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Text;

namespace MWD_Template_Creator
{
    public partial class frmTemplates : Form
    {
        private DataSet dsTemplates;
        private DataTable dtSubjectTemplate;
        private DataTable dtBodyTemplate;
        private string[] strarrDataSource;
        private int[] iarrFontSize;
        private string strTemplateFilePath;
        private bool bTemplatesChanged;
        private bool bUnsavedChanges;

        private string GetTemplateFilePath
        {
            get
            {
                return strTemplateFilePath;
            }
            set
            {
                strTemplateFilePath = value;
            }
        }

        public bool AreTemplatesChanged
        {
            get
            {
                return bTemplatesChanged;
            }
            set
            {
                bTemplatesChanged = value;
            }
        }

        public DataSet GetBodyTemplateDataSet
        {
            get
            {
                return dsTemplates; 
            }
        }

        public frmTemplates(string strCurrentTemplate, string strFilePath)
        {
            InitializeComponent();

            //this.MaximizeBox = false;

            GetTemplateFilePath = strFilePath;

            bUnsavedChanges = false;

            strarrDataSource = new string[] { "", "Activity", "Beyond_Bit_A/B_Plan", "Beyond_Bit_A/B_Target_Line", "Beyond_Bit_Azm", "Beyond_Bit_DLS", "Beyond_Bit_East",
                "Beyond_Bit_Header", "Beyond_Bit_Inc", "Beyond_Bit_MD", "Beyond_Bit_North", "Beyond_Bit_R/L_Plan", "Beyond_Bit_Tool_Face", "Beyond_Bit_TVD",
                "Beyond_Bit_VS", "Bit_A/B_Plan", "Bit_A/B_Target_Line", "Bit_Azm", "Bit_DLS", "Bit_East", "Bit_Inc", "Bit_MD", "Bit_North", "Bit_R/L_Plan", "Bit_RPM", "Bit_Tool_Face",
                "Bit_TVD", "Bit_VS", "Client", "CM_Temp", "Comments", "County", "Differential_Pressure", "DLS_Needed", "ECD", "ESD", "Flow_Rate", "Gamma", "Job_Number",
                "Motor_RPM", "Motor_Yield", "Mud_Weight", "Pick_Up", "Plan_Number", "Rev/Gal", "Rig", "ROP_Rotating", "ROP_Sliding", "Rotary_Torque", "Rotate_Weight",
                "Slack_Off", "Slide_Ahead", "Slide_Rotate_Footage", "Slide_Seen", "Slide_Tool_Face", "SPP", "Surface_RPM", "Survey_A/B_Plan", "Survey_A/B_Target_Line",
                "Survey_Azm", "Survey_DLS", "Survey_East", "Survey_Inc", "Survey_MD", "Survey_North", "Survey_R/L_Plan", "Survey_Temp", "Survey_Tool_Face", "Survey_TVD",
                "Survey_VS", "Target_Azm", "Target_Inc", "Tool_Face_Mode", "Well_Name", "WOB"};            

            iarrFontSize = new int[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };

            cboTemplates.DropDownStyle = ComboBoxStyle.DropDownList;
            cboTemplates.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboTemplates.AutoCompleteMode = AutoCompleteMode.Suggest;

            dsTemplates = new DataSet("Templates");
            dtSubjectTemplate = new DataTable("SubjectTemplate");
            dtBodyTemplate = new DataTable("BodyTemplate");            
            dsTemplates.Tables.Add(CreateSubjectTemplateTable());
            dsTemplates.Tables.Add(CreateBodyTemplateTable());
            CreateSubjectTemplateDGV();
            CreateBodyTemplateDGV();

            if (File.Exists(GetTemplateFilePath))
            {
                dsTemplates.ReadXml(GetTemplateFilePath);
                dsTemplates.AcceptChanges();
            }
            else
            {
                dsTemplates.WriteXml(GetTemplateFilePath);
                dsTemplates.AcceptChanges();
            }            

            for (int i = 0; i < dtSubjectTemplate.Rows.Count; i++)
            {
                if (!cboTemplates.Items.Contains(dtSubjectTemplate.Rows[i]["Name"].ToString()))
                {
                    cboTemplates.Items.Add(dtSubjectTemplate.Rows[i]["Name"].ToString());
                }
            }

            //for (int i = 0; i < dtBodyTemplate.Rows.Count; i++)
            //{
            //    dtBodyTemplate.Rows[i]["LineHeaderFontColor"] = -16777216;
            //    dtBodyTemplate.Rows[i]["DataSourceFontColor"] = -16777216;
            //}

            dgvSubjectTemplate.DataSource = dtSubjectTemplate;
            dgvBodyTemplate.DataSource = dtBodyTemplate;

            cboTemplates.Text = strCurrentTemplate;

            //if (strCurrentTemplate.Replace(" ", "").ToString() != "")
            //{
            //    dtSubjectTemplate.DefaultView.RowFilter = "Name = '" + strCurrentTemplate + "'";
                
            //    dgvBodyTemplate.Refresh();
            //    //ChangeBodyTemplateDefaultViewRowFilter(strCurrentTemplate);
            //}                       

            AreTemplatesChanged = false;
        }

        private void frmTemplates_Shown(object sender, EventArgs e)
        {
            dgvSubjectTemplate.Columns["Name"].Visible = false;
            dgvBodyTemplate.Columns["Name"].Visible = false;

            ChangeBodyTemplateDefaultViewRowFilter(cboTemplates.Text);            
        }

        private DataTable CreateSubjectTemplateTable()
        {
            dtSubjectTemplate.Columns.Add("Name", typeof(string));
            dtSubjectTemplate.Columns.Add("DataSource", typeof(string));            
            dtSubjectTemplate.Columns.Add("Separator", typeof(string));    

            return dtSubjectTemplate;
        }

        private void AddSubjectTemplateRow(string strName)
        {
            int iInsertIndex = dgvSubjectTemplate.Rows.Count;

            if (dgvSubjectTemplate.SelectedRows.Count == 1)
            {
                for (int i = 0; i < dtSubjectTemplate.Rows.Count; i++)
                {
                    if ((dtSubjectTemplate.Rows[i]["Name"].ToString() == cboTemplates.Text) &&
                        (dtSubjectTemplate.Rows[i]["DataSource"].ToString() == dgvSubjectTemplate.SelectedRows[0].Cells["DataSource"].Value.ToString()))
                    {
                        iInsertIndex = i + 1;
                        break;
                    }
                }
            }

            DataRow drRow = dtSubjectTemplate.NewRow();

            drRow["Name"] = strName;
            drRow["DataSource"] = "";
            drRow["Separator"] = "";

            dtSubjectTemplate.Rows.InsertAt(drRow, iInsertIndex);
        }

        private DataTable CreateBodyTemplateTable()
        {
            dtBodyTemplate.Columns.Add("Name", typeof(string));
            dtBodyTemplate.Columns.Add("LineHeader", typeof(string));
            dtBodyTemplate.Columns.Add("LineHeaderFontName", typeof(string));
            dtBodyTemplate.Columns.Add("LineHeaderFontSize", typeof(int));
            dtBodyTemplate.Columns.Add("LineHeaderFontColor", typeof(int));
            dtBodyTemplate.Columns.Add("LineHeaderUnderlined", typeof(bool));
            dtBodyTemplate.Columns.Add("LineHeaderBold", typeof(bool));
            dtBodyTemplate.Columns.Add("LineHeaderItalic", typeof(bool));
            dtBodyTemplate.Columns.Add("Spacing", typeof(string));
            dtBodyTemplate.Columns.Add("DataSource", typeof(string));
            dtBodyTemplate.Columns.Add("Unit", typeof(string));
            dtBodyTemplate.Columns.Add("DataSourceFontName", typeof(string));
            dtBodyTemplate.Columns.Add("DataSourceFontSize", typeof(int));
            dtBodyTemplate.Columns.Add("DataSourceFontColor", typeof(int));
            dtBodyTemplate.Columns.Add("DataSourceUnderlined", typeof(bool));
            dtBodyTemplate.Columns.Add("DataSourceBold", typeof(bool));
            dtBodyTemplate.Columns.Add("DataSourceItalic", typeof(bool));
            dtBodyTemplate.Columns.Add("BlankLines", typeof(int));
            dtBodyTemplate.Columns.Add("StaticIndent", typeof(int));
            dtBodyTemplate.Columns.Add("TableRowBackColor", typeof(int));

            return dtBodyTemplate;
        }

        private void AddBodyTemplateRow(string strName,
                                        string strLineHeaderFontName, int iLineHeaderFontSize, int iLineHeaderFontColor, bool bLineHeaderUnderlined, bool bLineHeaderBold, bool bLineHeaderItalic,
                                        string strDataSourceFontName, int iDataSourceFontSize, int iDataSourceFontColor, bool bDataSourceUnderlined, bool bDataSourceBold, bool bDataSourceItalic,
                                        string strSpacing, int iBlankLines, int iStaticIndent, int iTableRowBackColor)
        {
            int iInsertIndex = dtBodyTemplate.Rows.Count;

            if (dgvBodyTemplate.SelectedRows.Count == 1)
            {
                if (chkBodyDuplicateSettings.Checked)
                {
                    for (int i = 0; i < dtBodyTemplate.Rows.Count; i++)
                    {
                        if ((dtBodyTemplate.Rows[i]["Name"].ToString() == cboTemplates.Text) &&
                            (dtBodyTemplate.Rows[i]["LineHeader"].ToString() == dgvBodyTemplate.SelectedRows[0].Cells["LineHeader"].Value.ToString()) &&
                            (dtBodyTemplate.Rows[i]["DataSource"].ToString() == dgvBodyTemplate.SelectedRows[0].Cells["DataSource"].Value.ToString()))
                        {
                            iInsertIndex = i + 1;
                            break;
                        }
                    }
                }
                else
                {
                    for (int i = dtBodyTemplate.Rows.Count - 1; i >= 0; i--)
                    {
                        if (dtBodyTemplate.Rows[i]["Name"].ToString() == dgvBodyTemplate["Name", 0].Value.ToString())
                        {
                            iInsertIndex = i + 1;
                            break;
                        }
                    }
                }
            }


            DataRow drRow = dtBodyTemplate.NewRow();

            drRow["Name"] = strName;
            drRow["LineHeader"] = "";
            drRow["LineHeaderFontName"] = strLineHeaderFontName;
            drRow["LineHeaderFontSize"] = iLineHeaderFontSize;
            drRow["LineHeaderFontColor"] = iLineHeaderFontColor;
            drRow["LineHeaderUnderlined"] = bLineHeaderUnderlined;
            drRow["LineHeaderBold"] = bLineHeaderBold;
            drRow["LineHeaderItalic"] = bLineHeaderItalic;
            drRow["Spacing"] = strSpacing;
            drRow["DataSource"] = "";
            drRow["Unit"] = "";
            drRow["DataSourceFontName"] = strDataSourceFontName;
            drRow["DataSourceFontSize"] = iDataSourceFontSize;
            drRow["DataSourceFontColor"] = iDataSourceFontColor;
            drRow["DataSourceUnderlined"] = bDataSourceUnderlined;
            drRow["DataSourceBold"] = bDataSourceBold;
            drRow["DataSourceItalic"] = bDataSourceItalic;
            drRow["BlankLines"] = iBlankLines;
            drRow["StaticIndent"] = iStaticIndent;
            drRow["TableRowBackColor"] = iTableRowBackColor;

            
            dtBodyTemplate.Rows.InsertAt(drRow, iInsertIndex);
        }

        private void CreateSubjectTemplateDGV()
        {
            dgvSubjectTemplate.Columns.Add("Name", "Name");
            dgvSubjectTemplate.Columns["Name"].DataPropertyName = "Name";
            dgvSubjectTemplate.Columns["Name"].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubjectTemplate.Columns["Name"].ReadOnly = true;

            DataGridViewComboBoxColumn cbocDataSource = new DataGridViewComboBoxColumn();
            cbocDataSource.HeaderText = "Data Source";
            cbocDataSource.DataPropertyName = "DataSource";            
            cbocDataSource.Name = "DataSource";
            cbocDataSource.Items.AddRange(strarrDataSource);
            dgvSubjectTemplate.Columns.Add(cbocDataSource);

            dgvSubjectTemplate.Columns.Add("Separator", "Separator");
            dgvSubjectTemplate.Columns["Separator"].DataPropertyName = "Separator";
            dgvSubjectTemplate.Columns["Separator"].SortMode = DataGridViewColumnSortMode.NotSortable;            
        }

        private void CreateBodyTemplateDGV()
        {
            dgvBodyTemplate.Columns.Add("Name", "Name");
            dgvBodyTemplate.Columns["Name"].DataPropertyName = "Name";
            dgvBodyTemplate.Columns["Name"].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvBodyTemplate.Columns["Name"].ReadOnly = true;            

            dgvBodyTemplate.Columns.Add("LineHeader", "Line Header");
            dgvBodyTemplate.Columns["LineHeader"].DataPropertyName = "LineHeader";
            dgvBodyTemplate.Columns["LineHeader"].SortMode = DataGridViewColumnSortMode.NotSortable;

            DataGridViewButtonColumn btncLineHeaderFontName = new DataGridViewButtonColumn();
            btncLineHeaderFontName.HeaderText = "Line Header Font Name";
            btncLineHeaderFontName.DataPropertyName = "LineHeaderFontName";
            btncLineHeaderFontName.Name = "LineHeaderFontName";
            btncLineHeaderFontName.Text = "Courier New";
            btncLineHeaderFontName.UseColumnTextForButtonValue = false;
            dgvBodyTemplate.Columns.Add(btncLineHeaderFontName);

            DataGridViewComboBoxColumn cbocLineHeaderFontSize = new DataGridViewComboBoxColumn();
            cbocLineHeaderFontSize.HeaderText = "Line Header Font Size";
            cbocLineHeaderFontSize.DataPropertyName = "LineHeaderFontSize";
            cbocLineHeaderFontSize.Name = "LineHeaderFontSize";
            for (int i = 0; i < iarrFontSize.Length; i++)
            {
                cbocLineHeaderFontSize.Items.Add(iarrFontSize[i]);
            }
            dgvBodyTemplate.Columns.Add(cbocLineHeaderFontSize);

            DataGridViewButtonColumn btncLineHeaderFontColor = new DataGridViewButtonColumn();
            btncLineHeaderFontColor.HeaderText = "LH Color";
            btncLineHeaderFontColor.DataPropertyName = "LineHeaderFontColor";
            btncLineHeaderFontColor.Name = "LineHeaderFontColor";
            btncLineHeaderFontColor.Text = "Courier New";
            btncLineHeaderFontColor.UseColumnTextForButtonValue = false;
            btncLineHeaderFontColor.FlatStyle = FlatStyle.Popup;
            btncLineHeaderFontColor.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            dgvBodyTemplate.Columns.Add(btncLineHeaderFontColor);

            DataGridViewCheckBoxColumn chkcLineHeaderUnderlined = new DataGridViewCheckBoxColumn();
            chkcLineHeaderUnderlined.HeaderText = "U";
            chkcLineHeaderUnderlined.DataPropertyName = "LineHeaderUnderlined";
            chkcLineHeaderUnderlined.Name = "LineHeaderUnderlined";
            //chkcLineHeaderUnderlined.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            dgvBodyTemplate.Columns.Add(chkcLineHeaderUnderlined);

            DataGridViewCheckBoxColumn chkcLineHeaderBold = new DataGridViewCheckBoxColumn();
            chkcLineHeaderBold.HeaderText = "B";
            chkcLineHeaderBold.DataPropertyName = "LineHeaderBold";
            chkcLineHeaderBold.Name = "LineHeaderBold";
            //chkcLineHeaderBold.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            dgvBodyTemplate.Columns.Add(chkcLineHeaderBold);

            DataGridViewCheckBoxColumn chkcLineHeaderItalic = new DataGridViewCheckBoxColumn();
            chkcLineHeaderItalic.HeaderText = "I";
            chkcLineHeaderItalic.DataPropertyName = "LineHeaderItalic";
            chkcLineHeaderItalic.Name = "LineHeaderItalic";
            //chkcLineHeaderItalic.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            dgvBodyTemplate.Columns.Add(chkcLineHeaderItalic);

            DataGridViewComboBoxColumn cbocSpacing = new DataGridViewComboBoxColumn();
            cbocSpacing.HeaderText = "Spacing";
            cbocSpacing.DataPropertyName = "Spacing";
            cbocSpacing.Name = "Spacing";
            cbocSpacing.Items.Add("Static Indent");
            for (int i = 0; i <= 20; i++)
            {
                cbocSpacing.Items.Add(i.ToString());
            }
            dgvBodyTemplate.Columns.Add(cbocSpacing);

            DataGridViewComboBoxColumn cbocDataSource = new DataGridViewComboBoxColumn();
            cbocDataSource.HeaderText = "Data Source";
            cbocDataSource.DataPropertyName = "DataSource";
            cbocDataSource.Name = "DataSource";
            cbocDataSource.Items.AddRange(strarrDataSource);

            dgvBodyTemplate.Columns.Add(cbocDataSource); dgvBodyTemplate.Columns.Add("Unit", "Unit");
            dgvBodyTemplate.Columns["Unit"].DataPropertyName = "Unit";
            dgvBodyTemplate.Columns["Unit"].SortMode = DataGridViewColumnSortMode.NotSortable;

            DataGridViewButtonColumn btncDataSourceFontName = new DataGridViewButtonColumn();
            btncDataSourceFontName.HeaderText = "Data Source Font Name";
            btncDataSourceFontName.DataPropertyName = "DataSourceFontName";
            btncDataSourceFontName.Name = "DataSourceFontName";
            btncDataSourceFontName.Text = "Courier New";
            dgvBodyTemplate.Columns.Add(btncDataSourceFontName);

            DataGridViewComboBoxColumn cbocDataSourceFontSize = new DataGridViewComboBoxColumn();
            cbocDataSourceFontSize.HeaderText = "Data Source Font Size";
            cbocDataSourceFontSize.DataPropertyName = "DataSourceFontSize";
            cbocDataSourceFontSize.Name = "DataSourceFontSize"; for (int i = 0; i < iarrFontSize.Length; i++)
            {
                cbocDataSourceFontSize.Items.Add(iarrFontSize[i]);
            }
            dgvBodyTemplate.Columns.Add(cbocDataSourceFontSize);

            DataGridViewButtonColumn btncDataSourceFontColor = new DataGridViewButtonColumn();
            btncDataSourceFontColor.HeaderText = "DS Color";
            btncDataSourceFontColor.DataPropertyName = "DataSourceFontColor";
            btncDataSourceFontColor.Name = "DataSourceFontColor";
            btncDataSourceFontColor.Text = "Courier New";
            btncDataSourceFontColor.FlatStyle = FlatStyle.Popup;
            btncDataSourceFontColor.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            dgvBodyTemplate.Columns.Add(btncDataSourceFontColor);

            DataGridViewCheckBoxColumn chkcDataSourceUnderlined = new DataGridViewCheckBoxColumn();
            chkcDataSourceUnderlined.HeaderText = "U";
            chkcDataSourceUnderlined.DataPropertyName = "DataSourceUnderlined";
            chkcDataSourceUnderlined.Name = "DataSourceUnderlined";
            //chkcDataSourceUnderlined.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            dgvBodyTemplate.Columns.Add(chkcDataSourceUnderlined);

            DataGridViewCheckBoxColumn chkcDataSourceBold = new DataGridViewCheckBoxColumn();
            chkcDataSourceBold.HeaderText = "B";
            chkcDataSourceBold.DataPropertyName = "DataSourceBold";
            chkcDataSourceBold.Name = "DataSourceBold";
            //chkcDataSourceBold.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            dgvBodyTemplate.Columns.Add(chkcDataSourceBold);

            DataGridViewCheckBoxColumn chkcDataSourceItalic = new DataGridViewCheckBoxColumn();
            chkcDataSourceItalic.HeaderText = "I";
            chkcDataSourceItalic.DataPropertyName = "DataSourceItalic";
            chkcDataSourceItalic.Name = "DataSourceItalic";
            //chkcDataSourceItalic.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            dgvBodyTemplate.Columns.Add(chkcDataSourceItalic);

            DataGridViewComboBoxColumn cbocBlankLines = new DataGridViewComboBoxColumn();
            cbocBlankLines.HeaderText = "Blank Lines";
            cbocBlankLines.DataPropertyName = "BlankLines";
            cbocBlankLines.Name = "BlankLines";
            for (int i = 0; i <= 10; i++)
            {
                cbocBlankLines.Items.AddRange(i);
            }            
            dgvBodyTemplate.Columns.Add(cbocBlankLines);

            dgvBodyTemplate.Columns.Add("StaticIndent", "Static Indent");
            dgvBodyTemplate.Columns["StaticIndent"].DataPropertyName = "StaticIndent";
            dgvBodyTemplate.Columns["StaticIndent"].SortMode = DataGridViewColumnSortMode.NotSortable;

            DataGridViewButtonColumn btncTableRowBackColor = new DataGridViewButtonColumn();
            btncTableRowBackColor.HeaderText = "Table Row Color";
            btncTableRowBackColor.DataPropertyName = "TableRowBackColor";
            btncTableRowBackColor.Name = "TableRowBackColor";
            btncTableRowBackColor.Text = "Courier New";
            btncTableRowBackColor.UseColumnTextForButtonValue = false;
            btncTableRowBackColor.FlatStyle = FlatStyle.Popup;
            btncTableRowBackColor.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            dgvBodyTemplate.Columns.Add(btncTableRowBackColor);
        }

        private void btnNewTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                frmTemplateName frmName = new frmTemplateName();

                frmName.ShowDialog();

                bool bOKToAdd = true;

                if (frmName.OKToAdd)
                {
                    for (int i = 0; i < dtSubjectTemplate.Rows.Count; i++)
                    {
                        if (frmName.GetTemplateName == dtSubjectTemplate.Rows[i]["Name"].ToString())
                        {
                            MessageBox.Show("Duplicate Template Names are not permitted.");
                            bOKToAdd = false;
                            break;
                        }
                    }

                    if (bOKToAdd)
                    {
                        if (cboTemplates.Items.Count > 0)
                        {
                            DialogResult drResult = MessageBox.Show("Would you like to duplicate a Current Template?", "Duplicate Template?", MessageBoxButtons.YesNo);

                            if (drResult == DialogResult.Yes)
                            {
                                List<string> lstTemplates = new List<string>();
                                for (int i = 0; i < cboTemplates.Items.Count; i++)
                                {
                                    lstTemplates.Add(cboTemplates.Items[i].ToString());
                                }

                                frmTemplateName frmDupTemplate = new frmTemplateName(lstTemplates);

                                frmDupTemplate.ShowDialog();

                                if (frmDupTemplate.OKToAdd)
                                {
                                    string strTemplateToDuplicate = frmDupTemplate.GetTemplateToDuplicate;
                                    string strNewTemplateName = frmName.GetTemplateName;

                                    int iSubjectTemplateRows = dtSubjectTemplate.Rows.Count;
                                    for (int i = 0; i < iSubjectTemplateRows; i++)
                                    {
                                        if (dtSubjectTemplate.Rows[i]["Name"].ToString() == strTemplateToDuplicate)
                                        {
                                            DataRow drRow = dtSubjectTemplate.NewRow();

                                            drRow["Name"] = strNewTemplateName;
                                            drRow["DataSource"] = dtSubjectTemplate.Rows[i]["DataSource"].ToString();
                                            drRow["Separator"] = dtSubjectTemplate.Rows[i]["Separator"].ToString();

                                            dtSubjectTemplate.Rows.Add(drRow);
                                        }
                                    }

                                    int iBodyTemplateRows = dtBodyTemplate.Rows.Count;
                                    for (int i = 0; i < iBodyTemplateRows; i++)
                                    {
                                        if (dtBodyTemplate.Rows[i]["Name"].ToString() == strTemplateToDuplicate)
                                        {
                                            DataRow drRow = dtBodyTemplate.NewRow();

                                            drRow["Name"] = strNewTemplateName;
                                            drRow["LineHeader"] = dtBodyTemplate.Rows[i]["LineHeader"].ToString();
                                            drRow["LineHeaderFontName"] = dtBodyTemplate.Rows[i]["LineHeaderFontName"].ToString();
                                            drRow["LineHeaderFontSize"] = Convert.ToInt16(dtBodyTemplate.Rows[i]["LineHeaderFontSize"]);
                                            drRow["LineHeaderFontColor"] = Convert.ToInt32(dtBodyTemplate.Rows[i]["LineHeaderFontColor"]);
                                            drRow["LineHeaderUnderlined"] = Convert.ToBoolean(dtBodyTemplate.Rows[i]["LineHeaderUnderlined"]);
                                            drRow["LineHeaderBold"] = Convert.ToBoolean(dtBodyTemplate.Rows[i]["LineHeaderBold"]);
                                            drRow["LineHeaderItalic"] = Convert.ToBoolean(dtBodyTemplate.Rows[i]["LineHeaderItalic"]); 
                                            drRow["Spacing"] = dtBodyTemplate.Rows[i]["Spacing"].ToString();
                                            drRow["DataSource"] = dtBodyTemplate.Rows[i]["DataSource"].ToString();
                                            drRow["Unit"] = dtBodyTemplate.Rows[i]["Unit"].ToString();
                                            drRow["DataSourceFontName"] = dtBodyTemplate.Rows[i]["DataSourceFontName"].ToString();
                                            drRow["DataSourceFontSize"] = Convert.ToInt16(dtBodyTemplate.Rows[i]["DataSourceFontSize"]);
                                            drRow["DataSourceFontColor"] = Convert.ToInt32(dtBodyTemplate.Rows[i]["DataSourceFontColor"]);
                                            drRow["DataSourceUnderlined"] = Convert.ToBoolean(dtBodyTemplate.Rows[i]["DataSourceUnderlined"]);
                                            drRow["DataSourceBold"] = Convert.ToBoolean(dtBodyTemplate.Rows[i]["DataSourceBold"]);
                                            drRow["DataSourceItalic"] = Convert.ToBoolean(dtBodyTemplate.Rows[i]["DataSourceItalic"]);
                                            drRow["BlankLines"] = Convert.ToInt16(dtBodyTemplate.Rows[i]["BlankLines"]);
                                            drRow["StaticIndent"] = Convert.ToInt16(dtBodyTemplate.Rows[i]["StaticIndent"]);
                                            drRow["TableRowBackColor"] = Convert.ToInt32(dtBodyTemplate.Rows[i]["TableRowBackColor"]);

                                            dtBodyTemplate.Rows.Add(drRow);
                                        }
                                    }

                                    cboTemplates.Items.Add(frmName.GetTemplateName);
                                    cboTemplates.Text = frmName.GetTemplateName;
                                    bUnsavedChanges = true;
                                }
                            }
                        }
                        else
                        {
                            AddSubjectTemplateRow(frmName.GetTemplateName);
                            AddBodyTemplateRow(frmName.GetTemplateName, "Courier New", 10, Color.Black.ToArgb(), false, false, false,
                                               "Courier New", 10, Color.Black.ToArgb(), false, false, false, "0", 0, 0, Color.White.ToArgb());
                            cboTemplates.Items.Add(frmName.GetTemplateName);
                            cboTemplates.Text = frmName.GetTemplateName;
                            bUnsavedChanges = true;
                        }
                    }                
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void dgvBodyTemplate_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (((e.ColumnIndex == dgvBodyTemplate.Columns["LineHeaderFontName"].Index) || (e.ColumnIndex == dgvBodyTemplate.Columns["DataSourceFontName"].Index)) && e.RowIndex >= 0)
            {
                FontDialog fdFont = new FontDialog();
                DialogResult drResult = fdFont.ShowDialog();

                if (drResult == DialogResult.OK)
                {
                    if ((dtBodyTemplate.DefaultView[e.RowIndex]["Spacing"].ToString() != "Static Indent") ||
                        ((dtBodyTemplate.DefaultView[e.RowIndex]["Spacing"].ToString() == "Static Indent") &&
                        ((fdFont.Font.Name.ToString() == "Courier New") || (fdFont.Font.Name.ToString() == "Consolas"))))
                    {
                        //FontConverter fConverter = new FontConverter();
                        dtBodyTemplate.DefaultView[e.RowIndex][e.ColumnIndex] = fdFont.Font.Name.ToString();// fConverter.ConvertToString(fdFont.Font);
                        dtBodyTemplate.DefaultView[e.RowIndex][e.ColumnIndex + 1] = fdFont.Font.Size;
                        dtBodyTemplate.DefaultView[e.RowIndex][e.ColumnIndex + 3] = fdFont.Font.Underline;
                        dtBodyTemplate.DefaultView[e.RowIndex][e.ColumnIndex + 4] = fdFont.Font.Bold;
                        dtBodyTemplate.DefaultView[e.RowIndex][e.ColumnIndex + 5] = fdFont.Font.Italic;                        

                        if (dtBodyTemplate.DefaultView[e.RowIndex]["Spacing"].ToString() == "Static Indent")
                        {
                            if (dtBodyTemplate.DefaultView[e.RowIndex]["LineHeaderFontName"].ToString() != 
                                dtBodyTemplate.DefaultView[e.RowIndex]["DataSourceFontName"].ToString())
                            {
                                dtBodyTemplate.DefaultView[e.RowIndex]["DataSourceFontName"] = 
                                dtBodyTemplate.DefaultView[e.RowIndex]["LineHeaderFontName"].ToString();
                            }
                        }

                        UpdateBodyRichTextBox();
                        UpdateBodyRichTextBoxFont();
                        dgvBodyTemplate.Refresh();
                    }
                    else
                    {
                        MessageBox.Show("'Static Indent' can only be used with a monospace font.  Use 'Courier New' or 'Consolas'.");
                    }
                }
            }
            else if (((e.ColumnIndex == dgvBodyTemplate.Columns["LineHeaderFontColor"].Index) || (e.ColumnIndex == dgvBodyTemplate.Columns["DataSourceFontColor"].Index)) && e.RowIndex >= 0)
            {
                ColorDialog cdColor = new ColorDialog();
                DialogResult drResult = cdColor.ShowDialog();

                if (drResult == DialogResult.OK)
                {
                    //dgvBodyTemplate.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = cdColor.Color;
                    //dgvBodyTemplate.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = cdColor.Color;
                    dgvBodyTemplate.SelectedCells[0].Style.BackColor = cdColor.Color;
                    dgvBodyTemplate.SelectedCells[0].Style.ForeColor = cdColor.Color;
                    dgvBodyTemplate.SelectedCells[0].Style.SelectionBackColor = cdColor.Color;
                    dgvBodyTemplate.SelectedCells[0].Style.SelectionForeColor = cdColor.Color;
                    dtBodyTemplate.DefaultView[e.RowIndex][e.ColumnIndex] = cdColor.Color.ToArgb();

                    dgvBodyTemplate.Refresh();
                    UpdateBodyRichTextBox();
                    UpdateBodyRichTextBoxFont();

                }
            }
            else if (e.ColumnIndex == dgvBodyTemplate.Columns["TableRowBackColor"].Index)
            {
                ColorDialog cdColor = new ColorDialog();
                DialogResult drResult = cdColor.ShowDialog();

                if (drResult == DialogResult.OK)
                {
                    dgvBodyTemplate.SelectedCells[0].Style.BackColor = cdColor.Color;
                    dgvBodyTemplate.SelectedCells[0].Style.ForeColor = cdColor.Color;
                    dgvBodyTemplate.SelectedCells[0].Style.SelectionBackColor = cdColor.Color;
                    dgvBodyTemplate.SelectedCells[0].Style.SelectionForeColor = cdColor.Color;
                    dtBodyTemplate.DefaultView[e.RowIndex][e.ColumnIndex] = cdColor.Color.ToArgb();

                    dgvBodyTemplate.Refresh();
                }
            }
        }

        private void btnSaveTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                dtSubjectTemplate.AcceptChanges();
                dtBodyTemplate.AcceptChanges();
                dsTemplates.WriteXml(GetTemplateFilePath);

                ChangeBodyTemplateDefaultViewRowFilter(cboTemplates.Text);

                AreTemplatesChanged = true;
                bUnsavedChanges = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void frmTemplates_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (bUnsavedChanges)
            {
                DialogResult drResult = MessageBox.Show("Would you like to save changes before closing.  " +
                                        "Click 'Yes' to save, 'No' to close without saving, or 'Cancel' to keep " +
                                        "the 'Templates Form' open.", "Save, Close, or Continue Editing", MessageBoxButtons.YesNoCancel);

                if (drResult == DialogResult.Yes)
                {
                    dtSubjectTemplate.AcceptChanges();
                    dtBodyTemplate.AcceptChanges();
                    dsTemplates.WriteXml(GetTemplateFilePath);
                }
                else if (drResult == DialogResult.No)
                {
                    dtSubjectTemplate.RejectChanges();
                    dtBodyTemplate.RejectChanges();
                }
                else if (drResult == DialogResult.Cancel)
                {
                    e.Cancel = true;
                }
            }
        }

        private void cboTemplates_TextChanged(object sender, EventArgs e)
        {
            dtSubjectTemplate.DefaultView.RowFilter = "Name = '" + cboTemplates.Text + "'";
            ChangeBodyTemplateDefaultViewRowFilter(cboTemplates.Text);
            rtxtBodySample.Text = "";
            txtSubjectSample.Text = "";
        }

        private void btnAddSubjectRow_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(cboTemplates.Text.Replace(" ", "")))
            {
                AddSubjectTemplateRow(cboTemplates.Text);
                
                bUnsavedChanges = true;
            }
            else
            {
                MessageBox.Show("A Template must first be created or selected before adding a Subject Row.");
            }
        }

        private void ChangeBodyTemplateDefaultViewRowFilter(string strFilter)
        {
            try
            {
                dtBodyTemplate.DefaultView.RowFilter = "";
                dtBodyTemplate.DefaultView.RowFilter = "Name = '" + strFilter + "'";
                for (int i = 0; i < dtBodyTemplate.DefaultView.Count; i++)
                {
                    dgvBodyTemplate["LineHeaderFontColor", i].Selected = true;
                    dgvBodyTemplate.SelectedCells[0].Style.BackColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderFontColor"]));
                    dgvBodyTemplate.SelectedCells[0].Style.ForeColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderFontColor"]));
                    dgvBodyTemplate.SelectedCells[0].Style.SelectionBackColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderFontColor"]));
                    dgvBodyTemplate.SelectedCells[0].Style.SelectionForeColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderFontColor"]));
                    dgvBodyTemplate["DataSourceFontColor", i].Selected = true;
                    dgvBodyTemplate.SelectedCells[0].Style.BackColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]));
                    dgvBodyTemplate.SelectedCells[0].Style.ForeColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]));
                    dgvBodyTemplate.SelectedCells[0].Style.SelectionBackColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]));
                    dgvBodyTemplate.SelectedCells[0].Style.SelectionForeColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]));
                    dgvBodyTemplate["TableRowBackColor", i].Selected = true;
                    dgvBodyTemplate.SelectedCells[0].Style.BackColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["TableRowBackColor"]));
                    dgvBodyTemplate.SelectedCells[0].Style.ForeColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["TableRowBackColor"]));
                    dgvBodyTemplate.SelectedCells[0].Style.SelectionBackColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["TableRowBackColor"]));
                    dgvBodyTemplate.SelectedCells[0].Style.SelectionForeColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["TableRowBackColor"]));
                }

                if (dtBodyTemplate.DefaultView.Count > 0)
                {
                    dgvBodyTemplate["LineHeader", dgvBodyTemplate.Rows.Count - 1].Selected = true;
                    dgvBodyTemplate.Refresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in 'ChangeBodyTemplateDefaultViewRowFilter'.  " + ex.ToString());
            }
        }

        private void btnAddBodyRow_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(cboTemplates.Text.Replace(" ", "")))
            {               
                if ((chkBodyDuplicateSettings.Checked) && (dtBodyTemplate.DefaultView.Count > 0))
                {
                    int iRowIndex = dtBodyTemplate.DefaultView.Count - 1;

                    if (dgvBodyTemplate.SelectedRows.Count > 0)
                    {
                        iRowIndex = dgvBodyTemplate.SelectedRows[0].Index;
                    }

                    AddBodyTemplateRow(cboTemplates.Text,
                                        dtBodyTemplate.DefaultView[iRowIndex]["LineHeaderFontName"].ToString(),
                                        Convert.ToInt16(dtBodyTemplate.DefaultView[iRowIndex]["LineHeaderFontSize"]),
                                        Convert.ToInt32(dtBodyTemplate.DefaultView[iRowIndex]["LineHeaderFontColor"]),
                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[iRowIndex]["LineHeaderUnderlined"]),
                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[iRowIndex]["LineHeaderBold"]),
                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[iRowIndex]["LineHeaderItalic"]),
                                        dtBodyTemplate.DefaultView[iRowIndex]["DataSourceFontName"].ToString(),
                                        Convert.ToInt16(dtBodyTemplate.DefaultView[iRowIndex]["DataSourceFontSize"]),
                                        Convert.ToInt32(dtBodyTemplate.DefaultView[iRowIndex]["DataSourceFontColor"]),
                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[iRowIndex]["DataSourceUnderlined"]),
                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[iRowIndex]["DataSourceBold"]),
                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[iRowIndex]["DataSourceItalic"]),
                                        dtBodyTemplate.DefaultView[iRowIndex]["Spacing"].ToString(),
                                        Convert.ToInt16(dtBodyTemplate.DefaultView[iRowIndex]["BlankLines"]),
                                        Convert.ToInt16(dtBodyTemplate.DefaultView[iRowIndex]["StaticIndent"]),
                                        Convert.ToInt32(dtBodyTemplate.DefaultView[iRowIndex]["TableRowBackColor"]));
                    
                    ChangeBodyTemplateDefaultViewRowFilter(cboTemplates.Text);

                    Color cLineHeaderFontColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[iRowIndex]["LineHeaderFontColor"]));
                    Color cDataSourceFontColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[iRowIndex]["DataSourceFontColor"]));
                    Color cTableRowFontColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[iRowIndex]["TableRowBackColor"]));

                    dgvBodyTemplate["LineHeaderFontColor", iRowIndex + 1].Style.BackColor = cLineHeaderFontColor;
                    dgvBodyTemplate["LineHeaderFontColor", iRowIndex + 1].Style.ForeColor = cLineHeaderFontColor;
                    dgvBodyTemplate["LineHeaderFontColor", iRowIndex + 1].Style.SelectionBackColor = cLineHeaderFontColor;
                    dgvBodyTemplate["LineHeaderFontColor", iRowIndex + 1].Style.SelectionForeColor = cLineHeaderFontColor;

                    dgvBodyTemplate["DataSourceFontColor", iRowIndex + 1].Style.BackColor = cDataSourceFontColor;
                    dgvBodyTemplate["DataSourceFontColor", iRowIndex + 1].Style.ForeColor = cDataSourceFontColor;
                    dgvBodyTemplate["DataSourceFontColor", iRowIndex + 1].Style.SelectionBackColor = cDataSourceFontColor;
                    dgvBodyTemplate["DataSourceFontColor", iRowIndex + 1].Style.SelectionForeColor = cDataSourceFontColor;

                    dgvBodyTemplate["TableRowBackColor", iRowIndex + 1].Style.BackColor = cTableRowFontColor;
                    dgvBodyTemplate["TableRowBackColor", iRowIndex + 1].Style.ForeColor = cTableRowFontColor;
                    dgvBodyTemplate["TableRowBackColor", iRowIndex + 1].Style.SelectionBackColor = cTableRowFontColor;
                    dgvBodyTemplate["TableRowBackColor", iRowIndex + 1].Style.SelectionForeColor = cTableRowFontColor;
                }
                else
                {
                    AddBodyTemplateRow(cboTemplates.Text, "Courier New", 10, Color.Black.ToArgb(), false, false, false, "Courier New",
                                       10, Color.Black.ToArgb(), false, false, false, "0", 0, 0, Color.White.ToArgb());

                    ChangeBodyTemplateDefaultViewRowFilter(cboTemplates.Text);                   

                    dgvBodyTemplate["LineHeaderFontColor", dtBodyTemplate.DefaultView.Count - 1].Style.BackColor = Color.Black;
                    dgvBodyTemplate["LineHeaderFontColor", dtBodyTemplate.DefaultView.Count - 1].Style.ForeColor = Color.Black;
                    dgvBodyTemplate["LineHeaderFontColor", dtBodyTemplate.DefaultView.Count - 1].Style.SelectionBackColor = Color.Black;
                    dgvBodyTemplate["LineHeaderFontColor", dtBodyTemplate.DefaultView.Count - 1].Style.SelectionForeColor = Color.Black;

                    dgvBodyTemplate["DataSourceFontColor", dtBodyTemplate.DefaultView.Count - 1].Style.BackColor = Color.Black;
                    dgvBodyTemplate["DataSourceFontColor", dtBodyTemplate.DefaultView.Count - 1].Style.ForeColor = Color.Black;
                    dgvBodyTemplate["DataSourceFontColor", dtBodyTemplate.DefaultView.Count - 1].Style.SelectionBackColor = Color.Black;
                    dgvBodyTemplate["DataSourceFontColor", dtBodyTemplate.DefaultView.Count - 1].Style.SelectionForeColor = Color.Black;

                    dgvBodyTemplate["TableRowBackColor", dtBodyTemplate.DefaultView.Count - 1].Style.BackColor = Color.Black;
                    dgvBodyTemplate["TableRowBackColor", dtBodyTemplate.DefaultView.Count - 1].Style.ForeColor = Color.Black;
                    dgvBodyTemplate["TableRowBackColor", dtBodyTemplate.DefaultView.Count - 1].Style.SelectionBackColor = Color.Black;
                    dgvBodyTemplate["TableRowBackColor", dtBodyTemplate.DefaultView.Count - 1].Style.SelectionForeColor = Color.Black;
                }
            }
            else
            {
                MessageBox.Show("A Template must first be created or selected before adding a Body Row.");
            }

            bUnsavedChanges = true;
            dgvBodyTemplate.Refresh();
        }

        private void dgvSubjectTemplate_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            UpdateSubjectTextBox();
            bUnsavedChanges = true;
        }

        private void UpdateSubjectTextBox()
        {
            if (dtSubjectTemplate.DefaultView.Count > 0)
            {
                txtSubjectSample.Text = "";
                for (int i = 0; i < dtSubjectTemplate.DefaultView.Count; i++)
                {
                    txtSubjectSample.Text = txtSubjectSample.Text + dtSubjectTemplate.DefaultView[i]["DataSource"].ToString() + dtSubjectTemplate.DefaultView[i]["Separator"].ToString();
                }
            }
        }

        private void txtSubjectSample_DoubleClick(object sender, EventArgs e)
        {
            UpdateSubjectTextBox();
        }

        private void rtxtBodySample_DoubleClick(object sender, EventArgs e)
        {
            UpdateBodyRichTextBox();
            UpdateBodyRichTextBoxFont();
        }

        private List<string> GetLineHeader_UnitDefaults(string strDataSource)
        {
            List<string> lstDefaults = new List<string>();

            if ((strDataSource == "Survey_MD") || (strDataSource == "Bit_MD"))
            {                
                lstDefaults.Add("MD:");
                lstDefaults.Add("'");
            }
            else if ((strDataSource == "Survey_Inc") || (strDataSource == "Bit_Inc"))
            {
                lstDefaults.Add("INC:");
                lstDefaults.Add("°");
            }
            else if ((strDataSource == "Survey_Azm") || (strDataSource == "Bit_Azm"))
            {
                lstDefaults.Add("AZM:");
                lstDefaults.Add("°");
            }
            else if ((strDataSource == "Survey_TVD") || (strDataSource == "Bit_TVD"))
            {
                lstDefaults.Add("TVD:");
                lstDefaults.Add("'");
            }
            else if ((strDataSource == "Survey_VS") || (strDataSource == "Bit_VS"))
            {
                lstDefaults.Add("VS:");
                lstDefaults.Add("'");
            }
            else if ((strDataSource == "Survey_DLS") || (strDataSource == "Bit_DLS"))
            {
                lstDefaults.Add("DLS:");
                lstDefaults.Add("°");
            }
            else if ((strDataSource == "Survey_North") || (strDataSource == "Bit_North"))
            {
                lstDefaults.Add("N/S:");
                lstDefaults.Add("'");
            }
            else if ((strDataSource == "Survey_East") || (strDataSource == "Bit_East"))
            {
                lstDefaults.Add("E/W:");
                lstDefaults.Add("'");
            }
            else if ((strDataSource == "Survey_Temp") || (strDataSource == "CM_Temp"))
            {
                lstDefaults.Add("TEMP:");
                lstDefaults.Add("° F");
            }
            else if ((strDataSource == "Survey_A/B_Plan") || (strDataSource == "Bit_A/B_Plan"))
            {
                lstDefaults.Add("Above/Below Plan:");
                lstDefaults.Add("'");
            }
            else if ((strDataSource == "Survey_R/L_Plan") || (strDataSource == "Bit_R/L_Plan"))
            {
                lstDefaults.Add("Left/Right Plan:");
                lstDefaults.Add("'");
            }
            else if (strDataSource == "A/B_Target_Line")
            {
                lstDefaults.Add("Above/Below Target Center:");
                lstDefaults.Add("'");
            }
            else if (strDataSource == "ROP_Rotating")
            {
                lstDefaults.Add("ROP Rotating:");
                lstDefaults.Add(" ft/hr");
            }
            else if (strDataSource == "ROP_Sliding")
            {
                lstDefaults.Add("ROP Sliding:");
                lstDefaults.Add(" ft/hr");
            }
            else if (strDataSource == "Gamma")
            {
                lstDefaults.Add("Gamma:");
                lstDefaults.Add(" AAPI");
            }
            else if (strDataSource == "Activity")
            {
                lstDefaults.Add("Activity:");
                lstDefaults.Add("");
            }
            else if (strDataSource == "Client")
            {
                lstDefaults.Add("Client:");
                lstDefaults.Add("");
            }
            else if (strDataSource == "County")
            {
                lstDefaults.Add("County:");
                lstDefaults.Add("");
            }
            else if (strDataSource == "Differential_Pressure")
            {
                lstDefaults.Add("Delta Pressure:");
                lstDefaults.Add(" psi");
            }
            else if (strDataSource == "DLS_Needed")
            {
                lstDefaults.Add("DLS Needed:");
                lstDefaults.Add("°");
            }
            else if (strDataSource == "ECD")
            {
                lstDefaults.Add("ECD:");
                lstDefaults.Add("lb/gal");
            }
            else if (strDataSource == "ESD")
            {
                lstDefaults.Add("ESD:");
                lstDefaults.Add("lb/gal");
            }
            else if (strDataSource == "Flow_Rate")
            {
                lstDefaults.Add("Flow Rate:");
                lstDefaults.Add(" GPM");
            }
            else if (strDataSource == "Job_Number")
            {
                lstDefaults.Add("Job Number:");
                lstDefaults.Add("");
            }
            else if (strDataSource == "Motor_RPM")
            {
                lstDefaults.Add("Motor RPM:");
                lstDefaults.Add(" RPM");
            }
            else if (strDataSource == "Motor_Yield")
            {
                lstDefaults.Add("Motor Yield:");
                lstDefaults.Add("°");
            }
            else if (strDataSource == "Mud_Weight")
            {
                lstDefaults.Add("Mud_Weight:");
                lstDefaults.Add("lb/gal");
            }
            else if (strDataSource == "Pick_Up")
            {
                lstDefaults.Add("Pick Up:");
                lstDefaults.Add(" klbs");
            }
            else if (strDataSource == "Plan_Number")
            {
                lstDefaults.Add("Plan Number:");
                lstDefaults.Add("");
            }
            else if (strDataSource == "Rev/Gal")
            {
                lstDefaults.Add("Rev/Gal:");
                lstDefaults.Add(" RPG");
            }
            else if (strDataSource == "Rig")
            {
                lstDefaults.Add("Rig:");
                lstDefaults.Add("");
            }
            else if (strDataSource == "Rotary_Torque")
            {
                lstDefaults.Add("Rotary Torque:");
                lstDefaults.Add(" kft/lbs");
            }
            else if (strDataSource == "Rotate_Weight")
            {
                lstDefaults.Add("Rotate Weight:");
                lstDefaults.Add(" klbs");
            }
            else if (strDataSource == "Slack_Off")
            {
                lstDefaults.Add("Slack Off:");
                lstDefaults.Add(" klbs");
            }
            else if (strDataSource == "Slide_Ahead")
            {
                lstDefaults.Add("Slide Ahead:");
                lstDefaults.Add("'");
            }
            else if (strDataSource == "Slide_Seen")
            {
                lstDefaults.Add("Slide Seen:");
                lstDefaults.Add("'");
            }
            else if (strDataSource == "SPP")
            {
                lstDefaults.Add("SPP:");
                lstDefaults.Add(" psi");
            }
            else if (strDataSource == "Surface_RPM")
            {
                lstDefaults.Add("Surface RPM:");
                lstDefaults.Add(" RPM");
            }
            else if (strDataSource == "Target_Azm")
            {
                lstDefaults.Add("Target AZM:");
                lstDefaults.Add("°");
            }
            else if (strDataSource == "Target_Inc")
            {
                lstDefaults.Add("Target INC:");
                lstDefaults.Add("°");
            }
            else if (strDataSource == "Well_Name")
            {
                lstDefaults.Add("Well:");
                lstDefaults.Add("");
            }
            else if (strDataSource == "WOB")
            {
                lstDefaults.Add("WOB:");
                lstDefaults.Add(" klbs");
            }
            else
            {
                lstDefaults.Add("");
                lstDefaults.Add("");
            }

            return lstDefaults;
        }

        private void dgvBodyTemplate_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {          
            if (e.ColumnIndex == dgvBodyTemplate.Columns["DataSource"].Index)
            {
                if (string.IsNullOrEmpty(dtBodyTemplate.DefaultView[e.RowIndex]["LineHeader"].ToString().Replace(" ", "")))
                {
                    dtBodyTemplate.DefaultView[e.RowIndex]["LineHeader"] = GetLineHeader_UnitDefaults(dtBodyTemplate.DefaultView[e.RowIndex]["DataSource"].ToString())[0];
                }
                if (string.IsNullOrEmpty(dtBodyTemplate.DefaultView[e.RowIndex]["Unit"].ToString().Replace(" ", "")))
                {
                    dtBodyTemplate.DefaultView[e.RowIndex]["Unit"] = GetLineHeader_UnitDefaults(dtBodyTemplate.DefaultView[e.RowIndex]["DataSource"].ToString())[1];
                }
                dgvBodyTemplate.Refresh();
            }
            else if (dtBodyTemplate.DefaultView[e.RowIndex]["Spacing"].ToString() != "Static Indent")
            {                
                dtBodyTemplate.DefaultView[e.RowIndex]["StaticIndent"] = "0";
                dgvBodyTemplate.Refresh();
            }
            else if (dtBodyTemplate.DefaultView[e.RowIndex]["Spacing"].ToString() == "Static Indent")
            {
                bool bFontChanged = false;
                if ((dtBodyTemplate.DefaultView[e.RowIndex]["LineHeaderFontName"].ToString() != "Courier New") &&
                    (dtBodyTemplate.DefaultView[e.RowIndex]["LineHeaderFontName"].ToString() != "Consolas"))
                {
                    bFontChanged = true;
                    dtBodyTemplate.DefaultView[e.RowIndex]["LineHeaderFontName"] = "Courier New";
                    dgvBodyTemplate.Refresh();
                }
                if ((dtBodyTemplate.DefaultView[e.RowIndex]["DataSourceFontName"].ToString() != "Courier New") &&
                    (dtBodyTemplate.DefaultView[e.RowIndex]["DataSourceFontName"].ToString() != "Consolas"))
                {
                    bFontChanged = true;
                    dtBodyTemplate.DefaultView[e.RowIndex]["DataSourceFontName"] = "Courier New";
                    dgvBodyTemplate.Refresh();
                }

                if (bFontChanged)
                {
                    MessageBox.Show("'Static Indent' can only be used with a monospace font.  Use 'Courier New' or 'Consolas'.  " +
                                    "Font changed to 'Courier New'.");
                }

                if (dtBodyTemplate.DefaultView[e.RowIndex]["DataSource"].ToString() == "Comments")
                {
                    MessageBox.Show("'Static Indent' cannot be used with a 'Data Source' of 'Comments'.");
                    dtBodyTemplate.DefaultView[e.RowIndex]["Spacing"] = "0";
                }
            }
            else if ((e.ColumnIndex == dgvBodyTemplate.Columns["LineHeaderFontColor"].Index) ||
                      (e.ColumnIndex == dgvBodyTemplate.Columns["DataSourceFontColor"].Index) ||
                      (e.ColumnIndex == dgvBodyTemplate.Columns["TableNameBackColor"].Index))
            {
                int iColor = Convert.ToInt32(dtBodyTemplate.DefaultView[e.RowIndex][e.ColumnIndex]);
                dgvBodyTemplate[e.ColumnIndex, e.RowIndex].Style.BackColor = System.Drawing.Color.FromArgb(iColor);
            }

            UpdateBodyRichTextBox();
            UpdateBodyRichTextBoxFont();

            bUnsavedChanges = true;
        }

        private void UpdateBodyRichTextBox()
        {
            if (dtBodyTemplate.DefaultView.Count > 0)
            {
                int iStart = 0;
                int iLength = 0;
                rtxtBodySample.Text = "";
                for (int i = 0; i < dtBodyTemplate.DefaultView.Count; i++)
                {
                    if (rtxtBodySample.Text.Length > 0)
                    {
                        iStart = rtxtBodySample.Text.Length;
                    }
                    else
                    {
                        iStart = 0;
                    }
                    rtxtBodySample.Text = rtxtBodySample.Text + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString();
                    iLength = dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Length;

                    if (!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Replace(" ", "")))
                    {
                        iStart = rtxtBodySample.Text.Length;
                        if (dtBodyTemplate.DefaultView[i]["Spacing"].ToString() != "Static Indent")
                        {
                            iStart = iStart + Convert.ToInt16(dtBodyTemplate.DefaultView[i]["Spacing"]);
                            rtxtBodySample.Text = rtxtBodySample.Text + new string(' ', Convert.ToInt16(dtBodyTemplate.DefaultView[i]["Spacing"]));
                            rtxtBodySample.Text = rtxtBodySample.Text + dtBodyTemplate.DefaultView[i]["DataSource"].ToString() +
                                                                        dtBodyTemplate.DefaultView[i]["Unit"].ToString();
                        }
                        else
                        {
                            rtxtBodySample.Text = rtxtBodySample.Text + SpaceGenerator(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() +
                                                                                       dtBodyTemplate.DefaultView[i]["DataSource"].ToString(),// +
                                                                                       //dtBodyTemplate.DefaultView[i]["Unit"].ToString(), 
                                                                                       Convert.ToInt16(dtBodyTemplate.DefaultView[i]["StaticIndent"])) +
                                                                                       dtBodyTemplate.DefaultView[i]["DataSource"].ToString() +
                                                                                       dtBodyTemplate.DefaultView[i]["Unit"].ToString();
                        }
                        iLength = rtxtBodySample.Text.Length - iStart;                                     
                    }
                    for (int j = 0; j <= Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]); j++)
                    {
                        rtxtBodySample.Text = rtxtBodySample.Text + Environment.NewLine;
                    }
                }
                rtxtBodySample.Refresh();
            }
        }

        private void UpdateBodyRichTextBoxFont()
        {
            int iStart = 0;
            int iLength = 0;
            int iTotalIndexLength = 0;
            int iTotalStringLength = 0;
            int iLines = 0;
            string strDataSourceText;
            for (int i = 0; i < dtBodyTemplate.DefaultView.Count; i++)
            {
                iStart = 0;
                iLength = 0;

                if (!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Replace(" ", "")))
                {
                    if (rtxtBodySample.Lines[iLines].Contains(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString()))
                    {
                    //    if (dtBodyTemplate.DefaultView[i]["Spacing"].ToString() == "Static Indent")
                    //    {
                    //        //Applies the Font Setting for LineHeader to the whole line when Static Indent is selected. 
                    //        iStart = 0;//rtxtBodySample.GetFirstCharIndexFromLine(iLines);
                    //        iLength = rtxtBodySample.Lines[iLines].Length;
                    //    }
                    //    else
                    //    {
                            iStart = Convert.ToInt16(rtxtBodySample.Lines[iLines].IndexOf(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Trim(' '), 0));
                            iLength = dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Trim(' ').Length;
                        //}

                        AddFontToRichText(iStart + iTotalIndexLength, iLength,
                                        dtBodyTemplate.DefaultView[i]["LineHeaderFontName"].ToString(),
                                        Convert.ToInt16(dtBodyTemplate.DefaultView[i]["LineHeaderFontSize"]),
                                        Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderFontColor"]),
                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderUnderlined"]),
                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderBold"]),
                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderItalic"]));

                        //Needed to handle Static Indent when colors, underlined, bold or italic are different.  Font settings must be the same but other settings can be different.   
                        if (dtBodyTemplate.DefaultView[i]["Spacing"].ToString() == "Static Indent")
                        {
                        //    if ((Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderFontColor"]) !=
                        //        Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"])) ||
                        //        (Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderUnderlined"]) !=
                        //        Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceUnderlined"])) ||
                        //        (Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderBold"]) !=
                        //        Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceBold"])) ||
                        //        (Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderItalic"]) !=
                        //        Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceItalic"])))

                        //    {
                                //iStart = Convert.ToInt16(rtxtBodySample.Lines[iLines].IndexOf(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Trim(' '), 0));
                                //rtxtBodySample.SelectionStart = iStart + iTotalIndexLength;
                                //rtxtBodySample.SelectionLength = dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Trim(' ').Length;                               

                                //rtxtBodySample.SelectionColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderFontColor"]));

                            strDataSourceText = rtxtBodySample.Lines[iLines].Substring(iStart + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Length);

                            if (strDataSourceText.Contains(dtBodyTemplate.DefaultView[i]["DataSource"].ToString() +
                                                            dtBodyTemplate.DefaultView[i]["Unit"].ToString()))
                            {
                                //Static Indent Space.
                                rtxtBodySample.SelectionStart = Convert.ToInt16(rtxtBodySample.Lines[iLines].IndexOf(strDataSourceText, 0)) + 
                                                                iTotalIndexLength;

                                rtxtBodySample.SelectionLength = strDataSourceText.Length -
                                                                 dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Length -
                                                                 dtBodyTemplate.DefaultView[i]["Unit"].ToString().Length;

                                rtxtBodySample.SelectionFont = new System.Drawing.Font(dtBodyTemplate.DefaultView[i]["LineHeaderFontName"].ToString(),
                                                                                       Convert.ToInt16(dtBodyTemplate.DefaultView[i]["LineHeaderFontSize"]));

                                if (!Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceUnderlined"]))
                                {
                                    //Data Source and Unit.
                                    iStart = Convert.ToInt16(strDataSourceText.IndexOf(dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Trim(' '), 0)) +
                                                            (rtxtBodySample.Lines[iLines].Length - strDataSourceText.Length);

                                    iLength = dtBodyTemplate.DefaultView[i]["DataSource"].ToString().TrimStart(' ').Length +
                                                                        dtBodyTemplate.DefaultView[i]["Unit"].ToString().TrimEnd(' ').Length;

                                    AddFontToRichText(iStart + iTotalIndexLength, iLength,
                                                        dtBodyTemplate.DefaultView[i]["LineHeaderFontName"].ToString(),
                                                        Convert.ToInt16(dtBodyTemplate.DefaultView[i]["LineHeaderFontSize"]),
                                                        Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]),
                                                        false,
                                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceBold"]),
                                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceItalic"]));
                                }
                                else
                                {
                                    //Data Source. 
                                    iStart = Convert.ToInt16(strDataSourceText.IndexOf(dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Trim(' '), 0)) +
                                                                    (rtxtBodySample.Lines[iLines].Length - strDataSourceText.Length);

                                    iLength = dtBodyTemplate.DefaultView[i]["DataSource"].ToString().TrimStart(' ').Length;

                                    AddFontToRichText(iStart + iTotalIndexLength, iLength,
                                                        dtBodyTemplate.DefaultView[i]["LineHeaderFontName"].ToString(),
                                                        Convert.ToInt16(dtBodyTemplate.DefaultView[i]["LineHeaderFontSize"]),
                                                        Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]),
                                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceUnderlined"]),
                                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceBold"]),
                                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceItalic"]));

                                    //Unit. 
                                    iStart = Convert.ToInt16(strDataSourceText.IndexOf(dtBodyTemplate.DefaultView[i]["Unit"].ToString(), 0)) +
                                                                    (rtxtBodySample.Lines[iLines].Length - 
                                                                                          strDataSourceText.Length);

                                    iLength = dtBodyTemplate.DefaultView[i]["Unit"].ToString().Length;

                                    AddFontToRichText(iStart + iTotalIndexLength, iLength,
                                                        dtBodyTemplate.DefaultView[i]["LineHeaderFontName"].ToString(),
                                                        Convert.ToInt16(dtBodyTemplate.DefaultView[i]["LineHeaderFontSize"]),
                                                        Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]),
                                                        false,
                                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceBold"]),
                                                        Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceItalic"]));
                                }

                                

                                    //rtxtBodySample.SelectionColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]));

                                    //if (Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceUnderlined"]))
                                    //{
                                    //    rtxtBodySample.SelectionFont = new System.Drawing.Font("Courier New", 12, FontStyle.Underline | FontStyle.Bold | FontStyle.Italic);
                                    //}
                                }
                            //}
                        }
                    }
                }

                if ((dtBodyTemplate.DefaultView[i]["Spacing"].ToString() != "Static Indent") &&
                    (!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Replace(" ", ""))))
                {
                    strDataSourceText = rtxtBodySample.Lines[iLines].Substring(iStart + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Length);

                    if (strDataSourceText.Contains(dtBodyTemplate.DefaultView[i]["DataSource"].ToString() +
                                                          dtBodyTemplate.DefaultView[i]["Unit"].ToString()))
                    {
                        iStart = Convert.ToInt16(strDataSourceText.IndexOf(dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Trim(' '), 0));
                        iLength = dtBodyTemplate.DefaultView[i]["DataSource"].ToString().TrimStart(' ').Length +
                                    dtBodyTemplate.DefaultView[i]["Unit"].ToString().TrimEnd(' ').Length;


                        AddFontToRichText(iStart + iTotalIndexLength + (rtxtBodySample.Lines[iLines].Length - strDataSourceText.Length), iLength,
                                          dtBodyTemplate.DefaultView[i]["DataSourceFontName"].ToString(),
                                          Convert.ToInt16(dtBodyTemplate.DefaultView[i]["DataSourceFontSize"]),
                                          Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]),
                                          Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceUnderlined"]),
                                          Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceBold"]),
                                          Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceItalic"]));
                    }
                }

                if (i == 0)
                {
                    iTotalStringLength = -1;
                }

                iTotalStringLength = iTotalStringLength + rtxtBodySample.Lines[iLines].Length + ((Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]) + 1) * 2);
                iTotalIndexLength = iTotalIndexLength + rtxtBodySample.Lines[iLines].Length + 1 + Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]);
                
                iLines = iLines + 1 + Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]);
            }
        }

        //Adds Font to the Selected Text
        private void AddFontToRichText(int iStart, int iLength, string strFontName, int iFontSize, 
                                       int iFontColor, bool bUnderlined, bool bBold, bool bItalic)            
        {
            rtxtBodySample.SelectionStart = iStart;
            rtxtBodySample.SelectionLength = iLength;

            rtxtBodySample.SelectionColor = System.Drawing.Color.FromArgb(iFontColor);

            if (!bUnderlined && !bBold && !bItalic)
            {
                rtxtBodySample.SelectionFont = new System.Drawing.Font(strFontName, iFontSize);                
            }
            else if (bUnderlined && !bBold && !bItalic)
            {
                rtxtBodySample.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Underline);
            }
            else if (bUnderlined && bBold && !bItalic)
            {
                rtxtBodySample.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Underline | FontStyle.Bold);
            }
            else if (bUnderlined && !bBold && bItalic)
            {
                rtxtBodySample.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Underline | FontStyle.Italic);
            }
            else if (bUnderlined && bBold && bItalic)
            {
                rtxtBodySample.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Underline | FontStyle.Bold | FontStyle.Italic);
            }
            else if (!bUnderlined && bBold && !bItalic)
            {
                rtxtBodySample.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Bold);
            }
            else if (!bUnderlined && bBold && bItalic)
            {
                rtxtBodySample.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Bold | FontStyle.Italic);
            }
            else if (!bUnderlined && !bBold && bItalic)
            {
                rtxtBodySample.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Italic);
            }
        }

        //Generates needed spaces
        private string SpaceGenerator(string strText, int iStaticIndent)
        {
            string strSpaces = " ";
            for (int i = 0; i < iStaticIndent - strText.Length; i++)
            {
                strSpaces = strSpaces + " ";
            }

            return strSpaces;
        }

        private void lblStaticIndentInfo_DoubleClick(object sender, EventArgs e)
        {
            if (dgvSubjectTemplate.Columns["Name"].Visible == true)
            {
                dgvSubjectTemplate.Columns["Name"].Visible = false;
                dgvBodyTemplate.Columns["Name"].Visible = false;
            }
            else
            {
                dgvSubjectTemplate.Columns["Name"].Visible = true;
                dgvBodyTemplate.Columns["Name"].Visible = true;
            }
        }

        private void dgvSubjectTemplate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {                
                DialogResult drResult = MessageBox.Show("Delete Selected RowsTemplate?", "Delete Selected Rows", MessageBoxButtons.YesNo);

                if (drResult == DialogResult.Yes)
                {
                    for (int j = 0; j < dgvSubjectTemplate.SelectedRows.Count; j++)
                    {
                        for (int i = 0; i < dtSubjectTemplate.Rows.Count; i++)
                        {
                            if ((dtSubjectTemplate.Rows[i]["Name"].ToString() == dgvSubjectTemplate.SelectedRows[j].Cells["Name"].Value.ToString()) &&
                                (dtSubjectTemplate.Rows[i]["DataSource"].ToString() == dgvSubjectTemplate.SelectedRows[j].Cells["DataSource"].Value.ToString()))
                            {
                                dtSubjectTemplate.Rows[i].Delete();
                                dtSubjectTemplate.AcceptChanges();
                                break;
                            }
                        }
                    }
                    dgvSubjectTemplate.Refresh();
                    UpdateSubjectTextBox();
                }
                else
                {
                    e.Handled = false;
                }
            }
        }

        private void btnBodyMoveUp_Click(object sender, EventArgs e)
        {
            if (dgvBodyTemplate.SelectedRows.Count == 1)
            {
                if (dgvBodyTemplate.SelectedRows[0].Index > 0)
                {
                    int iDeleteIndex = -1;
                    int iInsertIndex = -1;
                    int iSelectedDGVIndex = dgvBodyTemplate.SelectedRows[0].Index - 1;

                    for (int i = 0; i < dtBodyTemplate.Rows.Count; i++)
                    {
                        if ((dgvBodyTemplate.SelectedRows[0].Cells["Name"].Value.ToString() == dtBodyTemplate.Rows[i]["Name"].ToString()) &&
                            (dgvBodyTemplate.SelectedRows[0].Cells["LineHeader"].Value.ToString() == dtBodyTemplate.Rows[i]["LineHeader"].ToString()) &&
                            (dgvBodyTemplate.SelectedRows[0].Cells["DataSource"].Value.ToString() == dtBodyTemplate.Rows[i]["DataSource"].ToString()) &&
                            (dgvBodyTemplate.SelectedRows[0].Cells["Unit"].Value.ToString() == dtBodyTemplate.Rows[i]["Unit"].ToString()) && (iDeleteIndex < 0))
                        {
                            iDeleteIndex = i;
                        }

                        if (iDeleteIndex >= 0)
                        {
                            for (int j = i - 1; j >= 0; j--)
                            {
                                if (dgvBodyTemplate.SelectedRows[0].Cells["Name"].Value.ToString() == dtBodyTemplate.Rows[j]["Name"].ToString())
                                {
                                    iInsertIndex = j;
                                    break;
                                }

                            }
                        }

                        if (iInsertIndex >= 0)
                        {
                            break;
                        }
                    }

                    if ((iDeleteIndex >= 0) && (iInsertIndex >= 0))
                    {
                        DataRow drRow = dtBodyTemplate.NewRow();

                        drRow["Name"] = dtBodyTemplate.Rows[iDeleteIndex]["Name"];
                        drRow["LineHeader"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeader"];
                        drRow["LineHeaderFontName"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderFontName"];
                        drRow["LineHeaderFontSize"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderFontSize"];
                        drRow["LineHeaderFontColor"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderFontColor"];
                        drRow["LineHeaderUnderlined"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderUnderlined"];
                        drRow["LineHeaderBold"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderBold"];
                        drRow["LineHeaderItalic"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderItalic"];
                        drRow["Spacing"] = dtBodyTemplate.Rows[iDeleteIndex]["Spacing"];
                        drRow["DataSource"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSource"];
                        drRow["Unit"] = dtBodyTemplate.Rows[iDeleteIndex]["Unit"];
                        drRow["DataSourceFontName"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSourceFontName"];
                        drRow["DataSourceFontSize"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSourceFontSize"];
                        drRow["DataSourceFontColor"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSourceFontColor"];
                        drRow["DataSourceUnderlined"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSourceUnderlined"];
                        drRow["DataSourceBold"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSourceBold"];
                        drRow["DataSourceItalic"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSourceItalic"];
                        drRow["BlankLines"] = dtBodyTemplate.Rows[iDeleteIndex]["BlankLines"];
                        drRow["StaticIndent"] = dtBodyTemplate.Rows[iDeleteIndex]["StaticIndent"];
                        drRow["TableRowBackColor"] = dtBodyTemplate.Rows[iDeleteIndex]["TableRowBackColor"];

                        Color cLineHeaderFontColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderFontColor"]));
                        Color cDataSourceFontColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.Rows[iDeleteIndex]["DataSourceFontColor"]));
                        Color cTableRowFontColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.Rows[iDeleteIndex]["TableRowBackColor"]));


                        dtBodyTemplate.Rows.RemoveAt(iDeleteIndex);

                        dtBodyTemplate.Rows.InsertAt(drRow, iInsertIndex);

                        dgvBodyTemplate.Refresh();

                        dgvBodyTemplate.Rows[iSelectedDGVIndex].Selected = true;

                        RefreshDGVColors(cLineHeaderFontColor, cDataSourceFontColor, cTableRowFontColor);

                        bUnsavedChanges = true;
                    }
                }
            }
        }

        private void btnBodyMoveDown_Click(object sender, EventArgs e)
        {
            if (dgvBodyTemplate.SelectedRows.Count == 1)
            {
                if (dgvBodyTemplate.SelectedRows[0].Index < (dgvBodyTemplate.Rows.Count - 2))
                {
                    int iDeleteIndex = -1;
                    int iInsertIndex = -1;
                    int iSelectedDGVIndex = dgvBodyTemplate.SelectedRows[0].Index + 1;

                    for (int i = 0; i < dtBodyTemplate.Rows.Count; i++)
                    {
                        if ((dgvBodyTemplate.SelectedRows[0].Cells["Name"].Value.ToString() == dtBodyTemplate.Rows[i]["Name"].ToString()) &&
                            (dgvBodyTemplate.SelectedRows[0].Cells["LineHeader"].Value.ToString() == dtBodyTemplate.Rows[i]["LineHeader"].ToString()) &&
                            (dgvBodyTemplate.SelectedRows[0].Cells["DataSource"].Value.ToString() == dtBodyTemplate.Rows[i]["DataSource"].ToString()) &&
                            (dgvBodyTemplate.SelectedRows[0].Cells["Unit"].Value.ToString() == dtBodyTemplate.Rows[i]["Unit"].ToString()) && (iDeleteIndex < 0))
                        {
                            iDeleteIndex = i;
                        }

                        if (iDeleteIndex >= 0)
                        {
                            if (dgvBodyTemplate.SelectedRows[0].Cells["Name"].Value.ToString() == dtBodyTemplate.Rows[i]["Name"].ToString())
                            {
                                iInsertIndex = i + 1;
                                break;
                            }
                        }
                    }


                    if ((iDeleteIndex >= 0) && (iInsertIndex >= 0))
                    {
                        DataRow drRow = dtBodyTemplate.NewRow();

                        drRow["Name"] = dtBodyTemplate.Rows[iDeleteIndex]["Name"];
                        drRow["LineHeader"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeader"];
                        drRow["LineHeaderFontName"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderFontName"];
                        drRow["LineHeaderFontSize"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderFontSize"];
                        drRow["LineHeaderFontColor"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderFontColor"];
                        drRow["LineHeaderUnderlined"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderUnderlined"];
                        drRow["LineHeaderBold"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderBold"];
                        drRow["LineHeaderItalic"] = dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderItalic"];
                        drRow["Spacing"] = dtBodyTemplate.Rows[iDeleteIndex]["Spacing"];
                        drRow["DataSource"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSource"];
                        drRow["Unit"] = dtBodyTemplate.Rows[iDeleteIndex]["Unit"];
                        drRow["DataSourceFontName"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSourceFontName"];
                        drRow["DataSourceFontSize"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSourceFontSize"];
                        drRow["DataSourceFontColor"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSourceFontColor"];
                        drRow["DataSourceUnderlined"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSourceUnderlined"];
                        drRow["DataSourceBold"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSourceBold"];
                        drRow["DataSourceItalic"] = dtBodyTemplate.Rows[iDeleteIndex]["DataSourceItalic"];
                        drRow["BlankLines"] = dtBodyTemplate.Rows[iDeleteIndex]["BlankLines"];
                        drRow["StaticIndent"] = dtBodyTemplate.Rows[iDeleteIndex]["StaticIndent"];
                        drRow["TableRowBackColor"] = dtBodyTemplate.Rows[iDeleteIndex]["TableRowBackColor"];

                        Color cLineHeaderFontColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.Rows[iDeleteIndex]["LineHeaderFontColor"]));
                        Color cDataSourceFontColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.Rows[iDeleteIndex]["DataSourceFontColor"]));
                        Color cTableRowFontColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.Rows[iDeleteIndex]["TableRowBackColor"]));
                                          

                        dtBodyTemplate.Rows.RemoveAt(iDeleteIndex);

                        dtBodyTemplate.Rows.InsertAt(drRow, iInsertIndex);

                        dgvBodyTemplate.Refresh();

                        dgvBodyTemplate.Rows[iSelectedDGVIndex].Selected = true;

                        RefreshDGVColors(cLineHeaderFontColor, cDataSourceFontColor, cTableRowFontColor);

                        bUnsavedChanges = true;
                    }
                }
            }
        }

        private void RefreshDGVColors(Color cLineHeaderFontColor, Color cDataSourceFontColor, Color cTableRowFontColor)
        {            

            dgvBodyTemplate.SelectedRows[0].Cells["LineHeaderFontColor"].Style.BackColor = cLineHeaderFontColor;
            dgvBodyTemplate.SelectedRows[0].Cells["LineHeaderFontColor"].Style.ForeColor = cLineHeaderFontColor;
            dgvBodyTemplate.SelectedRows[0].Cells["LineHeaderFontColor"].Style.SelectionBackColor = cLineHeaderFontColor;
            dgvBodyTemplate.SelectedRows[0].Cells["LineHeaderFontColor"].Style.SelectionForeColor = cLineHeaderFontColor;

            dgvBodyTemplate.SelectedRows[0].Cells["DataSourceFontColor"].Style.BackColor = cDataSourceFontColor;
            dgvBodyTemplate.SelectedRows[0].Cells["DataSourceFontColor"].Style.ForeColor = cDataSourceFontColor;
            dgvBodyTemplate.SelectedRows[0].Cells["DataSourceFontColor"].Style.SelectionBackColor = cDataSourceFontColor;
            dgvBodyTemplate.SelectedRows[0].Cells["DataSourceFontColor"].Style.SelectionForeColor = cDataSourceFontColor;

            dgvBodyTemplate.SelectedRows[0].Cells["TableRowBackColor"].Style.BackColor = cTableRowFontColor;
            dgvBodyTemplate.SelectedRows[0].Cells["TableRowBackColor"].Style.ForeColor = cTableRowFontColor;
            dgvBodyTemplate.SelectedRows[0].Cells["TableRowBackColor"].Style.SelectionBackColor = cTableRowFontColor;
            dgvBodyTemplate.SelectedRows[0].Cells["TableRowBackColor"].Style.SelectionForeColor = cTableRowFontColor;
        }

        private void btnSubjectMoveUp_Click(object sender, EventArgs e)
        {
            if (dgvSubjectTemplate.SelectedRows.Count == 1)
            {
                if (dgvSubjectTemplate.SelectedRows[0].Index > 0)
                {
                    int iDeleteIndex = -1;
                    int iInsertIndex = -1;
                    int iSelectedDGVIndex = dgvSubjectTemplate.SelectedRows[0].Index - 1;

                    for (int i = 0; i < dtSubjectTemplate.Rows.Count; i++)
                    {
                        if ((dgvSubjectTemplate.SelectedRows[0].Cells["Name"].Value.ToString() == dtSubjectTemplate.Rows[i]["Name"].ToString()) &&
                            (dgvSubjectTemplate.SelectedRows[0].Cells["DataSource"].Value.ToString() == dtSubjectTemplate.Rows[i]["DataSource"].ToString()) &&
                            (dgvSubjectTemplate.SelectedRows[0].Cells["Separator"].Value.ToString() == dtSubjectTemplate.Rows[i]["Separator"].ToString()) && (iDeleteIndex < 0))
                        {
                            iDeleteIndex = i;
                        }

                        if (iDeleteIndex >= 0)
                        {
                            for (int j = i - 1; j >= 0; j--)
                            {
                                if (dgvSubjectTemplate.SelectedRows[0].Cells["Name"].Value.ToString() == dtSubjectTemplate.Rows[j]["Name"].ToString())
                                {
                                    iInsertIndex = j;
                                    break;
                                }
                                
                            }
                        }

                        if (iInsertIndex >= 0)
                        {
                            break;
                        }
                    }

                    if ((iDeleteIndex >= 0) && (iInsertIndex >= 0))
                    {

                        DataRow drRow = dtSubjectTemplate.NewRow();

                        drRow["Name"] = dtSubjectTemplate.Rows[iDeleteIndex]["Name"];
                        drRow["DataSource"] = dtSubjectTemplate.Rows[iDeleteIndex]["DataSource"];
                        drRow["Separator"] = dtSubjectTemplate.Rows[iDeleteIndex]["Separator"];

                        dtSubjectTemplate.Rows.RemoveAt(iDeleteIndex);

                        dtSubjectTemplate.Rows.InsertAt(drRow, iInsertIndex);

                        dgvSubjectTemplate.Refresh();

                        dgvSubjectTemplate.Rows[iSelectedDGVIndex].Selected = true;

                        bUnsavedChanges = true;
                    }
                }
            }
        }

        private void btnSubjectMoveDown_Click(object sender, EventArgs e)
        {
            if (dgvSubjectTemplate.SelectedRows.Count == 1)
            {
                if (dgvSubjectTemplate.SelectedRows[0].Index < (dgvSubjectTemplate.Rows.Count - 1))
                {
                    int iDeleteIndex = -1;
                    int iInsertIndex = -1;
                    int iSelectedDGVIndex = dgvSubjectTemplate.SelectedRows[0].Index + 1;

                    for (int i = 0; i < dtSubjectTemplate.Rows.Count; i++)
                    {
                        if ((dgvSubjectTemplate.SelectedRows[0].Cells["Name"].Value.ToString() == dtSubjectTemplate.Rows[i]["Name"].ToString()) &&
                            (dgvSubjectTemplate.SelectedRows[0].Cells["DataSource"].Value.ToString() == dtSubjectTemplate.Rows[i]["DataSource"].ToString()) &&
                            (dgvSubjectTemplate.SelectedRows[0].Cells["Separator"].Value.ToString() == dtSubjectTemplate.Rows[i]["Separator"].ToString()) && (iDeleteIndex < 0))
                        {
                            iDeleteIndex = i;
                        }

                        if (iDeleteIndex >= 0)
                        {
                            if (dgvSubjectTemplate.SelectedRows[0].Cells["Name"].Value.ToString() == dtSubjectTemplate.Rows[i]["Name"].ToString())
                            {
                                iInsertIndex = i + 1;
                                break;
                            }
                        }
                    }

                    if ((iDeleteIndex >= 0) && (iInsertIndex >= 0))
                    {

                        DataRow drRow = dtSubjectTemplate.NewRow();

                        drRow["Name"] = dtSubjectTemplate.Rows[iDeleteIndex]["Name"];
                        drRow["DataSource"] = dtSubjectTemplate.Rows[iDeleteIndex]["DataSource"];
                        drRow["Separator"] = dtSubjectTemplate.Rows[iDeleteIndex]["Separator"];

                        dtSubjectTemplate.Rows.RemoveAt(iDeleteIndex);

                        dtSubjectTemplate.Rows.InsertAt(drRow, iInsertIndex);

                        dgvSubjectTemplate.Refresh();

                        dgvSubjectTemplate.Rows[iSelectedDGVIndex].Selected = true;

                        bUnsavedChanges = true;
                    }
                }
            }
        }

        private void btnDeleteTemplate_Click(object sender, EventArgs e)
        {
            DialogResult drResult = MessageBox.Show("OK to delete " + cboTemplates.Text + 
                                                    "?  This will be permanent and saved to the DB.", "Delete Template", 
                                                    MessageBoxButtons.YesNo);
            
            if (drResult == DialogResult.Yes)
            {
                for (int i = 0; i < dtSubjectTemplate.Rows.Count; i++)
                {
                    if (dtSubjectTemplate.Rows[i]["Name"].ToString() == cboTemplates.Text)
                    {
                        dtSubjectTemplate.Rows[i].Delete();
                    }
                }

                dtSubjectTemplate.AcceptChanges();

                for (int i = 0; i < dtBodyTemplate.Rows.Count; i++)
                {
                    if (dtBodyTemplate.Rows[i]["Name"].ToString() == cboTemplates.Text)
                    {
                        dtBodyTemplate.Rows[i].Delete();
                    }
                }
                
                dtBodyTemplate.AcceptChanges();

                cboTemplates.Items.Clear();

                for (int i = 0; i < dtSubjectTemplate.Rows.Count; i++)
                {
                    if (!cboTemplates.Items.Contains(dtSubjectTemplate.Rows[i]["Name"].ToString()))
                    {
                        cboTemplates.Items.Add(dtSubjectTemplate.Rows[i]["Name"].ToString());
                    }
                }               

                dsTemplates.WriteXml(GetTemplateFilePath);
                if (cboTemplates.Items.Count > 0)
                {
                    cboTemplates.Text = cboTemplates.Items[0].ToString();
                    ChangeBodyTemplateDefaultViewRowFilter(cboTemplates.Text);
                }

                AreTemplatesChanged = true;
                bUnsavedChanges = false;
            }


        }
    }
}
