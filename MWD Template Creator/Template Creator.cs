using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Text;
using System.IO;
using System.Collections;
using OutlookTools = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using Microsoft.VisualBasic;

namespace MWD_Template_Creator
{
    public partial class frmTemplateCreator : Form
    {
        StringBuilder builder = new StringBuilder();       

        //Integer that tracks the number of SF values equal to or below 1.50. 
        int intSFDangerCount = 0;
        string strSubjectLine;
        decimal dSurveyMD;
        decimal dBitMD;
        decimal dInc;
        decimal dAzm;        
        decimal dTVD;
        decimal dNorth;
        decimal dEast;
        decimal dVS;
        int iSVNumOfProj;
        DateTime dtBottomLine;
        int iSurveyACCounter;
        int iSignatureCounter;
        string strSurveyPath;
        string strACPath;
        bool bMailFilesUpdated;
        bool bAllFilesUpdated;

        private DataSet dsTemplates;
        private DataTable dtSubjectTemplate;
        private DataTable dtBodyTemplate;
        private DataTable dtVariables;

        List<string> lstACHeaders = new List<string>();
        public frmTemplateCreator()
        {
            InitializeComponent();

            //this.MaximizeBox = false;

            dsTemplates = new DataSet("Templates");
            dtSubjectTemplate = new DataTable("SubjectTemplate");
            dtBodyTemplate = new DataTable("BodyTemplate");
            dsTemplates.Tables.Add(CreateBodyTemplateTable());
            dsTemplates.Tables.Add(CreateSubjectTemplateTable());

            txtTemplateFilePath.Text = Properties.Settings.Default.TemplateFile;

            if (File.Exists(txtTemplateFilePath.Text))
            {
                try
                {
                    dsTemplates.ReadXml(txtTemplateFilePath.Text);
                    TemplateFileChanged();                    
                }
                catch
                {
                    FileIsNotValid(txtTemplateFilePath.Text);
                }
            }
            else
            {
                FileDoesNotExist();
            }

            lblProgress.Text = "";

            txtJobNumber.Text = Properties.Settings.Default.JobNumber;
            txtClient.Text = Properties.Settings.Default.Client;
            txtWellName.Text = Properties.Settings.Default.WellName;
            txtRig.Text = Properties.Settings.Default.Rig;
            txtSurfaceRPM.Text = Properties.Settings.Default.RPM;
            txtFlowRate.Text = Properties.Settings.Default.FlowRate;
            txtGamma.Text = Properties.Settings.Default.Gamma;
            txtWOB.Text = Properties.Settings.Default.WOB;
            txtDifferentialPressure.Text = Properties.Settings.Default.Diff;
            txtSPP.Text = Properties.Settings.Default.SPP;
            txtACFileName.Text = Properties.Settings.Default.ACFileName;
            txtACThreshold.Text = Properties.Settings.Default.ACThreshold;
            txtSurveyFileName.Text = Properties.Settings.Default.SurveyFileName;
            txtTargetInc.Text = Properties.Settings.Default.TargetInc;
            txtTargetAzm.Text = Properties.Settings.Default.TargetAzm;
            txtROPRotating.Text = Properties.Settings.Default.ROPRotating;
            chkManualOrientation.Checked = Properties.Settings.Default.ManualOrientation;
            chkWithoutAC.Checked = Properties.Settings.Default.WithoutAC;
            txtSVNumOfProj.Text = Properties.Settings.Default.SVProjections;
            txtPlanNorthZeroReference.Text = Properties.Settings.Default.PlanNorthZeroReference;
            txtPlanEastZeroReference.Text = Properties.Settings.Default.PlanEastZeroReference;
            txtSurveyTemp.Text = Properties.Settings.Default.ToolTemp;
            chkLateralTargetLine.Checked = Properties.Settings.Default.LateralTargetLine;
            txtTargetLineBeginningVS.Text = Properties.Settings.Default.LateralTargetLineVS;
            txtTargetLineBeginningTVD.Text = Properties.Settings.Default.LateralTargetLineTVD;
            txtTargetLineInc.Text = Properties.Settings.Default.LateralTargetLineInc;
            rtxtRecipientList.Text = Properties.Settings.Default.RecipientList;
            rtxtMailFiles.Text = Properties.Settings.Default.MailFiles;
            rtxtDaySignature.Text = Properties.Settings.Default.DaySignature;
            rtxtNightSignature.Text = Properties.Settings.Default.NightSignature;
            txtMailFolder.Text = Properties.Settings.Default.MailFolder;
            txtCounty.Text = Properties.Settings.Default.County;
            txtActivity.Text = Properties.Settings.Default.Activity;
            txtPlanNumber.Text = Properties.Settings.Default.PlanNumber;
            txtROPSliding.Text = Properties.Settings.Default.ROPSliding;
            chkBottomLineChecker.Checked = Properties.Settings.Default.UseBottomLine;
            txtMWDSurveyFileName.Text = Properties.Settings.Default.MWDSurveyFile;            
            txtDLSNeeded.Text = Properties.Settings.Default.DLSNeeded;
            txtRotaryTorque.Text = Properties.Settings.Default.RotaryTorque;
            txtRPG.Text = Properties.Settings.Default.RPG;
            cboTemplates.Text = Properties.Settings.Default.CurrentTemplate;
            txtSurveyCourseLength.Text = Properties.Settings.Default.SurveyCourseLength;
            txtSlideAhead.Text = Properties.Settings.Default.SlideAhead;
            txtSlideSeen.Text = Properties.Settings.Default.SlideSeen;
            txtRotateWeight.Text = Properties.Settings.Default.RotateWeight;
            txtSlackOff.Text = Properties.Settings.Default.SlackOff;
            txtPickUp.Text = Properties.Settings.Default.PickUp;
            txtCMTemp.Text = Properties.Settings.Default.CMTemp;
            txtBitAboveBelow.Text = Properties.Settings.Default.BitAboveBelow;
            txtBitRightLeft.Text = Properties.Settings.Default.BitRightLeft;
            txtSurveyAboveBelow.Text = Properties.Settings.Default.SurveyAboveBelow;
            txtSurveyRightLeft.Text = Properties.Settings.Default.SurveyRightLeft;
            rtxtComments.Text = Properties.Settings.Default.Comments;
            chkComments.Checked = Properties.Settings.Default.CommentsIncluded;
            chkThirdParty.Checked = Properties.Settings.Default.ThirdParty;
            chkChangePlanNSEWZeroReference.Checked = Properties.Settings.Default.UsePlanZeroReference;            
            txtTFMode.Text = Properties.Settings.Default.TFMode;
            txtSlideTF.Text = Properties.Settings.Default.SlideTF;
            rtxtCCRecipientList.Text = Properties.Settings.Default.CCRecipientList;
            txtMudWeight.Text = Properties.Settings.Default.MudWeight;
            txtECD.Text = Properties.Settings.Default.ECD;
            txtESD.Text = Properties.Settings.Default.ESD;

            if (Properties.Settings.Default.SigToUse == "Day")
            {
                rbDay.Checked = true;
            }
            else if (Properties.Settings.Default.SigToUse == "Night")
            {
                rbNight.Checked = true;
            }
            else
            {
                rbAutomatic.Checked = true;
            }

            ScreenSetup();

            CreateVariableDataTable();

            if (string.IsNullOrEmpty(txtACThreshold.Text.Replace(" ", "")))
            {
                txtACThreshold.Text = "1.50";
            }            

            InstalledFontCollection ifcFonts = new InstalledFontCollection();

            iSVNumOfProj = Convert.ToInt16(txtSVNumOfProj.Text);
                        
            iSurveyACCounter = 1;
            iSignatureCounter = 1;

            Timer tmrSurveyAC = new Timer();
            tmrSurveyAC.Tick += new EventHandler(FilesUpdatedTimerEventProcessor);
            tmrSurveyAC.Interval = 5000;
            tmrSurveyAC.Start();            
            
            if (txtSurveyFileName.Text == "")
            {
                txtSurveyFileName.Text = @"c:\";
            }
            if (txtACFileName.Text == "")
            {
                txtACFileName.Text = @"c:\";
            }
            if (txtMailFolder.Text == "")
            {
                txtMailFolder.Text = @"c:\";
            }
            if (txtMWDSurveyFileName.Text == "")
            {                
                txtMWDSurveyFileName.Text = @"c:\";
            }

            ShowGrpBottomLine();
            //lblBottomLineChecked.Visible = chkBottomLineChecker.Checked;

            grpLateralTargetLine.Visible = chkLateralTargetLine.Checked;

            strSurveyPath = txtSurveyFileName.Text;
            //strACPath = txtACFileName.Text;

            //if (!chkManualOrientation.Checked)
            //{
            //    if (!chkThirdParty.Checked)
            //    {
            //        txtAboveBelow.Enabled = false;
            //    }
            //}

            if (string.IsNullOrEmpty(dtBodyTemplate.Rows[0]["TableRowBackColor"].ToString().Replace(" ", "")))
            {
                for (int i = 0; i < dtBodyTemplate.Rows.Count; i++)
                {
                    dtBodyTemplate.Rows[i]["TableRowBackColor"] = Color.White.ToArgb();
                }
                dsTemplates.WriteXml(txtTemplateFilePath.Text);
            }
            
        }

        private void FileIsNotValid(string strPath)
        {
            DialogResult drResult = MessageBox.Show("The selected file was not a valid Template File.  " +
                                                            "To overwrite the selected file click 'Yes'.  To select a different file click 'No'.",
                                                            "Invalid File", MessageBoxButtons.YesNo);

            if (drResult == DialogResult.Yes)
            {
                dsTemplates.Clear();
                dsTemplates.WriteXml(strPath);
                cboTemplates.Text = "";
                cboTemplates.Items.Clear();
                txtTemplateFilePath.Text = strPath;
                TemplateFileChanged();
            }
            else
            {
                FileDoesNotExist();
            }
        }

        private void FileDoesNotExist()
        {
            DialogResult drResult = MessageBox.Show("To create a new 'Template File' select 'Yes' and select the folder location.  " + Environment.NewLine +
                                                    "To select a valid template file select 'No'.", "Create New or Find File", MessageBoxButtons.YesNo);

            if (drResult == DialogResult.Yes)
            {
                CreateNewFile();
            }
            else if (drResult == DialogResult.No)
            {
                dsTemplates.Clear();
                string strFilePath = GetFilePath(@"c:\", "Text Files|*.txt");

                if (File.Exists(strFilePath))
                {
                    try
                    {
                        dsTemplates.ReadXml(strFilePath);
                        TemplateFileChanged();
                        txtTemplateFilePath.Text = strFilePath;
                    }
                    catch
                    {
                        FileIsNotValid(strFilePath);
                    }
                }
                else
                {
                    FileDoesNotExist();
                }
            }
        }

        private void CreateNewFile()
        {
            frmTemplateName frmName = new frmTemplateName();

            frmName.LabelText = "Template Storage File Name";
            frmName.FormText = "Storage File Name";

            frmName.ShowDialog();

            if (frmName.OKToAdd)
            {
                string strFolderPath = GetFolderPath(@"c:\", true);

                string strFilePath = strFolderPath + @"\" + frmName.GetTemplateName + ".txt";
            
                if (!File.Exists(strFilePath))
                {
                    SaveNewFile(strFilePath);
                }
                else
                {
                    DialogResult drResult = MessageBox.Show("This file already exists.  To overwrite select 'Yes'. To cancel click 'No'.", "Overwrite File", MessageBoxButtons.YesNo);

                    if (drResult == DialogResult.Yes)
                    {
                        SaveNewFile(strFilePath);
                    }
                }
            }

            frmName.Dispose();
        }

        private void btnNewTemplateFile_Click(object sender, EventArgs e)
        {
            CreateNewFile();
        }

        private void SaveNewFile(string strFilePath)
        {
            try
            {
                DialogResult drResult = MessageBox.Show("Would you like to create a new empty file?  If so then click 'Yes'." + Environment.NewLine +
                                                        "To create a new file with the current template settings click 'No'.", "Create New File", MessageBoxButtons.YesNoCancel);

                if (drResult == DialogResult.Yes)
                {
                    dsTemplates.Clear();
                    dsTemplates.WriteXml(strFilePath);
                    TemplateFileChanged();
                    txtTemplateFilePath.Text = strFilePath;
                }
                else if (drResult == DialogResult.No)
                {
                    dsTemplates.WriteXml(strFilePath);
                    TemplateFileChanged();
                    txtTemplateFilePath.Text = strFilePath;
                }
            }
            catch
            {
                MessageBox.Show("There was a problem in the 'Save New File' method.");
            }
        }

        private void TemplateFileChanged()
        {
            string strFirstTemplate = "";
            cboTemplates.Items.Clear();
            cboTemplates.Text = "";
            for (int i = 0; i < dtSubjectTemplate.Rows.Count; i++)
            {
                if (!cboTemplates.Items.Contains(dtSubjectTemplate.Rows[i]["Name"].ToString()))
                {
                    cboTemplates.Items.Add(dtSubjectTemplate.Rows[i]["Name"].ToString());
                }
                if (i == 0)
                {
                    strFirstTemplate = dtSubjectTemplate.Rows[i]["Name"].ToString();
                }

                if (dtSubjectTemplate.Rows[i]["Name"].ToString() == Properties.Settings.Default.CurrentTemplate)
                {
                    strFirstTemplate = Properties.Settings.Default.CurrentTemplate;
                }
            }

            dtSubjectTemplate.DefaultView.RowFilter = "";
            dtBodyTemplate.DefaultView.RowFilter = "";
            dtSubjectTemplate.DefaultView.RowFilter = "Name = '" + strFirstTemplate + "'";
            dtBodyTemplate.DefaultView.RowFilter = "Name = '" + strFirstTemplate + "'";

            cboTemplates.Text = strFirstTemplate;

            Properties.Settings.Default.TemplateFile = txtTemplateFilePath.Text;
            Properties.Settings.Default.Save();

            ScreenSetup();
        }

        private void CreateVariableDataTable()
        {
            dtVariables = new DataTable("Variables");

            dtVariables.Columns.Add("Name", typeof(string));
            dtVariables.Columns.Add("Value", typeof(string));
            dtVariables.Columns.Add("Unit", typeof(string));
        }

        private DataTable CreateSubjectTemplateTable()
        {
            dtSubjectTemplate.Columns.Add("Name", typeof(string));
            dtSubjectTemplate.Columns.Add("DataSource", typeof(string));
            dtSubjectTemplate.Columns.Add("Separator", typeof(string));

            return dtSubjectTemplate;
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

        private void FilesUpdatedTimerEventProcessor(object sender, EventArgs e)
        {
            iSurveyACCounter += 1;
            iSignatureCounter += 1;
            if (iSurveyACCounter >= 5)
            {
                DateTime dtLastAccess;

                TimeSpan tsDiff;

                double dDiffInMinutes;
                int iWholeMinutes;

                DateTime dtLastUpdate;
                TimeSpan tsDiffUpdate;
                double dDiffInMinutesUpdate;
                int iWholeMinutesUpdate;

                //DD Survey File Check.
                //if (File.Exists(txtSurveyFileName.Text))
                if (!chkThirdParty.Checked)
                {
                    //dtLastAccess = File.GetLastAccessTime(txtSurveyFileName.Text);
                    ////dtLastAccess = File.GetLastWriteTime(txtSurveyFileName.Text);

                    //tsDiff = DateTime.Now - dtLastAccess;

                    //dDiffInMinutes = tsDiff.TotalMinutes;
                    //iWholeMinutes = (int)dDiffInMinutes;

                    dtLastUpdate = File.GetLastWriteTime(txtSurveyFileName.Text);

                    tsDiffUpdate = DateTime.Now - dtLastUpdate;

                    dDiffInMinutesUpdate = tsDiffUpdate.TotalMinutes;
                    iWholeMinutesUpdate = (int)dDiffInMinutesUpdate;

                    //if ((iWholeMinutes < 5) || (iWholeMinutesUpdate < 5))
                    if (iWholeMinutesUpdate < 5)
                    {
                        lblSurveyFileUpdated.BackColor = Color.MediumSeaGreen;
                        lblSurveyFileUpdated.ForeColor = Color.Black;
                        bAllFilesUpdated = true;
                    }
                    else
                    {
                        btnBuild.BackColor = Color.Tomato;
                        btnBuild.ForeColor = Color.Gainsboro;
                        lblBottomLineChecked.BackColor = Color.Tomato;
                        lblBottomLineChecked.ForeColor = Color.Gainsboro;
                        lblSurveyFileUpdated.BackColor = Color.Tomato;
                        lblSurveyFileUpdated.ForeColor = Color.Gainsboro;
                        bAllFilesUpdated = false;
                    }
                }

                //MWD Survey File Check.
                if ((chkBottomLineChecker.Checked) && (File.Exists(txtMWDSurveyFileName.Text)))
                {
                    //dtLastAccess = File.GetLastAccessTime(txtMWDSurveyFileName.Text);


                    //tsDiff = DateTime.Now - dtLastAccess;

                    //dDiffInMinutes = tsDiff.TotalMinutes;
                    //iWholeMinutes = (int)dDiffInMinutes;

                    dtLastUpdate = File.GetLastWriteTime(txtMWDSurveyFileName.Text);

                    tsDiffUpdate = DateTime.Now - dtLastUpdate;

                    dDiffInMinutesUpdate = tsDiffUpdate.TotalMinutes;
                    iWholeMinutesUpdate = (int)dDiffInMinutesUpdate;

                    //if ((iWholeMinutes < 10) || (iWholeMinutesUpdate < 10))
                    if (iWholeMinutesUpdate < 10)
                    {
                        lblMWDSurveyFileUpdated.BackColor = Color.MediumSeaGreen;
                        lblMWDSurveyFileUpdated.ForeColor = Color.Black;
                        bAllFilesUpdated = true;
                    }
                    else
                    {
                        btnBuild.BackColor = Color.Tomato;
                        btnBuild.ForeColor = Color.Gainsboro;
                        lblBottomLineChecked.BackColor = Color.Tomato;
                        lblBottomLineChecked.ForeColor = Color.Gainsboro;
                        lblMWDSurveyFileUpdated.BackColor = Color.Tomato;
                        lblMWDSurveyFileUpdated.ForeColor = Color.Gainsboro;
                        bAllFilesUpdated = false;
                    }
                }

                if ((!chkWithoutAC.Checked) && (File.Exists(txtACFileName.Text)))
                {
                    //dtLastAccess = File.GetLastAccessTime(txtACFileName.Text);

                    //tsDiff = DateTime.Now - dtLastAccess;

                    //dDiffInMinutes = tsDiff.TotalMinutes;
                    //iWholeMinutes = (int)dDiffInMinutes;

                    dtLastUpdate = File.GetLastWriteTime(txtACFileName.Text);

                    tsDiffUpdate = DateTime.Now - dtLastUpdate;

                    dDiffInMinutesUpdate = tsDiffUpdate.TotalMinutes;
                    iWholeMinutesUpdate = (int)dDiffInMinutesUpdate;

                    if (iWholeMinutesUpdate < 10)
                    {
                        lblACFileUpdated.BackColor = Color.MediumSeaGreen;
                        lblACFileUpdated.ForeColor = Color.Black;
                        if (bAllFilesUpdated)
                        {
                            bAllFilesUpdated = true;
                        }
                    }
                    else
                    {
                        btnBuild.BackColor = Color.Tomato;
                        btnBuild.ForeColor = Color.Gainsboro;
                        lblBottomLineChecked.BackColor = Color.Tomato;
                        lblBottomLineChecked.ForeColor = Color.Gainsboro;
                        lblACFileUpdated.BackColor = Color.Tomato;
                        lblACFileUpdated.ForeColor = Color.Gainsboro;
                        bAllFilesUpdated = false;
                    }
                }

                for (int j = 0; j < rtxtMailFiles.Lines.Length; j++)
                {
                    if (File.Exists(txtMailFolder.Text + @"\" + rtxtMailFiles.Lines[j]))
                    {
                        //dtLastAccess = File.GetLastAccessTime(txtMailFolder.Text + @"\" + rtxtMailFiles.Lines[j]);

                        //tsDiff = DateTime.Now - dtLastAccess;

                        //dDiffInMinutes = tsDiff.TotalMinutes;
                        //iWholeMinutes = (int)dDiffInMinutes;

                        dtLastUpdate = File.GetLastWriteTime(txtMailFolder.Text + @"\" + rtxtMailFiles.Lines[j]);

                        tsDiffUpdate = DateTime.Now - dtLastUpdate;

                        dDiffInMinutesUpdate = tsDiffUpdate.TotalMinutes;
                        iWholeMinutesUpdate = (int)dDiffInMinutesUpdate;

                        rtxtMailFiles.SelectionStart = rtxtMailFiles.Find(rtxtMailFiles.Lines[j]);
                        rtxtMailFiles.SelectionLength = rtxtMailFiles.Lines[j].Length;

                        if (iWholeMinutesUpdate < 10)
                        {
                            rtxtMailFiles.SelectionColor = Color.MediumSeaGreen;
                            if (j == 0)
                            {
                                bMailFilesUpdated = true;
                            }
                            else if (bMailFilesUpdated)
                            {
                                bMailFilesUpdated = true;
                            }
                        }
                        else
                        {
                            rtxtMailFiles.SelectionColor = Color.Black;
                            bMailFilesUpdated = false;
                        }
                    }
                }

                if (bMailFilesUpdated)
                {
                    lblMailFilesUpdated.BackColor = Color.MediumSeaGreen;
                    lblMailFilesUpdated.ForeColor = Color.Black;
                    if (bAllFilesUpdated)
                    {
                        btnBuild.BackColor = Color.MediumSeaGreen;
                        btnBuild.ForeColor = Color.Black;
                    }
                }
                else
                {
                    btnBuild.BackColor = Color.Tomato;
                    btnBuild.ForeColor = Color.Gainsboro;
                    lblBottomLineChecked.BackColor = Color.Tomato;
                    lblBottomLineChecked.ForeColor = Color.Gainsboro;
                    lblMailFilesUpdated.BackColor = Color.Tomato;
                    lblMailFilesUpdated.ForeColor = Color.Gainsboro;
                }
                
            }

            if (rbAutomatic.Checked)
            {
                if (iSignatureCounter >= 12)
                {
                    iSignatureCounter = 1;
                    if (GetTimeOfDay() == "Day")
                    {
                        rtxtDaySignature.Show();
                        rtxtNightSignature.Hide();
                    }
                    else
                    {
                        rtxtNightSignature.Show();
                        rtxtDaySignature.Hide();
                    }
                }
            }
        }        

        private bool CheckBottomLine()
        {
            lblProgress.Text = "Checking Bottom Line"; 
            lblProgress.Refresh();

            bool bBottomLineGood = false;
            bool bRedAlert = false;

            string strSurveyMessage = "Survey Depths did not match up. If the survey depths are correct then check the '# of Rows to Use in DD Survey File' textbox.  " + Environment.NewLine +
                                      "It needs to show the total number of surveys and projections used to create the email report.  I.E.  if you have one survey,  " + Environment.NewLine
                                      + "a bit projection and one projection beyond the bit then the value need to be 3.  If this is acceptable then click OK." + Environment.NewLine +
                                      "Otherwise click Cancel";

            PDDocument pdDoc = null;
            pdDoc = PDDocument.load(txtMWDSurveyFileName.Text);
            PDFTextStripper stripper = new PDFTextStripper();
            using (StringReader STREsurvFile = new StringReader(stripper.getText(pdDoc)))                
            {
                decimal dIncDiff = (-1);
                decimal dAzmDiff = (-1);
                decimal dTVDDiff = (-1);
                decimal dNorthDiff = (-1);
                decimal dEastDiff = (-1);
                decimal dVSDiff = (-1);
                bool bSurveysFound = false;
                string strTVDTemp;

                string strSurveyMD = dSurveyMD.ToString();

                if (!strSurveyMD.Contains("."))
                {
                    strSurveyMD = strSurveyMD + ".00";
                }

                string strMessage = "The Bottom Line at Survey Depth " + dSurveyMD.ToString() + " was not correct.";
                string templine;
                string strTemp;
                while ((templine = STREsurvFile.ReadLine()) != null)
                {
                    if (templine == "Well Head")
                    {
                        templine = STREsurvFile.ReadLine();
                        bSurveysFound = true;
                    }

                    if (bSurveysFound)
                    {
                        if (templine.Contains("Page"))
                        {
                            bSurveysFound = false;
                        }
                        else
                        {
                            if (templine.Contains(" "))
                            {
                                strTemp = templine.Substring(0, templine.IndexOf(" "));
                                if (strSurveyMD == strTemp)
                                {
                                    dIncDiff = Convert.ToDecimal(GetStringFromLine(1, templine, " ")) - dInc;
                                    dAzmDiff = Convert.ToDecimal(GetStringFromLine(2, templine, " ")) - dAzm;
                                    dNorthDiff = Convert.ToDecimal(GetStringFromLine(3, templine, " ")) - dNorth;
                                    dEastDiff = Convert.ToDecimal(GetStringFromLine(4, templine, " ")) - dEast;
                                    dVSDiff = Convert.ToDecimal(GetStringFromLine(5, templine, " ")) - dVS;
                                    strTVDTemp = GetStringFromLine(6, templine, " ");
                                    strTVDTemp = strTVDTemp.Substring(strTVDTemp.IndexOf(".") + 3, strTVDTemp.Length - (strTVDTemp.IndexOf(".") + 3));

                                    dTVDDiff = Convert.ToDecimal(strTVDTemp) - dTVD;
                                    lblProgress.Text = "Finished Checking Bottom Line";
                                    lblProgress.Refresh();
                                    break;
                                }
                            }
                        }
                    }
                }

                bBottomLineGood = true;
                strMessage = strMessage + Environment.NewLine;

                if (dIncDiff != 0)
                {
                    bBottomLineGood = false;
                    strMessage = strMessage + "  The Inc was off by " + dIncDiff + "." + Environment.NewLine;
                    bRedAlert = true;
                }
                if (dAzmDiff != 0)
                {
                    bBottomLineGood = false;
                    strMessage = strMessage + "  The Azm was off by " + dAzmDiff + "." + Environment.NewLine;
                    bRedAlert = true;
                }
                if (dTVDDiff != 0)
                {
                    bBottomLineGood = false;
                    strMessage = strMessage + "  The TVD was off by " + dTVDDiff + "." + Environment.NewLine;
                }
                if (dNorthDiff != 0)
                {
                    bBottomLineGood = false;
                    strMessage = strMessage + "  The North was off by " + dNorthDiff + "." + Environment.NewLine;
                }
                if (dEastDiff != 0)
                {
                    bBottomLineGood = false;
                    strMessage = strMessage + "  The East was off by " + dEastDiff + "." + Environment.NewLine;
                }
                if (dVSDiff != 0)
                {
                    bBottomLineGood = false;
                    strMessage = strMessage + "  The VS was off by " + dVSDiff + "." + Environment.NewLine;
                }

                strMessage = strMessage + Environment.NewLine +
                                    "  If these numbers are acceptable then click OK.  Otherwise click Cancel.";

                if (bSurveysFound)
                {
                    if (!bBottomLineGood)
                    {
                        //strMessage = strMessage + Environment.NewLine +
                        //            "  If these numbers are acceptable then click OK.  Otherwise click Cancel.";
                        frmBottomLineForm frmBLMessage = new frmBottomLineForm(bSurveysFound, bRedAlert, strSurveyMessage, strMessage);
                        //DialogResult drResult = MessageBox.Show(strMessage + "  If these numbers are acceptable then click OK.  Otherwise click Cancel.", "Bottom Line Check", MessageBoxButtons.OKCancel);

                        DialogResult drResult = frmBLMessage.ShowDialog();

                        if (drResult == DialogResult.OK)
                        {
                            bBottomLineGood = true;
                        }
                        else
                        {
                            bBottomLineGood = false;
                        }
                    }
                }
                else
                {
                    //strMessage = "Survey Depths did not match up.  If this is acceptable then click OK." + Environment.NewLine +
                    //             "Otherwise click Cancel";
                    frmBottomLineForm frmBLMessage = new frmBottomLineForm(bSurveysFound, bRedAlert, strSurveyMessage, strMessage);

                    //DialogResult drResult = MessageBox.Show("Survey Depths did not match up.  If this is acceptable then click OK.  Otherwise click Cancel", 
                    //                                        "Bottom Line Check", MessageBoxButtons.OKCancel);

                    DialogResult drResult = frmBLMessage.ShowDialog();

                    if (drResult == DialogResult.OK)
                    {
                        bBottomLineGood = true;
                    }
                    else
                    {
                        bBottomLineGood = false;
                    }
                }
                pdDoc.close();
                STREsurvFile.Close();
                STREsurvFile.Dispose();
            }
            
            return bBottomLineGood;
        }

        private bool ReadMWDFile()
        {
            try
            {
                //PDDocument pdDoc = null;
                ////pdDoc = PDDocument.load(txtMWDSurveyFileName.Text);
                //PDFTextStripper stripper = new PDFTextStripper();

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlBook = xlApp.Workbooks.Open(txtMWDSurveyFileName.Text, 0, true, 5, "", "", true, 
                                                            Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, 
                                                            "\t", false, false, 0, true, 1, 0);

                Excel.Worksheet xlSheet;
                xlSheet = (Excel.Worksheet)xlBook.Worksheets.get_Item(1);

                int iRows = 0;
                int iCols = 0;

                Excel.Range xlRange;

                xlRange = xlSheet.UsedRange;
                iRows = xlRange.Rows.Count;
                iCols = xlRange.Columns.Count;

                //int iCurrentRow = 7;
                //int iCurrentCol = 5;

                //MessageBox.Show(xlBook.Worksheets.Count.ToString());

                //MessageBox.Show((string)(xlRange.Cells[iCurrentRow, iCurrentCol] as Excel.Range).Text);
                
                string strSurvMD = "";
                string strSurvInc = "";
                string strSurvAzm = "";
                string strSurvTVD = "";
                string strSurvNorth = "";
                string strSurvNorthUnit = "";
                string strSurvEast = "";
                string strSurvEastUnit = "";
                string strSurvVS = "";
                string strSurvDLS = "";
                string strBitMD = "";
                string strBitInc = "";
                string strBitAzm = "";
                string strBitTVD = "";
                string strBitNorth = "";
                string strBitNorthUnit = "";
                string strBitEast = "";
                string strBitEastUnit = "";
                string strBitVS = "";
                bool bSurveysFound = false;
                bool bBitDepthFound = false;

                for (int i = 11; i < xlRange.Rows.Count - 2; i++)
                {
                    if (string.IsNullOrEmpty(((string)(xlRange.Cells[i + 1, 3] as Excel.Range).Text).ToString().Replace(" ", "")))
                    {
                        strSurvMD = ((string)(xlRange.Cells[i, 3] as Excel.Range).Text).Trim();
                        strSurvInc = ((string)(xlRange.Cells[i, 4] as Excel.Range).Text).Trim();
                        strSurvAzm = ((string)(xlRange.Cells[i, 5] as Excel.Range).Text).Trim();
                        strSurvTVD = ((string)(xlRange.Cells[i, 10] as Excel.Range).Text).Trim();
                        strSurvVS = ((string)(xlRange.Cells[i, 11] as Excel.Range).Text).Trim();
                        strSurvNorth = ((string)(xlRange.Cells[i, 12] as Excel.Range).Text).Trim();
                        strSurvNorthUnit = ((string)(xlRange.Cells[i, 13] as Excel.Range).Text).Trim();
                        strSurvEast = ((string)(xlRange.Cells[i, 14] as Excel.Range).Text).Trim();
                        strSurvEastUnit = ((string)(xlRange.Cells[i, 15] as Excel.Range).Text).Trim();
                        strSurvDLS = ((string)(xlRange.Cells[i, 18] as Excel.Range).Text).Trim();

                        bSurveysFound = true;

                        strBitMD = ((string)(xlRange.Cells[7, 5] as Excel.Range).Text).Trim();
                        strBitInc = ((string)(xlRange.Cells[7, 8] as Excel.Range).Text).Trim();
                        strBitAzm = ((string)(xlRange.Cells[7, 11] as Excel.Range).Text).Trim();
                        strBitTVD = ((string)(xlRange.Cells[7, 13] as Excel.Range).Text).Trim();
                        strBitVS = ((string)(xlRange.Cells[7, 16] as Excel.Range).Text).Trim();
                        strBitNorth = ((string)(xlRange.Cells[7, 18] as Excel.Range).Text).Trim();
                        strBitEast = ((string)(xlRange.Cells[7, 20] as Excel.Range).Text).Trim();

                        if (Convert.ToDecimal(strBitMD) > Convert.ToDecimal(strSurvMD))
                        {
                            bBitDepthFound = true;
                        }
                        break;
                    }
                }

                xlBook.Close();

                if (bSurveysFound)
                {
                    if (bBitDepthFound)
                    {
                        string strSurveyNorth;
                        string strSurveyEast;

                        DataRow drRow = dtVariables.NewRow();
                        drRow["Name"] = "Survey_MD";
                        drRow["Value"] = strSurvMD;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Survey_Inc";
                        drRow["Value"] = strSurvInc;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Survey_Azm";
                        drRow["Value"] = strSurvAzm;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Survey_TVD";
                        drRow["Value"] = strSurvTVD;
                        dtVariables.Rows.Add(drRow);
                        strSurveyNorth = strSurvNorth.ToString();
                        strSurveyEast = strSurvEast.ToString();

                        if (strSurvNorthUnit == "N")
                        {
                            strSurvNorthUnit = " North";
                        }
                        else
                        {
                            strSurvNorthUnit = " South";
                        }
                        if (strSurvEastUnit == "E")
                        {
                            strSurvEastUnit = " East";
                        }
                        else
                        {
                            strSurvEastUnit = " West";
                        }

                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Survey_North";
                        drRow["Value"] = Math.Abs(Convert.ToDouble(strSurveyNorth)).ToString();
                        drRow["Unit"] = strSurvNorthUnit;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Survey_East";
                        drRow["Value"] = Math.Abs(Convert.ToDouble(strSurveyEast)).ToString();
                        drRow["Unit"] = strSurvEastUnit;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Survey_VS";
                        drRow["Value"] = strSurvVS;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Survey_DLS";
                        drRow["Value"] = strSurvDLS;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();
                        dtVariables.Rows.Add(drRow);

                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Bit_MD";
                        drRow["Value"] = strBitMD;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Bit_Inc";
                        drRow["Value"] = strBitInc;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Bit_Azm";
                        drRow["Value"] = strBitAzm;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Bit_TVD";
                        drRow["Value"] = strBitTVD;
                        dtVariables.Rows.Add(drRow);

                        if (strBitNorth.Contains("N"))
                        {
                            strBitNorth = strBitNorth.Remove(strBitNorth.IndexOf("N"), 1);
                            strBitNorthUnit = " North";
                        }
                        else if (strBitNorth.Contains("S"))
                        {
                            strBitNorth = strBitNorth.Remove(strBitNorth.IndexOf("S"), 1);
                            strBitNorthUnit = " South";
                        }

                        if (strBitEast.Contains("E"))
                        {
                            strBitEast = strBitEast.Remove(strBitEast.IndexOf("E"), 1);
                            strBitEastUnit = " East";
                        }
                        else if (strBitEast.Contains("W"))
                        {
                            strBitEast = strBitEast.Remove(strBitEast.IndexOf("W"), 1);
                            strBitEastUnit = " West";
                        }

                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Bit_North";
                        drRow["Value"] = Math.Abs(Convert.ToDouble(strBitNorth)).ToString();
                        drRow["Unit"] = strBitNorthUnit;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Bit_East";
                        drRow["Value"] = Math.Abs(Convert.ToDouble(strBitEast)).ToString();
                        drRow["Unit"] = strBitEastUnit;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Bit_VS";
                        drRow["Value"] = strBitVS;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();

                        string strABUnit;
                        string strRLUnit;

                        string strAboveBelow = txtBitAboveBelow.Text;
                        string strRightLeft = txtBitRightLeft.Text;

                        if (Convert.ToDecimal(strAboveBelow) > 0)
                        {
                            strABUnit = " Below";
                        }
                        else if (Convert.ToDecimal(strAboveBelow) < 0)
                        {
                            strABUnit = " Above";
                        }
                        else
                        {
                            strABUnit = "";
                        }

                        if (Convert.ToDecimal(strRightLeft) > 0)
                        {
                            strRLUnit = " Left";
                        }
                        else if (Convert.ToDecimal(strRightLeft) < 0)
                        {
                            strRLUnit = " Right";
                        }
                        else
                        {
                            strRLUnit = "";
                        }

                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Bit_A/B_Plan";
                        drRow["Value"] = strAboveBelow;
                        drRow["Unit"] = strABUnit;
                        dtVariables.Rows.Add(drRow);
                        drRow = dtVariables.NewRow();
                        drRow["Name"] = "Bit_R/L_Plan";
                        drRow["Value"] = strRightLeft;
                        drRow["Unit"] = strRLUnit;
                        dtVariables.Rows.Add(drRow);

                        return true;
                    }
                    else
                    {
                        MessageBox.Show("There was a problem finding the Bit Depth in the MWD Excel File.");
                        return false;
                    }
                }
                else
                {
                    MessageBox.Show("There was a problem finding the Survey Depth in the MWD Excel File.");
                    return false;
                }                
            }
            catch(Exception ex)
            {
                if (!chkThirdParty.Checked)
                {
                    MessageBox.Show(ex.ToString() + Environment.NewLine +
                                    "There was an issue reading the MWD Survey PDF File.  It may need to be closed.");
                }
                else
                {
                    MessageBox.Show(ex.ToString() + Environment.NewLine +
                                    "There was an issue reading the MWD Survey Excel File.  It may need to be closed.");
                }
                return false;
            }            
        }

        protected virtual bool IsFileLocked(string strFilePath)
        {
            FileInfo fiFile = new FileInfo(strFilePath);

            FileStream stream = null;

            try
            {
                if (File.Exists(strFilePath))
                {
                    stream = fiFile.Open(FileMode.Open, FileAccess.Read, FileShare.None);
                }
                else
                {
                    FileDoesNotExistMessage(strFilePath);
                    return true;
                }

            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }

        private void FileDoesNotExistMessage(string strPath)
        {
            MessageBox.Show(Path.GetFileName(strPath) + " does not exist at it's specified location.");
        }

        private List<string> CheckAllFilesForLocked()
        {
            List<string> lstOldFiles = new List<string>();

            if (IsFileLocked(txtSurveyFileName.Text))
            {
                lstOldFiles.Add(txtSurveyFileName.Text);
            }
            

            if (chkBottomLineChecker.Checked)
            {
                if (IsFileLocked(txtMWDSurveyFileName.Text))
                {
                    lstOldFiles.Add(txtMWDSurveyFileName.Text);
                }
            }

            //for (int i = 0; i < rtxtMailFiles.Lines.Length; i++)
            //{
            //    if (IsFileLocked(txtMailFolder.Text + @"\" + rtxtMailFiles.Lines[i]))
            //    {
            //        lstOldFiles.Add(txtMailFolder.Text + @"\" + rtxtMailFiles.Lines[i].ToString());
            //    }
            //}

            if (!chkWithoutAC.Checked)
            {
                if (IsFileLocked(txtACFileName.Text))
                {
                    lstOldFiles.Add(txtACFileName.Text);
                }
            }

            return lstOldFiles;
        }

        private ArrayList CheckFilesUpdated()
        {
            lblProgress.Text = "Checking For Updated Files";
            lblProgress.Refresh();
            ArrayList arrFilesUpdated = new ArrayList();

            DateTime dtLastUpdate;
                
            

            TimeSpan tsDiff;



            double dDiffInMinutes;
            int iWholeMinutes;

            if (File.Exists(txtSurveyFileName.Text))
            {
                dtLastUpdate = File.GetLastAccessTime(txtSurveyFileName.Text);
                tsDiff = DateTime.Now - dtLastUpdate;

                dDiffInMinutes = tsDiff.TotalMinutes;
                iWholeMinutes = (int)dDiffInMinutes;

                if (iWholeMinutes > 10)
                {
                    arrFilesUpdated.Add(Path.GetFileName(txtSurveyFileName.Text));
                }
            }
            else
            {
                MessageBox.Show("File " + txtSurveyFileName.Text + " does not exist.  Update the Survey File Path with a valid file.");
            }

            if (!chkWithoutAC.Checked)
            {
                if (File.Exists(txtACFileName.Text))
                {
                    dtLastUpdate = File.GetLastAccessTime(txtACFileName.Text);

                    tsDiff = DateTime.Now - dtLastUpdate;
                    dDiffInMinutes = tsDiff.TotalMinutes;
                    iWholeMinutes = (int)dDiffInMinutes;

                    if (iWholeMinutes > 10)
                    {
                        arrFilesUpdated.Add(Path.GetFileName(txtACFileName.Text));
                    }
                }
                else
                {
                    MessageBox.Show("File " + txtACFileName.Text + " does not exist.  Update the Survey File Path with a valid file.");
                }
            }

            for (int i = 0; i < rtxtMailFiles.Lines.Length; i++)
            {
                if (!string.IsNullOrEmpty(rtxtMailFiles.Lines[i].Replace(" ", "")))
                {
                    if (File.Exists(txtMailFolder.Text + @"\" + rtxtMailFiles.Lines[i]))
                    {
                        dtLastUpdate = File.GetLastAccessTime(txtMailFolder.Text + @"\" + rtxtMailFiles.Lines[i]);

                        tsDiff = DateTime.Now - dtLastUpdate;
                        dDiffInMinutes = tsDiff.TotalMinutes;
                        iWholeMinutes = (int)dDiffInMinutes;

                        if (iWholeMinutes > 10)
                        {
                            dtLastUpdate = File.GetLastWriteTime(txtMailFolder.Text + @"\" + rtxtMailFiles.Lines[i]);

                            tsDiff = DateTime.Now - dtLastUpdate;
                            dDiffInMinutes = tsDiff.TotalMinutes;
                            iWholeMinutes = (int)dDiffInMinutes;
                            if (iWholeMinutes > 10)
                            {
                                arrFilesUpdated.Add(rtxtMailFiles.Lines[i]);
                            }
                        }
                    }

                    else
                    {
                        MessageBox.Show("File " + txtMailFolder.Text + @"\" + rtxtMailFiles.Lines[i] + " does not exist.  Update the Survey File Path with a valid file.");
                    }
                }
            }

            return arrFilesUpdated;
        }

        private bool IsReadyToSend()
        {
            lblProgress.Text = "Checking For Locked Files";
            lblProgress.Refresh();
            bool bOkToSend = false;

            //string strLockedFiles = CheckAllFilesForLocked();

            List<string> lstLockedFiles = CheckAllFilesForLocked();



            if (lstLockedFiles.Count == 0)
            {
                ArrayList arrFilesUpdated = new ArrayList();
                arrFilesUpdated = CheckFilesUpdated();

                if (arrFilesUpdated.Count == 0)
                {
                    if (chkBottomLineChecker.Checked)
                    {
                        bOkToSend = GetSurveyInfo();

                        if (bOkToSend)
                        {
                            bOkToSend = CheckBottomLine();
                            if (!bOkToSend)
                            {
                                lblBottomLineChecked.Text = "Bottom Line Wrong";
                                lblBottomLineChecked.BackColor = Color.Tomato;
                                lblBottomLineChecked.ForeColor = Color.Gainsboro;
                            }
                            else
                            {
                                dtBottomLine = DateTime.Now;
                                lblBottomLineChecked.Text = "Bottom Line Last Checked @ " + dtBottomLine.ToString();
                                lblBottomLineChecked.BackColor = Color.MediumSeaGreen;
                                lblBottomLineChecked.ForeColor = Color.Black;
                            }
                        }
                        else
                        {
                            MessageBox.Show("There was an error calculating the Survey Info from " + Path.GetFileName(txtSurveyFileName.Text));
                        }
                    }
                    else
                    {
                        bOkToSend = true;
                    }
                }
                else
                {
                    string strFiles = "";
                    for (int i = 0; i < arrFilesUpdated.Count; i++)
                    {
                        strFiles = strFiles + Environment.NewLine + arrFilesUpdated[i].ToString();
                    }
                    DialogResult drResult = MessageBox.Show("The following files have not been updated in the last 10 minutes." + strFiles + Environment.NewLine +
                                                            "If this is OK then press 'OK', otherwise press 'Cancel'.", "Files Not Updated.", MessageBoxButtons.OKCancel);
                    if (drResult == DialogResult.OK)
                    {
                        if (chkBottomLineChecker.Checked)
                        {
                            bOkToSend = GetSurveyInfo();

                            if (bOkToSend)
                            {
                                bOkToSend = CheckBottomLine();
                                if (!bOkToSend)
                                {
                                    lblBottomLineChecked.Text = "Bottom Line Wrong";
                                    lblBottomLineChecked.BackColor = Color.Tomato;
                                    lblBottomLineChecked.ForeColor = Color.Gainsboro;
                                }
                                else
                                {
                                    dtBottomLine = DateTime.Now;
                                    lblBottomLineChecked.Text = "Bottom Line Last Checked @ " + dtBottomLine.ToString();
                                    lblBottomLineChecked.BackColor = Color.MediumSeaGreen;
                                    lblBottomLineChecked.ForeColor = Color.Black;
                                }
                            }
                            else
                            {
                                MessageBox.Show("There was an error calculating the Survey Info from " + Path.GetFileName(txtSurveyFileName.Text));
                            }
                        }
                        else
                        {
                            bOkToSend = true;
                        }
                    }
                    else
                    {
                        lblProgress.Text = "";
                        lblProgress.Refresh();
                        bOkToSend = false;
                    }
                }
            }
            else
            {
                for (int j = 0; j < lstLockedFiles.Count; j++)
                {
                    if (!File.Exists(lstLockedFiles[j]))
                    {
                        return false;
                    }
                }

                string strMessage = lstLockedFiles[0] + Environment.NewLine;
                for (int i = 1; i < lstLockedFiles.Count; i++)
                {
                    strMessage = strMessage + lstLockedFiles[i] + Environment.NewLine;
                }
                lblProgress.Text = "Needed Files are Locked.";
                lblProgress.Refresh();
                MessageBox.Show(strMessage + " is still showing to be locked.  Check to make sure all files are closed and try again.");
                bOkToSend = false;
                lblProgress.Text = "";
                lblProgress.Refresh();
            }

            return bOkToSend;
        }

        private void CreateEMail()
        {
            Outlook.Application oApplication = new Outlook.Application();            
            Outlook._MailItem oMailItem = (Outlook._MailItem)oApplication.CreateItem(Outlook.OlItemType.olMailItem);
            
            oMailItem.Subject = strSubjectLine;

            //oMailItem.BodyFormat = Outlook.OlBodyFormat.olFormatRichText;
            oMailItem.To = rtxtRecipientList.Text;
            oMailItem.CC = rtxtCCRecipientList.Text;            

            if (Directory.Exists(txtMailFolder.Text))
            {
                OpenFileDialog ofdAttachment = new OpenFileDialog();
                foreach (string strLine in rtxtMailFiles.Lines)
                {
                    if (!string.IsNullOrEmpty(strLine.Replace(" ", "")))
                    {
                        ofdAttachment.FileName = txtMailFolder.Text + @"\" + strLine;
                        if (File.Exists(ofdAttachment.FileName))
                        {
                            oMailItem.Attachments.Add(ofdAttachment.FileName, Outlook.OlAttachmentType.olByValue, 1, Path.GetFileName(ofdAttachment.FileName));
                        }
                        else
                        {
                            MessageBox.Show(strLine + " does not exist in Folder " + txtMailFolder.Text + ".");
                        }
                    }
                }
            }           



            if (chkGenerateTable.Checked)
            {
                oMailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                oMailItem.HTMLBody = UpdateTableHTML() + oMailItem.HTMLBody;     
                
                oMailItem.Display(oMailItem);
                
            }
            else
            {
                oMailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                oMailItem.HTMLBody = UpdateBodyHTML() + oMailItem.HTMLBody;

                oMailItem.Display(oMailItem);
                //Clipboard.Clear();
                //rtxtTemplate.Text = rtxtTemplate.Text + rtxtSignature.Text;
                //rtxtTemplate.SelectAll();
                //rtxtTemplate.Copy();

                //oMailItem.BodyFormat = Outlook.OlBodyFormat.olFormatRichText;
                ////rtxtTemplate.Text = rtxtTemplate.Text + rtxtSignature.Text;
                ////rtxtSignature.SelectAll();
                ////oMailItem.RTFBody = rtxtSignature.SelectedRtf;
                //oMailItem.Display(oMailItem);
                ////oMailItem.Display(oMailItem);
            }


        }

        private string UpdateBodyHTML()
        {
            builder.Remove(0, builder.Length);
            
            builder.Append("<html>");
            builder.Append("<body>");

            //builder.Append("<style type='text/css'>.headerStyle { column-span:all; width:500px; text-align:center; }</style>");
            //builder.Append("<style type='text/css'>.multicol {  text-align:center; } .multicol h3 { column-span: all; }</style>");
            //builder.Append("newrow {  text-align:center; } h3 { column-span: all; }");

            //builder.Append("<style type='text/css'>.columnoneStyle { width:300px; text-align:center; }</style>");
            //builder.Append("<style type='text/css'>.columntwoStyle { width:100px; text-align:right; padding-right:15px }</style>");
            //builder.Append("<style type='text/css'>.columnthreeStyle { width:100px; text-align:left; padding-left:15px }</style>");

            List<string> lstCurrentVariable = new List<string>();
            string strCurrentUnit;
            string strLineHeaderFontFamily;
            string strDataSourceFontFamily;
            string strLineHeaderColor;
            string strDataSourceColor;
            string strLineHeaderFontSize;
            string strDataSourceFontSize;

            string strLineHeaderUnderlined;
            string strDataSourceUnderlined;
            string strLineHeaderBold;
            string strDataSourceBold;
            string strLineHeaderItalic;
            string strDataSourceItalic;

            int iBeyondCount = 0;
            int iFirstBeyondRow = 0;

            for (int i = 0; i < dtBodyTemplate.DefaultView.Count; i++)
            {
                strLineHeaderFontFamily = dtBodyTemplate.DefaultView[i]["LineHeaderFontName"].ToString();
                strDataSourceFontFamily = dtBodyTemplate.DefaultView[i]["DataSourceFontName"].ToString();
                strLineHeaderColor = GetRGBString(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderFontColor"]));
                strDataSourceColor = GetRGBString(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]));

                if (Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderBold"]))
                {
                    strLineHeaderBold = "bold";
                }
                else
                {
                    strLineHeaderBold = "normal";
                }

                if (Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceBold"]))
                {
                    strDataSourceBold = "bold";
                }
                else
                {
                    strDataSourceBold = "normal";
                }

                if (Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderUnderlined"]))
                {
                    strLineHeaderUnderlined = "underline";
                }
                else
                {
                    strLineHeaderUnderlined = "none";
                }

                if (Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceUnderlined"]))
                {
                    strDataSourceUnderlined = "underline";
                }
                else
                {
                    strDataSourceUnderlined = "none";
                }

                if (Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderItalic"]))
                {
                    strLineHeaderItalic = "italic";
                }
                else
                {
                    strLineHeaderItalic = "normal";
                }

                if (Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceItalic"]))
                {
                    strDataSourceItalic = "italic";
                }
                else
                {
                    strDataSourceItalic = "normal";
                }

                strLineHeaderFontSize = dtBodyTemplate.DefaultView[i]["LineHeaderFontSize"].ToString() + "pt";
                strDataSourceFontSize = dtBodyTemplate.DefaultView[i]["DataSourceFontSize"].ToString() + "pt";

                //builder.Append("<p>");

                if ((!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Replace(" ", ""))) &&
                    ((string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Replace(" ", ""))) ||
                    (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Beyond_Bit_Header")))
                {
                    if (iSVNumOfProj > 2)
                    {
                        if (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Beyond_Bit_Header")
                        {
                            if (iBeyondCount == 0)
                            {
                                iFirstBeyondRow = i - 1;
                            }
                            iBeyondCount = iBeyondCount + 1;
                        }
                        else if (!dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond"))
                        {
                            iBeyondCount = 0;
                        }
                    }
                    if (((dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() != "Comments") ||
                        ((dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() == "Comments") && (chkComments.Checked))) &&
                        ((!dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) ||
                        ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) && (iSVNumOfProj > 2))))
                    {
                        builder.Append("<span style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; color:" + strLineHeaderColor + "; " +
                                        "font-weight:" + strLineHeaderBold + "; text-decoration:" + strLineHeaderUnderlined + "; " +
                                        "font-style:" + strLineHeaderItalic + ";'>");
                        if ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Beyond_Bit_Header") &&
                                (i < (dtBodyTemplate.DefaultView.Count - 2)) &&
                                (dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString() == "Beyond_Bit_MD") &&
                                (!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString().Replace(" ", ""))))
                        {
                            lstCurrentVariable.Clear();
                            lstCurrentVariable = GetCurrentVariable("Beyond_Bit_MD", false, iBeyondCount);
                            string strDiffInProj = (Convert.ToDecimal(lstCurrentVariable[0]) - dBitMD).ToString();

                            builder.Append(strDiffInProj + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString());
                        }
                        else
                        {
                            //builder.Append("<tr><th colspan='3' bgcolor='strTableRowBackColor'>");
                            builder.Append(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString());

                            //builder.Append("<newrow colspan=3;>");
                            //builder.Append("<h3 colspan=3;>" + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() + "</h3><span>Testing This Out</span><span>And This</span><span>and this</span>");
                            //builder.Append("</newrow>");
                        }
                        builder.Append("</span>");
                    }
                }
                else if ((string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Replace(" ", ""))) &&
                        (!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Replace(" ", ""))) &&
                        (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Comments"))
                {
                    lstCurrentVariable.Clear();
                    lstCurrentVariable = GetCurrentVariable(dtBodyTemplate.DefaultView[i]["DataSource"].ToString(), false, iBeyondCount);

                    builder.Append("<span style=' font-family:" + strDataSourceFontFamily + "; font-size:" + strDataSourceFontSize + "; color:" + strDataSourceColor + "; " +
                                    "font-weight:" + strDataSourceBold + ";  text-decoration:" + strDataSourceUnderlined + "; " +
                                    "font-style:" + strDataSourceItalic + ";'>");
                    builder.Append(lstCurrentVariable[0]);
                    builder.Append("</span>");
                }
                else if (((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Beyond_Bit_Header")) &&
                        ((!dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) ||
                        ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) && (iSVNumOfProj > 2)) && (iFirstBeyondRow > 0)))

                {
                    if (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Comments")
                    {
                        lstCurrentVariable.Clear();
                        lstCurrentVariable = GetCurrentVariable(dtBodyTemplate.DefaultView[i]["DataSource"].ToString(), false, iBeyondCount);
                        if (lstCurrentVariable.Count == 2)
                        {
                            strCurrentUnit = dtBodyTemplate.DefaultView[i]["Unit"].ToString() + lstCurrentVariable[1].ToString();
                        }
                        else
                        {
                            strCurrentUnit = dtBodyTemplate.DefaultView[i]["Unit"].ToString();
                        }

                        string strSpaces;

                        if (dtBodyTemplate.DefaultView[i]["Spacing"].ToString() != "Static Indent")
                        {
                            strSpaces = HTMLSpaceGenerator(Convert.ToInt16(dtBodyTemplate.DefaultView[i]["Spacing"])) + lstCurrentVariable[0] + strCurrentUnit;
                        }
                        else
                        {
                            strSpaces = HTMLSpaceGenerator(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() +
                                                              lstCurrentVariable[0], Convert.ToInt16(dtBodyTemplate.DefaultView[i]["StaticIndent"])) +
                                                              lstCurrentVariable[0] + strCurrentUnit;
                        }

                        builder.Append(
                                        "<span style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; color:" + strLineHeaderColor + "; " +
                                        "font-weight:" + strLineHeaderBold + "; text-decoration:" + strLineHeaderUnderlined + "; " +
                                        "font-style:" + strLineHeaderItalic + ";'>" + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() + "</span>" +
                                        "<span style=' font-family:" + strDataSourceFontFamily + "; font-size:" + strDataSourceFontSize + "; color:" + strDataSourceColor + "; " +
                                        "font-weight:" + strDataSourceBold + "; text-decoration:" + strDataSourceUnderlined + "; " +
                                        "font-style:" + strDataSourceItalic + ";'>" + strSpaces + //"</span>" + //lstCurrentVariable[0] + "</span>" +
                                        //"<span style=' font-family:" + strDataSourceFontFamily + "; font-size:" + strDataSourceFontSize + "; color:" + strDataSourceColor + "; " +
                                        //"font-weight:" + strDataSourceBold + "; text-decoration:" + strDataSourceUnderlined + "; " +
                                        //"font-style:" + strDataSourceItalic + ";'>" + strCurrentUnit +
                                        "</span>");
                    }
                    else if ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Comments") && (chkComments.Checked))
                    {
                        //    builder.Append("</table>");
                        //    builder.Append("<style>table { width: 500px; border-collapse: collapse; } table, th, td { border: 1px solid black; } </style>");
                        //    builder.Append("<table>");
                        for (int j = 0; j < rtxtComments.Lines.Length; j++)
                        {
                            if (!string.IsNullOrEmpty(rtxtComments.Lines[j].Replace(" ", "")))
                            {
                                string strSpaces;

                                if (j > 0)
                                {
                                    strSpaces = HTMLSpaceGenerator(Convert.ToInt16(dtBodyTemplate.DefaultView[i]["Spacing"]));

                                    builder.Append("<span style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; " +
                                                   "color:" + strLineHeaderColor + "; font-weight:" + strLineHeaderBold + "; text-decoration:" + strLineHeaderUnderlined + "; " +
                                                   "font-style:" + strLineHeaderItalic + ";'>");
                                    builder.Append(strSpaces + rtxtComments.Lines[j].ToString());
                                    builder.Append("</span>");
                                }
                                else
                                {
                                    strSpaces = HTMLSpaceGenerator(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString(), Convert.ToInt16(dtBodyTemplate.DefaultView[i]["Spacing"]));

                                    builder.Append("<span style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; " +
                                                   "color:" + strLineHeaderColor + "; font-weight:" + strLineHeaderBold + "; text-decoration:" + strLineHeaderUnderlined + "; " +
                                                   "font-style:" + strLineHeaderItalic + ";'>");
                                    builder.Append(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() + strSpaces + rtxtComments.Lines[j].ToString());
                                    builder.Append("</span>");
                                }
                                builder.Append("<br>");
                                
                            }
                        }
                    }

                }

                if (Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]) > 0)
                {
                    for (int j = 0; j < Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]); j++)
                    {
                        builder.Append("<br>");//</br>");
                    }
                }

                if (iBeyondCount > 0)
                {
                    if (i <= (dtBodyTemplate.DefaultView.Count - 1))
                    {
                        if (((dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) &&
                            (i == (dtBodyTemplate.DefaultView.Count - 1))) ||
                            ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) &&
                            (!dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString().Contains("Beyond"))))
                        {
                            if ((iSVNumOfProj - iBeyondCount) > 2)
                            {
                                i = iFirstBeyondRow;

                                for (int j = 0; j < Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]); j++)
                                {
                                    builder.Append("<br></br>");
                                }
                            }
                        }
                    }
                }
                builder.Append("<br>");
                //builder.Append("</p>");
            }

            if (!chkWithoutAC.Checked)
            {
                ReadACFile();
            }

            RichTextBox rtxtRichTextSig = new RichTextBox();

            if (rbAutomatic.Checked)
            {
                if (GetTimeOfDay() == "Day")
                {
                    rtxtRichTextSig.Text = rtxtDaySignature.Text;
                }
                else
                {
                    rtxtRichTextSig.Text = rtxtNightSignature.Text;
                }
            }
            else if (rbDay.Checked)
            {
                rtxtRichTextSig.Text = rtxtDaySignature.Text;
            }
            else
            {
                rtxtRichTextSig.Text = rtxtNightSignature.Text;
            }

            if (!string.IsNullOrEmpty(rtxtRichTextSig.Text.Replace(" ", "")))
            {
                foreach (string strLine in rtxtRichTextSig.Lines)
                {
                    builder.Append("<br>");
                    string strTemp = "";
                    strTemp = strLine;
                    for (int j = 0; j < 100; j++)
                    {
                        if (strTemp.Contains("<"))
                        {
                            builder.Append(strTemp.Substring(0, strTemp.IndexOf("<")));
                            strTemp = strTemp.Substring(strTemp.IndexOf("<"));
                            string strHyperLink = "";
                            string strHyperLinkText = "";
                            if (strTemp.Contains("<mailto:"))
                            {
                                //builder.Append()
                                strTemp = strTemp.Substring(1);
                                strHyperLink = "<a href='" + strTemp.Substring(0, strTemp.IndexOf(">")) + "'>";
                                strTemp = strTemp.Substring("mailto:".Length);
                                strHyperLinkText = strTemp.Substring(0, strTemp.IndexOf(">"));
                                builder.Append(strHyperLink);
                                builder.Append(strHyperLinkText);
                                builder.Append("</a>");
                            }
                            else if (strTemp.Contains("<"))
                            {
                                strTemp = strTemp.Substring(1);
                                strHyperLink = "<a href='" + strTemp.Substring(0, strTemp.IndexOf(">")) + "'>";
                                strHyperLinkText = strTemp.Substring(0, strTemp.IndexOf(">"));
                                builder.Append(strHyperLink);
                                builder.Append(strHyperLinkText);
                                builder.Append("</a>");
                            }
                            strTemp = strTemp.Substring(strTemp.IndexOf(">") + 1);
                        }
                        else
                        {
                            builder.Append(strTemp);
                            break;
                        }
                    }
                    builder.Append("</br>");

                }
            }            

            builder.Append("</body>" +
                           "</html>");


            return builder.ToString();
        }

        private string UpdateTableHTML()
        {
            bool bTableOpen = false;

            builder.Remove(0, builder.Length);

            builder.Append("<html>");
            builder.Append("<body>");
            
            builder.Append("<style>table { width: 500px; border-collapse: separate; border: 5px solid black; } th, td { border-collapse: collapse; border: 1px solid black; } </style>");
            builder.Append("<table>");
            bTableOpen = true;

            builder.Append("<style type='text/css'>.headerStyle { column-span:all; width:500px; text-align:center; }</style>");

            builder.Append("<style type='text/css'>.columnoneStyle { width:300px; text-align:center; }</style>");
            builder.Append("<style type='text/css'>.columntwoStyle { width:100px; text-align:right; padding-right:15px }</style>");
            builder.Append("<style type='text/css'>.columnthreeStyle { width:100px; text-align:left; padding-left:15px }</style>");

            builder.Append("<style type='text/css'>.columnoneSplitStyle { width:250px; text-align:center; }</style>");
            builder.Append("<style type='text/css'>.columntwoSplitStyle { width:250px; text-align:center; }</style>");

            List<string> lstCurrentVariable = new List<string>();
            string strCurrentUnit;
            string strLineHeaderFontFamily;
            string strDataSourceFontFamily;
            string strLineHeaderColor;
            string strDataSourceColor;
            string strTableRowBackColor = GetRGBString(0);
            string strLineHeaderFontSize;
            string strDataSourceFontSize;

            string strLineHeaderUnderlined;
            string strDataSourceUnderlined;
            string strLineHeaderBold;
            string strDataSourceBold;
            string strLineHeaderItalic;
            string strDataSourceItalic;

            int iBeyondCount = 0;
            int iFirstBeyondRow = 0;

            for (int i = 0; i < dtBodyTemplate.DefaultView.Count; i++)
            {
                strLineHeaderFontFamily = dtBodyTemplate.DefaultView[i]["LineHeaderFontName"].ToString();
                strDataSourceFontFamily = dtBodyTemplate.DefaultView[i]["DataSourceFontName"].ToString();
                strLineHeaderColor = GetRGBString(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderFontColor"]));
                strDataSourceColor = GetRGBString(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]));
                strTableRowBackColor = GetRGBString(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["TableRowBackColor"]));

                if (Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderBold"]))
                {
                    strLineHeaderBold = "bold";
                }
                else
                {
                    strLineHeaderBold = "normal";
                }

                if (Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceBold"]))
                {
                    strDataSourceBold = "bold";
                }
                else
                {
                    strDataSourceBold = "normal";
                }

                if (Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderUnderlined"]))
                {
                    strLineHeaderUnderlined = "underline";
                }
                else
                {
                    strLineHeaderUnderlined = "none";
                }

                if (Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceUnderlined"]))
                {
                    strDataSourceUnderlined = "underline";
                }
                else
                {
                    strDataSourceUnderlined = "none";
                }

                if (Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderItalic"]))
                {
                    strLineHeaderItalic = "italic";
                }
                else
                {
                    strLineHeaderItalic = "normal";
                }

                if (Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceItalic"]))
                {
                    strDataSourceItalic = "italic";
                }
                else
                {
                    strDataSourceItalic = "normal";
                }

                strLineHeaderFontSize = dtBodyTemplate.DefaultView[i]["LineHeaderFontSize"].ToString() + "pt";
                strDataSourceFontSize = dtBodyTemplate.DefaultView[i]["DataSourceFontSize"].ToString() + "pt";


                if ((!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Replace(" ", ""))) &&
                    ((string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Replace(" ", ""))) ||
                    (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Beyond_Bit_Header")))
                {
                    if (iSVNumOfProj > 2)
                    {
                        if (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Beyond_Bit_Header")
                        {
                            if (iBeyondCount == 0)
                            {
                                iFirstBeyondRow = i - 1;
                            }
                            iBeyondCount = iBeyondCount + 1;
                        }
                        else if (!dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond"))
                        {
                            iBeyondCount = 0;
                        }
                    }
                    if (((dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() != "Comments") ||
                        ((dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() == "Comments") && (chkComments.Checked))) &&
                        ((!dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) ||
                        ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) && (iSVNumOfProj > 2))))
                    {
                        builder.Append("<tr><th colspan='3' style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; color:" + strLineHeaderColor + "; " +
                                        "font-weight:" + strLineHeaderBold + "; text-decoration:" + strLineHeaderUnderlined + "; " +
                                        "font-style:" + strLineHeaderItalic + "; background -color:" + strTableRowBackColor + "'>");
                        if ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Beyond_Bit_Header") &&
                                (i < (dtBodyTemplate.DefaultView.Count - 2)) &&
                                (dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString() == "Beyond_Bit_MD") &&
                                (!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString().Replace(" ", ""))))
                        {
                            lstCurrentVariable.Clear();
                            lstCurrentVariable = GetCurrentVariable("Beyond_Bit_MD", false, iBeyondCount);
                            string strDiffInProj = (Convert.ToDecimal(lstCurrentVariable[0]) - dBitMD).ToString();

                            builder.Append(strDiffInProj + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString());
                        }
                        else
                        {   
                            builder.Append(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString());
                        }
                        builder.Append("</th></tr>");
                    }
                }
                else if ((string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Replace(" ", ""))) &&
                        (!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Replace(" ", ""))) &&
                        (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Comments"))
                {
                    lstCurrentVariable.Clear();
                    lstCurrentVariable = GetCurrentVariable(dtBodyTemplate.DefaultView[i]["DataSource"].ToString(), false, iBeyondCount);

                    builder.Append("<tr><th colspan='3' style=' font-family:" + strDataSourceFontFamily + "; font-size:" + strDataSourceFontSize + "; color:" + strDataSourceColor + "; " +
                                    "font-weight:" + strDataSourceBold + ";  text-decoration:" + strDataSourceUnderlined + "; " +
                                    "font-style:" + strDataSourceItalic + "; background-color:" + strTableRowBackColor + "'>");
                    builder.Append(lstCurrentVariable[0]);
                    builder.Append("</th></tr>");                    
                }
                else if (((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Beyond_Bit_Header")) &&
                        ((!dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) ||
                        ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) && (iSVNumOfProj > 2)) && (iFirstBeyondRow > 0)))
                         
                {
                    if (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Comments")
                    {
                        lstCurrentVariable.Clear();
                        lstCurrentVariable = GetCurrentVariable(dtBodyTemplate.DefaultView[i]["DataSource"].ToString(), false, iBeyondCount);
                        if (lstCurrentVariable.Count == 2)
                        {
                            strCurrentUnit = dtBodyTemplate.DefaultView[i]["Unit"].ToString() + lstCurrentVariable[1].ToString();
                        }
                        else
                        {
                            strCurrentUnit = dtBodyTemplate.DefaultView[i]["Unit"].ToString();
                        }

                        if (dtBodyTemplate.DefaultView[i]["Unit"].ToString() != "")
                        {
                            builder.Append("<tr>" +
                                            "<td class='columnoneStyle'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; color:" + strLineHeaderColor + "; " +
                                            "font-weight:" + strLineHeaderBold + "; text-decoration:" + strLineHeaderUnderlined + "; " +
                                            "font-style:" + strLineHeaderItalic + "; background-color:" + strTableRowBackColor + "'>" + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() + "</td>" +
                                            "<td class='columntwoStyle'; style=' font-family:" + strDataSourceFontFamily + "; font-size:" + strDataSourceFontSize + "; color:" + strDataSourceColor + "; " +
                                            "font-weight:" + strDataSourceBold + "; text-decoration:" + strDataSourceUnderlined + "; " +
                                            "font-style:" + strDataSourceItalic + "; background-color:" + strTableRowBackColor + "'>" + lstCurrentVariable[0] + "</td>" +
                                            "<td class='columnthreeStyle'; style=' font-family:" + strDataSourceFontFamily + "; font-size:" + strDataSourceFontSize + "; color:" + strDataSourceColor + "; " +
                                            "font-weight:" + strDataSourceBold + "; text-decoration:" + strDataSourceUnderlined + "; " +
                                            "font-style:" + strDataSourceItalic + "; background-color:" + strTableRowBackColor + "'>" + strCurrentUnit + "</td>" +
                                            "</tr>");
                        }
                        else
                        {
                            builder.Append("<tr>" +
                                            "<th colspan='1' style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; color:" + strLineHeaderColor + "; " +
                                            "font-weight:" + strLineHeaderBold + "; text-decoration:" + strLineHeaderUnderlined + "; " +
                                            "font-style:" + strLineHeaderItalic + "; background-color:" + strTableRowBackColor + "'>" + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() + "</th>" +
                                            "<th colspan='2' style=' font-family:" + strDataSourceFontFamily + "; font-size:" + strDataSourceFontSize + "; color:" + strDataSourceColor + "; " +
                                            "font-weight:" + strDataSourceBold + "; text-decoration:" + strDataSourceUnderlined + "; " +
                                            "font-style:" + strDataSourceItalic + "; background-color:" + strTableRowBackColor + "'>" + lstCurrentVariable[0] +
                                            "</th></tr>");
                        }
                    }
                    else if ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Comments") && (chkComments.Checked))
                    {
                        builder.Append("<tr><th colspan='3' class='headerStyle'; style=' column-span:all; font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; " +
                                       "color:" + strLineHeaderColor + "; font-weight:" + strLineHeaderBold + "; text-decoration:" + strLineHeaderUnderlined + "; " +
                                       "font-style:" + strLineHeaderItalic + "; background-color:" + strTableRowBackColor + "'>");
                        builder.Append(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() + new string(' ', Convert.ToInt16(dtBodyTemplate.DefaultView[i]["Spacing"])) + rtxtComments.Text);
                        builder.Append("</th></tr>");
                    }
                    
                }


                if (((dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) && (iSVNumOfProj > 2)) ||
                            (!dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")))
                {
                    if (Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]) > 0)
                    {
                        if (bTableOpen)
                        {
                            builder.Append("</table>");
                            bTableOpen = false;
                        }

                        for (int j = 0; j < Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]); j++)
                        {
                            builder.Append("<br></br>");
                        }

                        if (i < (dtBodyTemplate.DefaultView.Count - 1))
                        {
                            if (((dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString().Contains("Beyond")) && (iSVNumOfProj > 2)) ||
                                (!dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString().Contains("Beyond")))
                            {
                                if (!bTableOpen)
                                {
                                    builder.Append("<table>");
                                    bTableOpen = true;
                                }
                            }
                        }
                    }
                }
                
                

                if (iBeyondCount > 0)
                {
                    if (i <= (dtBodyTemplate.DefaultView.Count - 1))
                    {
                        if (((dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) &&
                            (i == (dtBodyTemplate.DefaultView.Count - 1))) ||
                            ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) &&
                            (!dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString().Contains("Beyond"))))
                        {
                            if ((iSVNumOfProj - iBeyondCount) > 2)
                            {
                                i = iFirstBeyondRow;


                                if (bTableOpen)
                                {
                                    builder.Append("</table>");
                                    bTableOpen = false;
                                }

                                for (int j = 0; j < Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]); j++)
                                {
                                    builder.Append("<br></br>");
                                }

                                if (i < (dtBodyTemplate.DefaultView.Count - 1))
                                {
                                    if (!bTableOpen)
                                    {
                                        builder.Append("<table>");
                                        bTableOpen = true;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if (bTableOpen)
            {
                builder.Append("</table>");
                bTableOpen = false;
            }

            if (!chkWithoutAC.Checked)
            {
                ReadACFile();
            }

            builder.Append("<table style ='background-color:" + strTableRowBackColor + "'>");
            bTableOpen = true;
            bool bFirstRow = true;

            RichTextBox rtxtRichTextSig = new RichTextBox();

            if (rbAutomatic.Checked)
            {
                if (GetTimeOfDay() == "Day")
                {
                    rtxtRichTextSig.Text = rtxtDaySignature.Text;
                }
                else
                {
                    rtxtRichTextSig.Text = rtxtNightSignature.Text;
                }
            }
            else if (rbDay.Checked)
            {
                rtxtRichTextSig.Text = rtxtDaySignature.Text;
            }
            else
            {
                rtxtRichTextSig.Text = rtxtNightSignature.Text;
            }

            if (!string.IsNullOrEmpty(rtxtRichTextSig.Text.Replace(" ", "")))
            {
                foreach (string strLine in rtxtRichTextSig.Lines)
                {
                    if (!bFirstRow)
                    {
                        builder.Append("<br>");
                    }
                    bFirstRow = false;
                    string strTemp = "";
                    strTemp = strLine;
                    for (int j = 0; j < 100; j++)
                    {                        
                        if (strTemp.Contains("<"))
                        {
                            builder.Append(strTemp.Substring(0, strTemp.IndexOf("<")));
                            strTemp = strTemp.Substring(strTemp.IndexOf("<"));
                            string strHyperLink = ""; 
                            string strHyperLinkText = "";
                            if (strTemp.Contains("<mailto:"))
                            {
                                strTemp = strTemp.Substring(1);
                                strHyperLink = "<a href='" + strTemp.Substring(0, strTemp.IndexOf(">")) + "'>";
                                strTemp = strTemp.Substring("mailto:".Length);
                                strHyperLinkText = strTemp.Substring(0, strTemp.IndexOf(">"));
                                builder.Append(strHyperLink);
                                builder.Append(strHyperLinkText);
                                builder.Append("</a>");
                            }
                            else if (strTemp.Contains("<"))
                            {
                                strTemp = strTemp.Substring(1);
                                strHyperLink = "<a href='" + strTemp.Substring(0, strTemp.IndexOf(">")) + "'>";
                                strHyperLinkText = strTemp.Substring(0, strTemp.IndexOf(">"));
                                builder.Append(strHyperLink);
                                builder.Append(strHyperLinkText);
                                builder.Append("</a>");
                            }
                            strTemp = strTemp.Substring(strTemp.IndexOf(">") + 1);
                        }
                        else
                        {
                            builder.Append(strTemp);
                            break;
                        }
                    }
                    //builder.Append("</br>");
                }
            }

            rtxtRichTextSig.Dispose();

            if (bTableOpen)
            {
                builder.Append("</table>");
                bTableOpen = false;
            }

            builder.Append("</body>" +
                           "</html>");


            return builder.ToString();
        }

        private string GetRGBString(Int32 intColor)
        {
            string strRGB;
            try
            {
                strRGB = "RGB(";

                strRGB = strRGB + System.Drawing.Color.FromArgb(intColor).R.ToString();
                strRGB = strRGB + ",";
                strRGB = strRGB + System.Drawing.Color.FromArgb(intColor).G.ToString();
                strRGB = strRGB + ",";
                strRGB = strRGB + System.Drawing.Color.FromArgb(intColor).B.ToString();
                strRGB = strRGB + ")";
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error occurred in the GetRGBString routine.");
                strRGB = "RGB(255,255,255)";
            }

            return strRGB;
        }

        private void btnBuild_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(cboTemplates.Text.Replace(" ", "")))
            {
                bool bOkToSend = false;
                intSFDangerCount = 0;
                dtVariables.Rows.Clear();

                if (IsReadyToSend())
                {

                    lblBottomLineChecked.BackColor = Color.MediumSeaGreen;
                    lblBottomLineChecked.ForeColor = Color.Black;

                    bOkToSend = true;
                }

                if (bOkToSend)
                {
                    rtxtTemplate.Clear();

                    if (!chkThirdParty.Checked)
                    {
                        if (ReadSurveyFile())
                        {

                            UpdateBodyRichTextBox();

                            UpdateSubjectTextBox();

                            if (chkWithoutAC.Checked == false)
                            {
                                /////ReadACFile();
                            }

                            UpdateBodyRichTextBoxFont();

                            if (intSFDangerCount > 0)
                            {
                                int intStart = 0;
                                for (int i = 0; i < intSFDangerCount; i++)
                                {
                                    intStart = ColorText(intStart);
                                }
                            }

                            if (chkGenerateEmail.Checked)
                            {
                                CreateEMail();
                            }
                            else
                            {
                                frmBodySample frmSample = new frmBodySample(strSubjectLine, rtxtTemplate.Rtf);

                                frmSample.ShowDialog();

                                frmSample.Dispose();
                            }
                        }
                    }
                    else
                    {
                        if (ReadMWDFile())
                        {
                            UpdateBodyRichTextBox();

                            UpdateBodyRichTextBoxFont();

                            UpdateSubjectTextBox();

                            if (chkGenerateEmail.Checked)
                            {
                                CreateEMail();
                            }
                            else
                            {
                                frmBodySample frmSample = new frmBodySample(strSubjectLine, rtxtTemplate.Rtf);

                                frmSample.ShowDialog();

                                frmSample.Dispose();
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("A Template must first be selected/created before clicking 'Build Template'.");
            }
        }

        private List<string> GetCurrentVariable(string strDataSource, bool bSubjectLine, int iBeyondCount = 0)
        {
            List<string> arrList = new List<string>();

            if (string.IsNullOrEmpty(strDataSource.Replace(" ", "")))
            {
                arrList.Add("");                
                return arrList;
            }

            for (int i = 0; i < dtVariables.Rows.Count; i++)
            {                              
                if ((dtVariables.Rows[i]["Name"].ToString() == strDataSource) ||
                    ((strDataSource == "Motor_Yield") && (dtVariables.Rows[i]["Name"].ToString() == "Survey_DLS")) ||
                    ((strDataSource.Contains("Beyond")) && (dtVariables.Rows[i]["Name"].ToString() == "Beyond_Bit_Header"))) 
                {
                    if ((strDataSource != "Motor_Yield") && (!strDataSource.Contains("Beyond")))
                    {
                        if ((bSubjectLine) && (strDataSource.Contains("North")))
                        {
                            if ((dtVariables.Rows[i]["Unit"].ToString().Contains("South"))  || (dtVariables.Rows[i]["Unit"].ToString().Contains("West")))
                            {
                                if (Convert.ToDecimal(dtVariables.Rows[i]["Value"]) < 0)
                                {
                                    arrList.Add((Convert.ToDecimal(dtVariables.Rows[i]["Value"]) * (-1)).ToString());
                                }
                                else
                                {
                                    arrList.Add(dtVariables.Rows[i]["Value"].ToString());
                                }
                                arrList.Add(dtVariables.Rows[i]["Unit"].ToString());
                            }
                            else
                            {
                                arrList.Add(dtVariables.Rows[i]["Value"].ToString());
                                arrList.Add(dtVariables.Rows[i]["Unit"].ToString());
                            }
                        }
                        else
                        {
                            arrList.Add(dtVariables.Rows[i]["Value"].ToString());
                            arrList.Add(dtVariables.Rows[i]["Unit"].ToString());
                        }
                        return arrList;
                    }
                    else if (strDataSource.Contains("Beyond"))
                    {
                        if (dtVariables.Rows[i]["Name"].ToString() == "Beyond_Bit_Header")
                        {
                            if (Convert.ToInt16(dtVariables.Rows[i]["Value"]) == iBeyondCount)
                            {
                                for (int j = (i + 1); j < dtVariables.Rows.Count; j++)
                                {
                                    if (dtVariables.Rows[j]["Name"].ToString() == strDataSource)
                                    {
                                        arrList.Add(Convert.ToDecimal(dtVariables.Rows[j]["Value"]).ToString());
                                        arrList.Add(dtVariables.Rows[j]["Unit"].ToString());
                                        return arrList;
                                    }
                                }
                            }
                        }
                    }
                    else if (strDataSource == "Motor_Yield")
                    {
                        if (!string.IsNullOrEmpty(txtSurveyCourseLength.Text.Replace(" ", "")))
                        {
                            if (!string.IsNullOrEmpty(txtSlideSeen.Text.Replace(" ", "")))
                            {
                                arrList.Add(((Convert.ToDouble(dtVariables.Rows[i]["Value"]) / 
                                              Convert.ToDouble(txtSlideSeen.Text)) * 
                                              Convert.ToDouble(txtSurveyCourseLength.Text)).ToString("0.00"));
                            }
                            else
                            {
                                MessageBox.Show("Motor Yield cannot be calculated without 'Slide Seen'.");
                                arrList.Add("");                                
                            }
                        }
                        else
                        {
                            MessageBox.Show("Motor Yield cannot be calculated without 'Survey Course Length'.");
                            arrList.Add("");
                        }
                    }
                }
            }

            if (strDataSource == "Activity")
            {
                arrList.Add(txtActivity.Text);
                return arrList;
            }
            else if (strDataSource == "Target_Inc")
            {
                arrList.Add(txtTargetInc.Text);
                return arrList;
            }
            else if (strDataSource == "Slide_Tool_Face")
            {
                Decimal dTF = Convert.ToDecimal(txtSlideTF.Text);
                string strUnit = "";

                if ((dTF < 0))
                {
                    dTF = Math.Abs(360 + dTF);
                    strUnit = " L";
                }
                else if ((dTF > 180))
                {
                    dTF = Math.Abs(360 - dTF);
                    strUnit = " L";
                }
                else
                {
                    strUnit = " R";
                }
                arrList.Add(dTF.ToString());
                arrList.Add(strUnit);
                return arrList;
            }
            else if (strDataSource == "Slide_Rotate_Footage")
            {
                arrList.Add(txtSlideSeen.Text + " / " + txtSurveyCourseLength.Text);
                return arrList;
            }
            else if (strDataSource == "Gamma")
            {
                arrList.Add(txtGamma.Text);
                return arrList;
            }
            else if (strDataSource == "Surface_RPM")
            {
                arrList.Add(txtSurfaceRPM.Text);
                return arrList;
            }
            else if (strDataSource == "Rotate_Weight")
            {
                arrList.Add(txtRotateWeight.Text);
                return arrList;
            }
            else if (strDataSource == "DLS_Needed")
            {
                arrList.Add(txtDLSNeeded.Text);
                return arrList;
            }
            else if (strDataSource == "Survey_Temp")
            {
                arrList.Add(txtSurveyTemp.Text);
                return arrList;
            }
            else if (strDataSource == "Job_Number")
            {
                arrList.Add(txtJobNumber.Text);
                return arrList;
            }
            else if (strDataSource == "Rig")
            {
                arrList.Add(txtRig.Text);
                return arrList;
            }
            else if (strDataSource == "Rotary_Torque")
            {
                arrList.Add(txtRotaryTorque.Text);
                return arrList;
            }
            else if (strDataSource == "Target_Azm")
            {
                arrList.Add(txtTargetAzm.Text);
                return arrList;
            }
            else if (strDataSource == "ROP_Rotating")
            {
                arrList.Add(txtROPRotating.Text);
                return arrList;
            }
            else if (strDataSource == "SPP")
            {
                arrList.Add(txtSPP.Text);
                return arrList;
            }
            else if (strDataSource == "Pick_Up")
            {
                arrList.Add(txtPickUp.Text);
                return arrList;
            }
            else if (strDataSource == "WOB")
            {
                arrList.Add(txtWOB.Text);
                return arrList;
            }
            else if (strDataSource == "CM_Temp")
            {
                arrList.Add(txtCMTemp.Text);
                return arrList;
            }
            else if (strDataSource == "County")
            {
                arrList.Add(txtCounty.Text);
                return arrList;
            }
            else if (strDataSource == "Rev/Gal")
            {
                arrList.Add(txtRPG.Text);
                return arrList;
            }
            else if (strDataSource == "Flow_Rate")
            {
                arrList.Add(txtFlowRate.Text);
                return arrList;
            }
            else if (strDataSource == "ROP_Sliding")
            {
                arrList.Add(txtROPSliding.Text);
                return arrList;
            }
            else if (strDataSource == "Differential_Pressure")
            {
                arrList.Add(txtDifferentialPressure.Text);
                return arrList;
            }
            else if (strDataSource == "Slack_Off")
            {
                arrList.Add(txtSlackOff.Text);
                return arrList;
            }
            else if (strDataSource == "Slide_Ahead")
            {
                arrList.Add(txtSlideAhead.Text);
                return arrList;
            }
            else if (strDataSource == "Slide_Seen")
            {
                arrList.Add(txtSlideSeen.Text);
                return arrList;
            }
            else if(strDataSource == "Tool_Face_Mode")
            {
                arrList.Add(txtTFMode.Text);
                return arrList;
            }
            else if (strDataSource == "Client")
            {
                arrList.Add(txtClient.Text);
                return arrList;
            }
            else if (strDataSource == "Plan_Number")
            {
                arrList.Add(txtPlanNumber.Text);
                return arrList;
            }
            else if (strDataSource == "Well_Name")
            {
                arrList.Add(txtWellName.Text);
                return arrList;
            }
            else if (strDataSource == "Mud_Weight")
            {
                arrList.Add(txtMudWeight.Text);
                return arrList;
            }
            else if (strDataSource == "ESD")
            {
                arrList.Add(txtESD.Text);
                return arrList;
            }
            else if (strDataSource == "ECD")
            {
                arrList.Add(txtECD.Text);
                return arrList;
            }
            else if (strDataSource == "Motor_RPM")
            {
                if (!string.IsNullOrEmpty(txtFlowRate.Text.Replace(" ", "")))
                {
                    if (!string.IsNullOrEmpty(txtRPG.Text.Replace(" ", "")))
                    {
                        arrList.Add((Convert.ToDouble(txtRPG.Text) * Convert.ToDouble(txtFlowRate.Text)).ToString());
                    }
                    else
                    {
                        MessageBox.Show("Motor RPM cannot be calculated without 'Rev/Gal'.");
                        arrList.Add("");
                    }
                }
                else
                {    
                    MessageBox.Show("Motor RPM cannot be calculated without 'Flow Rate'.");
                    arrList.Add("");
                }
            }
            else if (strDataSource == "Bit_RPM")
            {
                if (!string.IsNullOrEmpty(txtFlowRate.Text.Replace(" ", "")))
                {
                    if (!string.IsNullOrEmpty(txtRPG.Text.Replace(" ", "")))
                    {                        
                        if (!string.IsNullOrEmpty(txtSurfaceRPM.Text.Replace(" ", "")))
                        {
                            arrList.Add(((Convert.ToDouble(txtRPG.Text) * Convert.ToDouble(txtFlowRate.Text)) + Convert.ToDouble(txtSurfaceRPM.Text)) .ToString());
                        }
                        else
                        {
                            MessageBox.Show("Bit RPM cannot be calculated without 'Surface RPM'.");
                            arrList.Add("");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Bit RPM cannot be calculated without 'Rev/Gal'.");
                        arrList.Add("");
                    }
                }
                else
                {
                    MessageBox.Show("Bit RPM cannot be calculated without 'Flow Rate'.");
                    arrList.Add("");
                }
            }

            arrList.Add("");
            return arrList;
        }

        private void UpdateSubjectTextBox()
        {
            if (dtSubjectTemplate.DefaultView.Count > 0)
            {
                strSubjectLine = "";
                List<string> lstCurrentVariable = new List<string>();
                for (int i = 0; i < dtSubjectTemplate.DefaultView.Count; i++)
                {
                    lstCurrentVariable.Clear();
                    lstCurrentVariable = GetCurrentVariable(dtSubjectTemplate.DefaultView[i]["DataSource"].ToString(), true);
                    strSubjectLine = strSubjectLine + lstCurrentVariable[0] + dtSubjectTemplate.DefaultView[i]["Separator"].ToString();
                }
            }
        }

        private void UpdateBodyRichTextBox()
        {
            if (dtBodyTemplate.DefaultView.Count > 0)
            {
                int iStart = 0;
                int iLength = 0;
                List<string> lstCurrentVariable = new List<string>();
                string strCurrentUnit;
                rtxtTemplate.Text = "";
                int iBeyondCount = 0;
                int iFirstBeyondRow = 0;
                for (int i = 0; i < dtBodyTemplate.DefaultView.Count; i++)
                {
                    if (iSVNumOfProj > 2)
                    {
                        if (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Beyond_Bit_Header")
                        {
                            if (iBeyondCount == 0)
                            {
                                iFirstBeyondRow = i - 1;
                            }
                            iBeyondCount = iBeyondCount + 1;
                        }
                        else if (!dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond"))
                        {
                            iBeyondCount = 0;
                        }
                    }
                        
                    if (((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Comments") ||
                        ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Comments") && (chkComments.Checked))) && 
                        ((!dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) ||
                        ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) && (iSVNumOfProj > 2))))
                    {
                        if (rtxtTemplate.Text.Length > 0)
                        {
                            iStart = rtxtTemplate.Text.Length;
                        }
                        else
                        {
                            iStart = 0;
                        }
                        if ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Comments") ||
                            ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Comments") && (chkComments.Checked)))
                        {
                            if ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Beyond_Bit_Header") &&
                                (i < (dtBodyTemplate.DefaultView.Count - 2)) &&
                                (dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString() == "Beyond_Bit_MD") &&
                                (!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString().Replace(" ", ""))))
                            {
                                lstCurrentVariable.Clear();
                                lstCurrentVariable = GetCurrentVariable("Beyond_Bit_MD", false, iBeyondCount);
                                string strDiffInProj = (Convert.ToDecimal(lstCurrentVariable[0]) - dBitMD).ToString();

                                rtxtTemplate.Text = rtxtTemplate.Text + strDiffInProj + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString();
                                iLength = strDiffInProj.Length + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Length;
                            }
                            else
                            {
                                rtxtTemplate.Text = rtxtTemplate.Text + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString();
                                iLength = dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Length;
                            }
                        }

                        if ((!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Replace(" ", ""))) &&
                            (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Beyond_Bit_Header"))
                        {
                            lstCurrentVariable.Clear();
                            lstCurrentVariable = GetCurrentVariable(dtBodyTemplate.DefaultView[i]["DataSource"].ToString(), false, iBeyondCount);
                            if (lstCurrentVariable.Count == 2)
                            {
                                strCurrentUnit = dtBodyTemplate.DefaultView[i]["Unit"].ToString() + lstCurrentVariable[1].ToString();
                            }
                            else
                            {
                                strCurrentUnit = dtBodyTemplate.DefaultView[i]["Unit"].ToString();
                            }
                            iStart = rtxtTemplate.Text.Length;

                            if (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Comments")
                            {
                                if (dtBodyTemplate.DefaultView[i]["Spacing"].ToString() != "Static Indent")
                                {
                                    iStart = iStart + Convert.ToInt16(dtBodyTemplate.DefaultView[i]["Spacing"]);
                                    rtxtTemplate.Text = rtxtTemplate.Text + new string(' ', Convert.ToInt16(dtBodyTemplate.DefaultView[i]["Spacing"]));
                                    rtxtTemplate.Text = rtxtTemplate.Text + lstCurrentVariable[0] + strCurrentUnit;
                                }
                                else
                                {
                                    rtxtTemplate.Text = rtxtTemplate.Text + SpaceGenerator(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() +
                                                                                           lstCurrentVariable[0], Convert.ToInt16(dtBodyTemplate.DefaultView[i]["StaticIndent"])) +
                                                                                           lstCurrentVariable[0] + strCurrentUnit;
                                }
                            }
                            else
                            {
                                if (chkComments.Checked)
                                {
                                    rtxtTemplate.Text = rtxtTemplate.Text + new string(' ', Convert.ToInt16(dtBodyTemplate.DefaultView[i]["Spacing"])) + rtxtComments.Lines[0];

                                    if (rtxtComments.Lines.Length > 1)
                                    {
                                        rtxtTemplate.Text = rtxtTemplate.Text + Environment.NewLine;
                                    }

                                    for (int j = 1; j < rtxtComments.Lines.Length; j++)
                                    {
                                        rtxtTemplate.Text = rtxtTemplate.Text + new string(' ', Convert.ToInt16(dtBodyTemplate.DefaultView[i]["Spacing"]) + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Length) +
                                                            rtxtComments.Lines[j].ToString();
                                        if (j < (rtxtComments.Lines.Length - 1))
                                        {
                                            rtxtTemplate.Text = rtxtTemplate.Text + Environment.NewLine;
                                        }
                                    }
                                }
                            }

                            iLength = rtxtTemplate.Text.Length - iStart;
                        }
                        for (int j = 0; j <= Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]); j++)
                        {
                            rtxtTemplate.Text = rtxtTemplate.Text + Environment.NewLine;
                        }
                    }

                    if (iBeyondCount > 0)
                    {
                        if (i < (dtBodyTemplate.DefaultView.Count - 1))
                        {
                            if ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) &&
                                ((!dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString().Contains("Beyond")) || (i == (dtBodyTemplate.DefaultView.Count - 1))))
                            {
                                if ((iSVNumOfProj - iBeyondCount) > 2)
                                {
                                    i = iFirstBeyondRow;
                                }
                            }
                        }
                    }
                }
                rtxtTemplate.Refresh();
            }
        }

        private void UpdateBodyRichTextBoxFont()
        {
            int iStart = 0;
            int iLength = 0;
            int iTotalIndexLength = 0;
            //int iTotalStringLength = 0;
            int iLines = 0;
            string strDataSourceText;
            string strCurrentUnit;
            int iCommentsStart = 0;
            int iCommentsLength = 0;
            int iCommentsLineCount = 0;
            int iCommentsLineHeaderLength = 0;
            int iBeyondCount = 0;
            int iFirstBeyondRow = 0;

            string strDiffInProj = "";
            List<string> lstCurrentVariable = new List<string>();
            for (int i = 0; i < dtBodyTemplate.DefaultView.Count; i++)
            {
                if (iSVNumOfProj > 2)
                {
                    if (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Beyond_Bit_Header")
                    {
                        if (iBeyondCount == 0)
                        {
                            iFirstBeyondRow = i - 1;
                        }
                        iBeyondCount = iBeyondCount + 1;
                    }
                    else if (!dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond"))
                    {
                        iBeyondCount = 0;
                    }
                }

                if (((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Comments") ||
                    ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Comments") && (chkComments.Checked))) &&
                    ((!dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) ||
                    ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) && (iSVNumOfProj > 2))))
                {
                    iStart = 0;
                    iLength = 0;
                    if (rtxtTemplate.Lines[iLines].Contains(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString()))
                    //if (rtxtTemplate.Lines[iLines].Contains(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString()))
                    {
                        ////////////////if (dtBodyTemplate.DefaultView[i]["Spacing"].ToString() == "Static Indent")
                        ////////////////{
                        ////////////////    //Applies the Font Setting for LineHeader to the whole line when Static Indent is selected. 
                        ////////////////    iStart = 0;//rtxtBodySample.GetFirstCharIndexFromLine(iLines);
                        ////////////////    iLength = rtxtTemplate.Lines[iLines].Length;
                        ////////////////}
                        ////////////////else
                        ////////////////{
                            if (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Comments")
                            {
                                if ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString() == "Beyond_Bit_Header") &&
                                    (i < (dtBodyTemplate.DefaultView.Count - 2)) &&
                                    (dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString() == "Beyond_Bit_MD") &&
                                    (!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString().Replace(" ", ""))))
                                {                                    
                                    lstCurrentVariable.Clear();
                                    lstCurrentVariable = GetCurrentVariable("Beyond_Bit_MD", false, iBeyondCount);
                                    strDiffInProj = (Convert.ToDecimal(lstCurrentVariable[0]) - dBitMD).ToString();

                                    iStart = Convert.ToInt16(rtxtTemplate.Lines[iLines].IndexOf(strDiffInProj + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Trim(' '), 0));
                                    iLength = strDiffInProj.Length + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Trim(' ').Length;
                                }
                                else
                                {
                                    iStart = Convert.ToInt16(rtxtTemplate.Lines[iLines].IndexOf(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Trim(' '), 0));
                                    iLength = dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Trim(' ').Length;
                                }
                            }
                            else
                            {
                                if (chkComments.Checked)
                                {
                                    iCommentsLineCount = rtxtComments.Lines.Length;
                                    iCommentsLineHeaderLength = dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Length;
                                    iStart = Convert.ToInt16(rtxtTemplate.Lines[iLines].IndexOf(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Trim(' '), 0));
                                    iLength = dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Trim(' ').Length;
                                    iCommentsStart = iStart;
                                    iCommentsLength = iLength;
                                }
                            }
                        ////////////////}

                        if (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Comments")
                        {
                            AddFontToRichText(iStart + iTotalIndexLength, iLength,
                                            dtBodyTemplate.DefaultView[i]["LineHeaderFontName"].ToString(),
                                            Convert.ToInt16(dtBodyTemplate.DefaultView[i]["LineHeaderFontSize"]),
                                            Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderFontColor"]),
                                            Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderUnderlined"]),
                                            Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderBold"]),
                                            Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderItalic"]));                            
                        }
                        else
                        {
                            if (chkComments.Checked)
                            {
                                int iLinesTemp = iLines;
                                int iTotalIndexLengthTemp = iTotalIndexLength;
                                for (int j = 0; j < rtxtComments.Lines.Length; j++)
                                {
                                    AddFontToRichText(iStart + iTotalIndexLengthTemp, iLength,
                                                dtBodyTemplate.DefaultView[i]["LineHeaderFontName"].ToString(),
                                                Convert.ToInt16(dtBodyTemplate.DefaultView[i]["LineHeaderFontSize"]),
                                                Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderFontColor"]),
                                                Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderUnderlined"]),
                                                Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderBold"]),
                                                Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["LineHeaderItalic"]));


                                    iTotalIndexLengthTemp = iTotalIndexLengthTemp + rtxtTemplate.Lines[iLinesTemp].Length + 1;// + Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]);
                                    iLinesTemp = iLinesTemp + 1;
                                }
                            }
                        }

                        if (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Beyond_Bit_Header")
                        {
                            //Needed to handle Static Indent when colors, underlined, bold or italic are different.  Font settings must be the same but other settings can be different.   
                            if (dtBodyTemplate.DefaultView[i]["Spacing"].ToString() == "Static Indent")
                            {
                                ////////////////if (Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderFontColor"]) !=
                                ////////////////    Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]))
                                ////////////////{
                                ////////////////    iStart = Convert.ToInt16(rtxtTemplate.Lines[iLines].IndexOf(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Trim(' '), 0));
                                ////////////////    rtxtTemplate.SelectionStart = iStart + iTotalIndexLength;
                                ////////////////    rtxtTemplate.SelectionLength = dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Trim(' ').Length;
                                ////////////////    rtxtTemplate.SelectionColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["LineHeaderFontColor"]));

                                strDataSourceText = rtxtTemplate.Lines[iLines].Substring(iStart + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Length);

                                lstCurrentVariable.Clear();
                                lstCurrentVariable = GetCurrentVariable(dtBodyTemplate.DefaultView[i]["DataSource"].ToString(), false, iBeyondCount);

                                if (lstCurrentVariable.Count == 2)
                                {
                                    strCurrentUnit = dtBodyTemplate.DefaultView[i]["Unit"].ToString() + lstCurrentVariable[1];
                                }
                                else
                                {
                                    strCurrentUnit = dtBodyTemplate.DefaultView[i]["Unit"].ToString();
                                }

                                if (strDataSourceText.Contains(lstCurrentVariable[0] + strCurrentUnit))
                                {
                                    ////////////////rtxtTemplate.SelectionStart = Convert.ToInt16(strDataSourceText.IndexOf(lstCurrentVariable[0].Trim(' '), 0)) +
                                    ////////////////                              iTotalIndexLength + (rtxtTemplate.Lines[iLines].Length - strDataSourceText.Length);
                                    ////////////////rtxtTemplate.SelectionLength = lstCurrentVariable[0].TrimStart(' ').Length + strCurrentUnit.TrimEnd(' ').Length;

                                    ////////////////rtxtTemplate.SelectionColor = System.Drawing.Color.FromArgb(Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]));
                                    ///

                                    //Static Indent Space.
                                    rtxtTemplate.SelectionStart = Convert.ToInt16(rtxtTemplate.Lines[iLines].IndexOf(strDataSourceText, 0)) +
                                                                    iTotalIndexLength;

                                    rtxtTemplate.SelectionLength = strDataSourceText.Length -
                                                                     lstCurrentVariable[0].Length - strCurrentUnit.Length;

                                    rtxtTemplate.SelectionFont = new System.Drawing.Font(dtBodyTemplate.DefaultView[i]["LineHeaderFontName"].ToString(),
                                                                                           Convert.ToInt16(dtBodyTemplate.DefaultView[i]["LineHeaderFontSize"]));

                                    if (!Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceUnderlined"]))
                                    {
                                        //Data Source and Unit.
                                        iStart = Convert.ToInt16(strDataSourceText.IndexOf(lstCurrentVariable[0].Trim(' '), 0)) +
                                                                (rtxtTemplate.Lines[iLines].Length - strDataSourceText.Length);

                                        iLength = lstCurrentVariable[0].TrimStart(' ').Length + strCurrentUnit.TrimEnd(' ').Length;

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
                                        iStart = Convert.ToInt16(strDataSourceText.IndexOf(lstCurrentVariable[0].Trim(' '), 0)) +
                                                                        (rtxtTemplate.Lines[iLines].Length - strDataSourceText.Length);

                                        iLength = lstCurrentVariable[0].TrimStart(' ').Length;

                                        AddFontToRichText(iStart + iTotalIndexLength, iLength,
                                                            dtBodyTemplate.DefaultView[i]["LineHeaderFontName"].ToString(),
                                                            Convert.ToInt16(dtBodyTemplate.DefaultView[i]["LineHeaderFontSize"]),
                                                            Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]),
                                                            Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceUnderlined"]),
                                                            Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceBold"]),
                                                            Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceItalic"]));

                                        //Unit. 
                                        iStart = Convert.ToInt16(strDataSourceText.IndexOf(strCurrentUnit.TrimEnd(' '), 0)) +
                                                                        (rtxtTemplate.Lines[iLines].Length - strDataSourceText.Length);

                                        iLength = strCurrentUnit.TrimEnd(' ').Length;

                                        AddFontToRichText(iStart + iTotalIndexLength, iLength,
                                                            dtBodyTemplate.DefaultView[i]["LineHeaderFontName"].ToString(),
                                                            Convert.ToInt16(dtBodyTemplate.DefaultView[i]["LineHeaderFontSize"]),
                                                            Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]),
                                                            false,
                                                            Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceBold"]),
                                                            Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceItalic"]));
                                    }

                                }
                                ////////////////}
                            }
                        }
                    }


                    if ((dtBodyTemplate.DefaultView[i]["Spacing"].ToString() != "Static Indent") &&
                        (!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Replace(" ", ""))) &&
                        (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Beyond_Bit_Header"))
                    {
                        strDataSourceText = rtxtTemplate.Lines[iLines].Substring(iStart + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Length);

                        lstCurrentVariable.Clear();
                        lstCurrentVariable = GetCurrentVariable(dtBodyTemplate.DefaultView[i]["DataSource"].ToString(), false, iBeyondCount);
                        if (lstCurrentVariable.Count == 2)
                        {
                            strCurrentUnit = dtBodyTemplate.DefaultView[i]["Unit"].ToString() + lstCurrentVariable[1];
                        }
                        else
                        {
                            strCurrentUnit = dtBodyTemplate.DefaultView[i]["Unit"].ToString();
                        }

                        if (strDataSourceText.Contains(lstCurrentVariable[0] + strCurrentUnit))
                        {
                            if (dtBodyTemplate.DefaultView[i]["DataSource"].ToString() != "Comments")
                            {
                                iStart = Convert.ToInt16(strDataSourceText.IndexOf(lstCurrentVariable[0].Trim(' '), 0));

                                iLength = lstCurrentVariable[0].TrimStart(' ').Length + strCurrentUnit.TrimEnd(' ').Length;

                                AddFontToRichText(iStart + iTotalIndexLength + (rtxtTemplate.Lines[iLines].Length - strDataSourceText.Length), iLength,
                                                  dtBodyTemplate.DefaultView[i]["DataSourceFontName"].ToString(),
                                                  Convert.ToInt16(dtBodyTemplate.DefaultView[i]["DataSourceFontSize"]),
                                                  Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]),
                                                  Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceUnderlined"]),
                                                  Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceBold"]),
                                                  Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceItalic"]));
                            }
                            else
                            {
                                if (chkComments.Checked)
                                {
                                    int iTotalIndexLengthTemp = iTotalIndexLength;
                                    for (int j = 0; j < rtxtComments.Lines.Length; j++)
                                    {
                                        strDataSourceText = rtxtTemplate.Lines[iLines].Substring(iStart + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Length);

                                        //if (!string.IsNullOrEmpty(dtBodyTemplate.DefaultView[i]["LineHeader"].ToString().Replace(" ", "")))
                                        //{
                                            iStart = Convert.ToInt16(strDataSourceText.IndexOf(lstCurrentVariable[0].Trim(' '), 0));
                                        //}
                                        //else
                                        //{

                                        //}

                                        iLength = rtxtComments.Lines[j].TrimEnd(' ').Length;

                                        AddFontToRichText(iStart + iTotalIndexLengthTemp + (rtxtTemplate.Lines[iLines].Length - strDataSourceText.Length), iLength + Convert.ToInt16(dtBodyTemplate.DefaultView[i]["Spacing"]),
                                                      dtBodyTemplate.DefaultView[i]["DataSourceFontName"].ToString(),
                                                      Convert.ToInt16(dtBodyTemplate.DefaultView[i]["DataSourceFontSize"]),
                                                      Convert.ToInt32(dtBodyTemplate.DefaultView[i]["DataSourceFontColor"]),
                                                      Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceUnderlined"]),
                                                      Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceBold"]),
                                                      Convert.ToBoolean(dtBodyTemplate.DefaultView[i]["DataSourceItalic"]));

                                        if (j < (rtxtComments.Lines.Length - 1))
                                        {
                                            iTotalIndexLengthTemp = iTotalIndexLengthTemp + rtxtTemplate.Lines[iLines].Length + 1;// + Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]);
                                            iLines = iLines + 1;
                                        }
                                        else
                                        {
                                            iTotalIndexLength = iTotalIndexLengthTemp;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //if (i == 0)
                    //{
                    //    iTotalStringLength = -1;
                    //}  

                    //iTotalStringLength = iTotalStringLength + rtxtTemplate.Lines[iLines].Length + ((Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]) + 1) * 2);
                    iTotalIndexLength = iTotalIndexLength + rtxtTemplate.Lines[iLines].Length + 1 + Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]);

                    iLines = iLines + 1 + Convert.ToInt16(dtBodyTemplate.DefaultView[i]["BlankLines"]);
                }

                if (iBeyondCount > 0)
                {
                    if (i < (dtBodyTemplate.DefaultView.Count - 1))
                    {
                        if ((dtBodyTemplate.DefaultView[i]["DataSource"].ToString().Contains("Beyond")) &&
                            ((!dtBodyTemplate.DefaultView[i + 1]["DataSource"].ToString().Contains("Beyond")) || (i == (dtBodyTemplate.DefaultView.Count - 1))))
                        {
                            if ((iSVNumOfProj - iBeyondCount) > 2)
                            {
                                i = iFirstBeyondRow;
                            }
                        }
                    }
                }
            }
            
            if (!chkWithoutAC.Checked)
            {
                int iACTextStart = 0;
                int iACTextLength = 0;
                for (int iACLines = iLines; iACLines < rtxtTemplate.Lines.Length; iACLines++)
                {
                    iACTextStart = rtxtTemplate.GetFirstCharIndexFromLine(iACLines);

                    iACTextLength = rtxtTemplate.Lines[iACLines].Length;

                    if (iACTextLength > 0)
                    {
                        if (rtxtTemplate.Lines[iACLines][0].ToString() == "*")
                        {
                            AddFontToRichText(iACTextStart, iACTextLength,
                                                  dtBodyTemplate.DefaultView[dtBodyTemplate.DefaultView.Count - 1]["LineHeaderFontName"].ToString(),                                                  
                                                  Convert.ToInt16(dtBodyTemplate.DefaultView[dtBodyTemplate.DefaultView.Count - 1]["LineHeaderFontSize"]),
                                                  Color.Black.ToArgb(), true, true,
                                                  Convert.ToBoolean(dtBodyTemplate.DefaultView[dtBodyTemplate.DefaultView.Count - 1]["LineHeaderItalic"]));
                        }
                        else
                        {
                            AddFontToRichText(iACTextStart, iACTextLength,
                                                      dtBodyTemplate.DefaultView[dtBodyTemplate.DefaultView.Count - 1]["LineHeaderFontName"].ToString(),
                                                      Convert.ToInt16(dtBodyTemplate.DefaultView[dtBodyTemplate.DefaultView.Count - 1]["LineHeaderFontSize"]),
                                                      Convert.ToInt32(dtBodyTemplate.DefaultView[dtBodyTemplate.DefaultView.Count - 1]["LineHeaderFontColor"]),
                                                      Convert.ToBoolean(dtBodyTemplate.DefaultView[dtBodyTemplate.DefaultView.Count - 1]["LineHeaderUnderlined"]),
                                                      Convert.ToBoolean(dtBodyTemplate.DefaultView[dtBodyTemplate.DefaultView.Count - 1]["LineHeaderBold"]),
                                                      Convert.ToBoolean(dtBodyTemplate.DefaultView[dtBodyTemplate.DefaultView.Count - 1]["LineHeaderItalic"]));
                        }
                    }
                }
            }
        }

        //Adds Font to the Selected Text
        private void AddFontToRichText(int iStart, int iLength, string strFontName, int iFontSize, 
                                       int iFontColor, bool bUnderlined, bool bBold, bool bItalic)
        {
            rtxtTemplate.SelectionStart = iStart;
            rtxtTemplate.SelectionLength = iLength;

            rtxtTemplate.SelectionColor = System.Drawing.Color.FromArgb(iFontColor);

            if (!bUnderlined && !bBold && !bItalic)
            {
                rtxtTemplate.SelectionFont = new System.Drawing.Font(strFontName, iFontSize);
            }
            else if (bUnderlined && !bBold && !bItalic)
            {
                rtxtTemplate.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Underline);
            }
            else if (bUnderlined && bBold && !bItalic)
            {
                rtxtTemplate.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Underline | FontStyle.Bold);
            }
            else if (bUnderlined && !bBold && bItalic)
            {
                rtxtTemplate.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Underline | FontStyle.Italic);
            }
            else if (bUnderlined && bBold && bItalic)
            {
                rtxtTemplate.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Underline | FontStyle.Bold | FontStyle.Italic);
            }
            else if (!bUnderlined && bBold && !bItalic)
            {
                rtxtTemplate.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Bold);
            }
            else if (!bUnderlined && bBold && bItalic)
            {
                rtxtTemplate.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Bold | FontStyle.Italic);
            }
            else if (!bUnderlined && !bBold && bItalic)
            {
                rtxtTemplate.SelectionFont = new System.Drawing.Font(strFontName, iFontSize, FontStyle.Italic);
            }
        }

        //Generates needed spaces
        private string SpaceGenerator(string strText, int iStaticIndent)
        {
            string strSpaces = "";

            if ((iStaticIndent - strText.Length) == 0)            
            {
                return "";
            }

            for (int i = 0; i < iStaticIndent - strText.Length; i++)
            {
                strSpaces = strSpaces + " ";
            }

            return strSpaces;
        }

        private string HTMLSpaceGenerator(int iSpaceCount)
        {
            string strSpaces = "";

            if (iSpaceCount == 0)
            { 
                return "";
            }

            for (int i = 0; i < iSpaceCount; i++)
            {
                strSpaces = strSpaces + "&nbsp;";
            }

            return strSpaces;
        }

        //Generates needed spaces for Static Indent.
        private string HTMLSpaceGenerator(string strText, int iStaticIndent)
        {
            string strSpaces = "";

            if (iStaticIndent == 0)
            {                
                return "";
            }

            for (int i = 0; i < iStaticIndent - strText.Length; i++)
            {
                strSpaces = strSpaces + "&nbsp;";
            }

            return strSpaces;
        }

        //Get the column number for the desired value.  
        private int GetColNum (string strToLookFor, string strLine)
        {
            int i = 0;
            int intColNum = 0;
            try
            {
                for (int j = 0; j < 25; j++)
                {
                    if (strLine.Substring(0, strToLookFor.Length).ToString() == strToLookFor)
                    {
                        intColNum = j;
                        j = 25;
                    }
                    else if (intColNum > 0)
                    {
                        j = 25;
                    }

                    i = strLine.IndexOf("\t");
                    strLine = strLine.Substring(i + 1);

                }
                if (intColNum == 0)
                {
                    MessageBox.Show("Please have the DD add " + strToLookFor + " to the either the Survey or AC file and then try again.");
                }
                return intColNum;
            }
            catch
            {
                MessageBox.Show("Please have the DD add " + strToLookFor + " to the either the Survey or AC file and then try again.");
                return 0;
            }
        }

        //Get the value from the line based on it's column number.  
        private string GetStringFromLine(int intPos, string strLine, string strSearchFor)
        {
            
            int i;
            for (int j = intPos; j > 0; j--)
                {
                    i = strLine.IndexOf(strSearchFor);
                    strLine = strLine.Substring(i + 1);
                }
            i = strLine.IndexOf(strSearchFor);
            if (i < 1)
            {
                i = strLine.Length;
            }
            return strLine.Substring(0, i);
        }

        //Determines the Above/Below for Lateral Target Line Changes based on VS, TVD, and Dip Angle.

        private string GetAboveBelowLateralTargetLine(double dBitTVD, double dBitVS)
        {
            double dDipAngle;
            if (Convert.ToDouble(txtTargetLineInc.Text) < 90)
            {
                dDipAngle = 90 - Convert.ToDouble(txtTargetLineInc.Text);
            }
            else if (Convert.ToDouble(txtTargetLineInc.Text) > 90)
            {
                dDipAngle = Convert.ToDouble(txtTargetLineInc.Text) - 90;
            }
            else
            {
                dDipAngle = 0;
            }

            double dChangeInTVD = Math.Tan(dDipAngle / 180 * Math.PI) * (dBitVS - Convert.ToDouble(txtTargetLineBeginningVS.Text));

            double dTargetLineTVD;
            if (Convert.ToDouble(txtTargetLineInc.Text) < 90)
            {
                dTargetLineTVD = Convert.ToDouble(txtTargetLineBeginningTVD.Text) + Math.Abs(dChangeInTVD);
            }
            else if (Convert.ToDouble(txtTargetLineInc.Text) > 90)
            {
                dTargetLineTVD = Convert.ToDouble(txtTargetLineBeginningTVD.Text) - Math.Abs(dChangeInTVD); 
            }
            else
            {
                dTargetLineTVD = Convert.ToDouble(txtTargetLineBeginningTVD.Text);
            }

            string strAboveBelow;

            if (dTargetLineTVD < dBitTVD)
            {
                strAboveBelow = (dBitTVD - dTargetLineTVD).ToString("0.00") + " Below";
            }
            else if (dTargetLineTVD > dBitTVD)
            {
                strAboveBelow = (dTargetLineTVD - dBitTVD).ToString("0.00") + " Above";
            }
            else
            {
                strAboveBelow = "0.00";
            }

            return strAboveBelow;
        }

        private bool GetSurveyInfo()
        {
            lblProgress.Text = "Getting Survey Info";
            lblProgress.Refresh();

            string survline = "";
            string templine = "";
            string[] strarrSurvLines;

            string strSurveyMD;
            int IncPos = 0;
            int AzmPos = 0;
            int TVDPos = 0;
            int NorthPos = 0;
            int EastPos = 0;
            int VSPos = 0;

            try
            {
                List<string> list = new List<string>();

                System.IO.StreamReader survfile = new System.IO.StreamReader(@"" + txtSurveyFileName.Text);

                while ((templine = survfile.ReadLine()) != null)
                {
                    if (templine.Contains("Survey Points:"))
                    {
                        templine = survfile.ReadLine();
                        IncPos = GetColNum("Inc", templine);
                        AzmPos = GetColNum("Az", templine);
                        TVDPos = GetColNum("TVD", templine);
                        NorthPos = GetColNum("N.Offset", templine);
                        EastPos = GetColNum("E.Offset", templine);
                        VSPos = GetColNum("VS", templine);
                    }
                    else if ((IncPos > 0) & (AzmPos > 0) & (TVDPos > 0) & (NorthPos > 0) & (EastPos > 0) & (VSPos > 0))
                    {
                        if (templine.Substring(0, 1) == "\t")
                        {
                            strarrSurvLines = list.ToArray();
                            if ((Convert.ToInt16(txtSVNumOfProj.Text) + 1) > strarrSurvLines.Length)
                            {
                                txtSVNumOfProj.Text = (strarrSurvLines.Length - 1).ToString();
                                MessageBox.Show("The 'Rows in DD Survey File' entry cannot exceed the actual number of Survey/Projection entries in the DD Survey Text File.  " +
                                                "Changing the number to " + txtSVNumOfProj.Text);
                            }

                            if (strarrSurvLines.Length > 2)
                            {
                                if (iSVNumOfProj < 2)
                                {
                                    iSVNumOfProj = 2;
                                }
                                survline = strarrSurvLines[strarrSurvLines.Length - (Convert.ToInt16(iSVNumOfProj))];

                                strSurveyMD = survline.Substring(0, survline.IndexOf("\t"));
                                if (strSurveyMD.Contains("(pr)"))
                                {
                                    strSurveyMD = strSurveyMD.Substring(0, strSurveyMD.Length - 5);
                                    strSurveyMD = (Math.Truncate(Convert.ToDouble(strSurveyMD))).ToString();
                                }

                                dSurveyMD = Convert.ToDecimal(strSurveyMD);
                                dInc = Convert.ToDecimal(GetStringFromLine(IncPos, survline, "\t"));
                                dAzm = Convert.ToDecimal(GetStringFromLine(AzmPos, survline, "\t"));
                                dTVD = Convert.ToDecimal(GetStringFromLine(TVDPos, survline, "\t"));
                                dNorth = Convert.ToDecimal(GetStringFromLine(NorthPos, survline, "\t"));
                                dEast = Convert.ToDecimal(GetStringFromLine(EastPos, survline, "\t"));
                                dVS = Convert.ToDecimal(GetStringFromLine(VSPos, survline, "\t"));
                                txtSVNumOfProj.Text = iSVNumOfProj.ToString();
                            }
                            else
                            {
                                MessageBox.Show("There is only one survey row in the 'DD Survey File'.  At least two rows are required.");
                                return false;
                            }
                            break;
                        }
                        list.Add(templine);
                    }
                }
                survfile.Close();
                survfile.Dispose();
                lblProgress.Text = "Survey Info Calculated";
                lblProgress.Refresh();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was an issue reading the DD Survey File.  " + ex.ToString());
                return false;
            }
                
        }

        private bool ReadSurveyFile()
        {
            try
            {
                string survline = "";
                string templine = "";
                string bitline = "";
                string beyondbitline = "";
                int IncPos = 0;
                int AzmPos = 0;
                int TVDPos = 0;
                int NorthPos = 0;
                int EastPos = 0;
                int DLSPos = 0;
                int VSPos = 0;
                int TFPos = 0;
                int AbovePos = 0;
                int RightPos = 0;

                string strSurveyDepth;
                string strSurveyTVD;
                string strSurveyNorth;
                string strSurveyEast;
                string strSurveyVS;

                string strBitDepth;
                string strBitInc;
                string strBitTVD;
                string strBitNorth;
                string strBitEast;
                string strBitVS;
                string strBitTF;
                string strAboveBelow;
                string strRightLeft;
                string strBeyondBitDepth;
                string strBeyondBitInc;
                string strBeyondBitTVD;
                string strBeyondBitNorth;
                string strBeyondBitEast;
                string strBeyondBitVS;
                string strBeyondBitTF;
                string strBeyondBitAboveBelow;
                string strBeyondBitRightLeft;
                string[] strarrSurvLines;
                  
            
                List<string> list = new List<string>();

                //System.IO.StreamReader survfile = new System.IO.StreamReader("c:\\" + txtSurveyFileName.Text.ToString());
                using (System.IO.StreamReader survfile = new System.IO.StreamReader(@"" + txtSurveyFileName.Text))
                {
                    while ((templine = survfile.ReadLine()) != null)
                    {
                        if (templine.Contains("Survey Points:"))
                        {
                            templine = survfile.ReadLine();
                            IncPos = GetColNum("Inc", templine);
                            AzmPos = GetColNum("Az", templine);
                            TVDPos = GetColNum("TVD", templine);
                            NorthPos = GetColNum("N.Offset", templine);
                            EastPos = GetColNum("E.Offset", templine);
                            DLSPos = GetColNum("DLS", templine);
                            VSPos = GetColNum("VS", templine);
                            TFPos = GetColNum("T.Face", templine);
                            AbovePos = GetColNum("High to Plan", templine);
                            RightPos = GetColNum("Right to Plan", templine);
                        }
                        else if ((IncPos > 0) & (AzmPos > 0) & (TVDPos > 0) & (NorthPos > 0) & (EastPos > 0) & (DLSPos > 0) & (VSPos > 0) & (AbovePos > 0) & (RightPos > 0))
                        {
                            if (templine.Substring(0, 1) == "\t")
                            {
                                strarrSurvLines = list.ToArray();

                                if (strarrSurvLines.Length > 2)
                                {
                                    try
                                    {
                                        if (Convert.ToDouble(txtSVNumOfProj.Text) < 2)
                                        {
                                            txtSVNumOfProj.Text = "2";
                                        }
                                        else if (String.IsNullOrEmpty(txtSVNumOfProj.Text))
                                        {
                                            txtSVNumOfProj.Text = "2";
                                        }

                                        if ((Convert.ToInt16(txtSVNumOfProj.Text) + 1) > strarrSurvLines.Length)
                                        {
                                            txtSVNumOfProj.Text = (strarrSurvLines.Length - 1).ToString();
                                            MessageBox.Show("The 'Rows in DD Survey File cannot exceed the actual number of Survey/Projection entries in the DD Survey Text File.  " +
                                                            "Changing the number to " + txtSVNumOfProj.Text);
                                        }

                                        survline = strarrSurvLines[strarrSurvLines.Length - (Convert.ToInt16(txtSVNumOfProj.Text))];
                                        bitline = strarrSurvLines[strarrSurvLines.Length - (Convert.ToInt16(txtSVNumOfProj.Text) - 1)];
                                        DataRow drRow = dtVariables.NewRow();


                                        strSurveyDepth = survline.Substring(0, survline.IndexOf("\t"));
                                        if (strSurveyDepth.Contains("(pr)"))
                                        {
                                            strSurveyDepth = strSurveyDepth.Substring(0, strSurveyDepth.Length - 5);
                                            strSurveyDepth = (Math.Truncate(Convert.ToDouble(strSurveyDepth))).ToString();
                                        }
                                        drRow["Name"] = "Survey_MD";
                                        drRow["Value"] = strSurveyDepth;
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Survey_Inc";
                                        drRow["Value"] = GetStringFromLine(IncPos, survline, "\t");
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Survey_Azm";
                                        drRow["Value"] = GetStringFromLine(AzmPos, survline, "\t");
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        strSurveyTVD = GetStringFromLine(TVDPos, survline, "\t");
                                        drRow["Name"] = "Survey_TVD";
                                        drRow["Value"] = GetStringFromLine(TVDPos, survline, "\t");
                                        dtVariables.Rows.Add(drRow);
                                        strSurveyNorth = GetStringFromLine(NorthPos, survline, "\t");
                                        strSurveyEast = GetStringFromLine(EastPos, survline, "\t");
                                        string strNorthOrSouth;
                                        string strEastOrWest;

                                        if (Convert.ToDouble(strSurveyNorth) >= 0)
                                        {
                                            strNorthOrSouth = " North";
                                        }
                                        else
                                        {
                                            strNorthOrSouth = " South";
                                        }
                                        if (Convert.ToDouble(strSurveyEast) >= 0)
                                        {
                                            strEastOrWest = " East";
                                        }
                                        else
                                        {
                                            strEastOrWest = " West";
                                        }
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Survey_North";
                                        drRow["Value"] = Math.Abs(Convert.ToDouble(strSurveyNorth)).ToString();
                                        drRow["Unit"] = strNorthOrSouth;
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Survey_East";
                                        drRow["Value"] = Math.Abs(Convert.ToDouble(strSurveyEast)).ToString();
                                        drRow["Unit"] = strEastOrWest;
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        strSurveyVS = GetStringFromLine(VSPos, survline, "\t");
                                        drRow["Name"] = "Survey_VS";
                                        drRow["Value"] = GetStringFromLine(VSPos, survline, "\t");
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Survey_Tool_Face";
                                        drRow["Value"] = GetStringFromLine(TFPos, survline, "\t");
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Survey_DLS";
                                        drRow["Value"] = GetStringFromLine(DLSPos, survline, "\t");
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Survey_A/B_Plan";
                                        drRow["Value"] = GetStringFromLine(AbovePos, survline, "\t");
                                        drRow["Unit"] = GetAboveBelowUnit(drRow["Value"].ToString());
                                        drRow["Value"] = Math.Abs(Convert.ToDecimal(drRow["Value"])).ToString();
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Survey_R/L_Plan";
                                        drRow["Value"] = GetStringFromLine(RightPos, survline, "\t");
                                        drRow["Unit"] = GetRightLeftUnit(drRow["Value"].ToString());
                                        drRow["Value"] = Math.Abs(Convert.ToDecimal(drRow["Value"])).ToString();
                                        dtVariables.Rows.Add(drRow);

                                        strBitDepth = bitline.Substring(0, bitline.IndexOf("\t"));
                                        if (strBitDepth.Contains("(pr)"))
                                        {
                                            strBitDepth = strBitDepth.Substring(0, strBitDepth.Length - 5);
                                            strBitDepth = (Math.Truncate(Convert.ToDouble(strBitDepth))).ToString();
                                        }
                                        drRow = dtVariables.NewRow();
                                        dBitMD = Convert.ToDecimal(strBitDepth);
                                        drRow["Name"] = "Bit_MD";
                                        drRow["Value"] = strBitDepth;
                                        dtVariables.Rows.Add(drRow);
                                        strBitInc = GetStringFromLine(IncPos, bitline, "\t");
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Bit_Inc";
                                        drRow["Value"] = strBitInc;
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Bit_Azm";
                                        drRow["Value"] = GetStringFromLine(AzmPos, bitline, "\t");
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        strBitTVD = GetStringFromLine(TVDPos, bitline, "\t");
                                        drRow["Name"] = "Bit_TVD";
                                        drRow["Value"] = strBitTVD;
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        strBitNorth = GetStringFromLine(NorthPos, bitline, "\t");
                                        drRow["Name"] = "Bit_North";
                                        drRow["Value"] = strBitNorth;
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        strBitEast = GetStringFromLine(EastPos, bitline, "\t");
                                        drRow["Name"] = "Bit_East";
                                        drRow["Value"] = strBitEast;
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        strBitVS = GetStringFromLine(VSPos, bitline, "\t");
                                        drRow["Name"] = "Bit_VS";
                                        drRow["Value"] = strBitVS;
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Bit_Tool_Face";
                                        drRow["Value"] = GetStringFromLine(TFPos, bitline, "\t");
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Bit_DLS";
                                        drRow["Value"] = GetStringFromLine(DLSPos, bitline, "\t"); 
                                        dtVariables.Rows.Add(drRow);
                                        //strBitAbove = GetStringFromLine(AbovePos, bitline, "\t");
                                        //strBitRight = GetStringFromLine(RightPos, bitline, "\t");
                                        strAboveBelow = GetStringFromLine(AbovePos, bitline, "\t");
                                        strRightLeft = GetStringFromLine(RightPos, bitline, "\t");

                                        bool bUseNSEWCoord = false;

                                        //Check Orientation for Above/Below, Right/Left. 
                                        if (chkChangePlanNSEWZeroReference.Checked)
                                        {
                                            DialogResult drToUseBoxValue = DialogResult.No;

                                            //Get Orientation based on plan North and East.                                         
                                            if ((Convert.ToDouble(strBitInc) < 5) && ((txtPlanNorthZeroReference.Text != "") && (txtPlanEastZeroReference.Text != "")))
                                            {
                                                drToUseBoxValue = MessageBox.Show(
                                                                    "To use the 'Above/Below' and 'Right/Left' based on plan north and east click 'Yes'.  " +
                                                                    "To use the 5D 'Above/Below' and 'Right/Left' click 'No'?",
                                                                    "Use Current 'Above/Below', 'Right/Left' Values?", MessageBoxButtons.YesNo);
                                            }
                                        
                                            //Get Orientation from 5D.  
                                            if (drToUseBoxValue == DialogResult.Yes)
                                            {
                                                if ((!string.IsNullOrEmpty(txtPlanEastZeroReference.Text.Replace(" ", ""))) &&
                                                    ((!string.IsNullOrEmpty(txtPlanNorthZeroReference.Text.Replace(" ", "")))))
                                                {
                                                    strAboveBelow = (Convert.ToDouble(strBitNorth) - Convert.ToDouble(txtPlanNorthZeroReference.Text)).ToString();
                                                    strRightLeft = (Convert.ToDouble(strBitEast) - Convert.ToDouble(txtPlanEastZeroReference.Text)).ToString();

                                                    bUseNSEWCoord = true;
                                                }
                                                else
                                                {
                                                    MessageBox.Show("'Plan North Zero Reference' and 'Plan East Zero Reference' need to both have values.");
                                                    return false;
                                                }
                                            }
                                        }

                                        string GetAboveBelowUnit(string strAB)
                                        {
                                            decimal dAB = Convert.ToDecimal(strAB);
                                            string strUnit = "";

                                            if (dAB > 0)
                                            {
                                                strUnit = " Below";
                                            }
                                            else
                                            {
                                                strUnit = " Above";
                                            }

                                            return strUnit;
                                        }

                                        string GetRightLeftUnit(string strRL)
                                        {
                                            decimal dRL = Convert.ToDecimal(strRL);
                                            string strUnit = "";

                                            if (dRL > 0)
                                            {
                                                strUnit = " Left";
                                            }
                                            else
                                            {
                                                strUnit = " Right";
                                            }

                                            return strUnit;
                                        }

                                        string strRLUnit;
                                        string strABUnit;

                                        if (!bUseNSEWCoord)
                                        {
                                            if (Convert.ToDouble(strAboveBelow) > 0)
                                            {
                                                strABUnit = " Below";
                                            }
                                            else if (Convert.ToDouble(strAboveBelow) < 0)
                                            {
                                                strABUnit = " Above";
                                            }
                                            else
                                            {
                                                strABUnit = "";
                                            }

                                            if (Convert.ToDouble(strRightLeft) > 0)
                                            {
                                                strRLUnit = " Left";
                                            }
                                            else if (Convert.ToDouble(strRightLeft) < 0)
                                            {
                                                strRLUnit = " Right";
                                            }
                                            else
                                            {
                                                strRLUnit = "";
                                            }
                                        }
                                        else
                                        {
                                            if (Convert.ToDouble(strAboveBelow) > 0)
                                            {
                                                strABUnit = " North";
                                            }
                                            else if (Convert.ToDouble(strAboveBelow) < 0)
                                            {
                                                strABUnit = " South";
                                            }
                                            else
                                            {
                                                strABUnit = "";
                                            }

                                            if (Convert.ToDouble(strRightLeft) > 0)
                                            {
                                                strRLUnit = " East";
                                            }
                                            else if (Convert.ToDouble(strRightLeft) < 0)
                                            {
                                                strRLUnit = " West";
                                            }
                                            else
                                            {
                                                strRLUnit = "";
                                            }
                                        }

                                        strAboveBelow = Math.Abs(Convert.ToDouble(strAboveBelow)).ToString("0.00");
                                        strRightLeft = Math.Abs(Convert.ToDouble(strRightLeft)).ToString("0.00");

                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Bit_A/B_Plan";
                                        drRow["Value"] = strAboveBelow;
                                        drRow["Unit"] = strABUnit;
                                        dtVariables.Rows.Add(drRow);
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Bit_R/L_Plan";
                                        drRow["Value"] = strRightLeft;
                                        drRow["Unit"] = strRLUnit;
                                        dtVariables.Rows.Add(drRow);

                                        //string strABTargetTemp;
                                        string strABTargetUnit = "";
                                        string strTargetTemp = "";
                                        if (chkManualOrientation.Checked)
                                        {
                                            strTargetTemp = txtBitAboveBelow.Text;
                                            if (Convert.ToDouble(strTargetTemp) > 0)
                                            {
                                                strABTargetUnit = " Below";
                                            }
                                            else if (Convert.ToDouble(strTargetTemp) < 0)
                                            {
                                                strABTargetUnit = " Above";
                                            }
                                        }
                                        else if (chkLateralTargetLine.Checked)
                                        {                                        
                                            if ((txtTargetLineBeginningVS.Text == "") || (txtTargetLineBeginningTVD.Text == "") || (txtTargetLineInc.Text == ""))
                                            {
                                                MessageBox.Show("All of the Lateral Target Change Line info must be filled out.");
                                                strTargetTemp = "N/A";
                                                break;
                                            }
                                            else
                                            {
                                                strTargetTemp = GetABTargetLine(strSurveyTVD, strSurveyVS)[0];
                                                strABTargetUnit = GetABTargetLine(strSurveyTVD, strSurveyVS)[1];

                                                drRow = dtVariables.NewRow();
                                                drRow["Name"] = "Survey_A/B_Target_Line";
                                                drRow["Value"] = strTargetTemp;
                                                drRow["Unit"] = strABTargetUnit;
                                                dtVariables.Rows.Add(drRow);

                                                strTargetTemp = GetABTargetLine(strBitTVD, strBitVS)[0];
                                                strABTargetUnit = GetABTargetLine(strBitTVD, strBitVS)[1];
                                            }                                        
                                        }
                                        else
                                        {
                                            //strTargetTemp = "N/A";
                                        }

                                        
                                        drRow = dtVariables.NewRow();
                                        drRow["Name"] = "Bit_A/B_Target_Line";
                                        drRow["Value"] = strTargetTemp;
                                        drRow["Unit"] = strABTargetUnit;
                                        dtVariables.Rows.Add(drRow);

                                        //Fill table for Projections Beyond the Bit. 
                                        int iHeaderCount = 1;
                                        for (int i = 2; i < iSVNumOfProj; i++ )
                                        {
                                            beyondbitline = strarrSurvLines[strarrSurvLines.Length - (Convert.ToInt16(txtSVNumOfProj.Text) - i)];

                                            strBeyondBitDepth = beyondbitline.Substring(0, beyondbitline.IndexOf("\t"));
                                            if (strBeyondBitDepth.Contains("(pr)"))
                                            {
                                                strBeyondBitDepth = strBeyondBitDepth.Substring(0, strBeyondBitDepth.Length - 5);
                                                strBeyondBitDepth = (Math.Truncate(Convert.ToDouble(strBeyondBitDepth))).ToString();
                                            }
                                            drRow = dtVariables.NewRow();
                                            drRow["Name"] = "Beyond_Bit_Header";
                                            drRow["Value"] = iHeaderCount.ToString();
                                            dtVariables.Rows.Add(drRow);

                                            iHeaderCount = iHeaderCount + 1;

                                            drRow = dtVariables.NewRow();
                                            drRow["Name"] = "Beyond_Bit_MD";
                                            drRow["Value"] = strBeyondBitDepth;
                                            dtVariables.Rows.Add(drRow);
                                            strBeyondBitInc = GetStringFromLine(IncPos, beyondbitline, "\t");
                                            drRow = dtVariables.NewRow();
                                            drRow["Name"] = "Beyond_Bit_Inc";
                                            drRow["Value"] = strBeyondBitInc;
                                            dtVariables.Rows.Add(drRow);
                                            drRow = dtVariables.NewRow();
                                            drRow["Name"] = "Beyond_Bit_Azm";
                                            drRow["Value"] = GetStringFromLine(AzmPos, beyondbitline, "\t");
                                            dtVariables.Rows.Add(drRow);
                                            drRow = dtVariables.NewRow();
                                            strBeyondBitTVD = GetStringFromLine(TVDPos, beyondbitline, "\t");
                                            drRow["Name"] = "Beyond_Bit_TVD";
                                            drRow["Value"] = strBeyondBitTVD;
                                            dtVariables.Rows.Add(drRow);
                                            drRow = dtVariables.NewRow();
                                            strBeyondBitNorth = GetStringFromLine(NorthPos, beyondbitline, "\t");
                                            drRow["Name"] = "Beyond_Bit_North";
                                            drRow["Value"] = strBeyondBitNorth;
                                            dtVariables.Rows.Add(drRow);
                                            drRow = dtVariables.NewRow();
                                            strBeyondBitEast = GetStringFromLine(EastPos, beyondbitline, "\t");
                                            drRow["Name"] = "Beyond_Bit_East";
                                            drRow["Value"] = strBeyondBitEast;
                                            dtVariables.Rows.Add(drRow);
                                            drRow = dtVariables.NewRow();
                                            strBeyondBitVS = GetStringFromLine(VSPos, beyondbitline, "\t");
                                            drRow["Name"] = "Beyond_Bit_VS";
                                            drRow["Value"] = strBeyondBitVS;
                                            dtVariables.Rows.Add(drRow);
                                            drRow = dtVariables.NewRow();
                                            drRow["Name"] = "Beyond_Bit_Tool_Face";
                                            drRow["Value"] = GetStringFromLine(TFPos, beyondbitline, "\t");
                                            dtVariables.Rows.Add(drRow);
                                            drRow = dtVariables.NewRow();
                                            drRow["Name"] = "Beyond_Bit_DLS";
                                            drRow["Value"] = GetStringFromLine(DLSPos, beyondbitline, "\t"); ;
                                            dtVariables.Rows.Add(drRow);
                                            //strBeyondBitAbove = GetStringFromLine(AbovePos, beyondbitline, "\t");
                                            //strBeyondBitRight = GetStringFromLine(RightPos, beyondbitline, "\t");
                                            strAboveBelow = GetStringFromLine(AbovePos, beyondbitline, "\t");
                                            strRightLeft = GetStringFromLine(RightPos, beyondbitline, "\t");

                                            drRow = dtVariables.NewRow();
                                            drRow["Name"] = "Beyond_Bit_A/B_Plan";
                                            drRow["Value"] = GetStringFromLine(AbovePos, beyondbitline, "\t");
                                            drRow["Unit"] = GetAboveBelowUnit(drRow["Value"].ToString());
                                            drRow["Value"] = Math.Abs(Convert.ToDecimal(drRow["Value"])).ToString();
                                            dtVariables.Rows.Add(drRow);
                                            drRow = dtVariables.NewRow();
                                            drRow["Name"] = "Beyond_Bit_R/L_Plan";
                                            drRow["Value"] = GetStringFromLine(RightPos, beyondbitline, "\t");
                                            drRow["Unit"] = GetRightLeftUnit(drRow["Value"].ToString());
                                            drRow["Value"] = Math.Abs(Convert.ToDecimal(drRow["Value"])).ToString();
                                            dtVariables.Rows.Add(drRow);

                                            if (chkLateralTargetLine.Checked)
                                            {
                                                strTargetTemp = GetABTargetLine(strBeyondBitTVD, strBeyondBitVS)[0];
                                                strABTargetUnit = GetABTargetLine(strBeyondBitTVD, strBeyondBitVS)[1];

                                                drRow = dtVariables.NewRow();
                                                drRow["Name"] = "Beyond_Bit_A/B_Target_Line";
                                                drRow["Value"] = strTargetTemp;
                                                drRow["Unit"] = strABTargetUnit;
                                                dtVariables.Rows.Add(drRow);
                                            }
                                        }

                                        survfile.Close();
                                        survfile.Dispose();
                                        return true;
                                    }
                                    catch (Exception ex)
                                    {
                                        survfile.Close();
                                        survfile.Dispose();
                                        MessageBox.Show(ex + "  There was a problem reading the 'DD Survey File'.");
                                        return false;
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("There is only one row in the DD Survey Text File.  At least two rows are required.");
                                    survfile.Close();
                                    survfile.Dispose();
                                    return false;
                                }
                            }

                            list.Add(templine);
                        }
                    }
                    survfile.Close();
                    survfile.Dispose();
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString() + Environment.NewLine +
                                "There was an issue reading the DD Survey Text File.");
            }
            return false;
        }

        private List<string> GetABTargetLine(string strTVD, string strVS)
        {
            string strABTargetTemp = GetAboveBelowLateralTargetLine(Convert.ToDouble(strTVD), Convert.ToDouble(strVS));
            string strTargetTemp = "";
            string strABTargetUnit = "";

            if (strABTargetTemp.Contains(" Above"))
            {
                strTargetTemp = strABTargetTemp.Remove(strABTargetTemp.IndexOf(" Above"));
                strABTargetUnit = " Above";
            }
            else if (strABTargetTemp.Contains(" Below"))
            {
                strTargetTemp = strABTargetTemp.Remove(strABTargetTemp.IndexOf(" Below"));
                strABTargetUnit = " Below";
            }

            List<string> lstTarget = new List<string>()
            {
                strTargetTemp,
                strABTargetUnit
            };

            return lstTarget;

        }

        private int ColorText(int intStartPos)
        {
            int intStart = rtxtTemplate.Text.IndexOf("Possible Collision", intStartPos, StringComparison.CurrentCultureIgnoreCase) - 5;
            if (intStart > 0)
            {               
                    rtxtTemplate.SelectionStart = intStart;
                    rtxtTemplate.SelectionLength = 23;
                    rtxtTemplate.SelectionColor = Color.Red;
                    rtxtTemplate.SelectionLength = 0;
            }

            return intStart + 23;
        }

        //Underline and Bolden Text
        private void UnderlineText(string heading, string strBoldOrUnderline)
        {            
            int start = rtxtTemplate.Text.IndexOf(heading, StringComparison.CurrentCultureIgnoreCase);
            if (start > (-1))
            {
                rtxtTemplate.SelectionStart = start;
                rtxtTemplate.SelectionLength = heading.Length;
                //rtxtTemplate.SelectionFont = new Font(rtxtTemplate.SelectionFont, FontStyle.Underline);
                if ((strBoldOrUnderline == "Underline") || (strBoldOrUnderline == "Both"))
                {
                    rtxtTemplate.SelectionFont = new System.Drawing.Font("Courier New", 11, FontStyle.Underline);
                }
                if ((strBoldOrUnderline == "Bold") || (strBoldOrUnderline == "Both"))
                {
                    rtxtTemplate.SelectionFont = new System.Drawing.Font(rtxtTemplate.SelectionFont, rtxtTemplate.SelectionFont.Style | FontStyle.Bold);
                }
                rtxtTemplate.SelectionLength = 0;
            }

        }

        //Get distance between survey projections and converts it to a string.  
        private string GetSurveyDistance (string strPrevMD, string strCurrentMD)
        {
            double dMD = Convert.ToDouble(strCurrentMD) - Convert.ToDouble(strPrevMD);

            return dMD.ToString("0");
        }

        //Generates needed spaces
        private string SpaceGenerator(string strText)
        {
            string strSpaces = " ";
            for (int i = 0; i < 38 - strText.Length; i++)
            {
                strSpaces = strSpaces + " ";
            }

            return strSpaces;
        }

        //Check SF to see if it is above SF Threshold.  
        private string QCSF(string strSF, string strPlanSurvey, bool bIsSecondary)
        {
            if ((strPlanSurvey == "s") || ((strPlanSurvey == "p") && (bIsSecondary)))
            {
                if (Convert.ToDouble(strSF) <= Convert.ToDouble(txtACThreshold.Text))
                {
                    if (intSFDangerCount == 0)
                    {
                        MessageBox.Show("There are SF @ Values that are " + txtACThreshold.Text + " or lower.  Alert the DD to a possible collision.");
                    }

                    strSF = Convert.ToDouble(strSF).ToString("0.00") + " Possible Collision";
                    intSFDangerCount = intSFDangerCount + 1;
                }
                else
                {
                    strSF = Convert.ToDouble(strSF).ToString("0.00");
                }
            }
            return strSF;
        }

        //Check SF to see it is above Threshold and to determine text color. 
        private string QCSFColor(string strQCSF)
        {
            string strColor;

            if (!strQCSF.Contains("Possible Collision"))
            {
                strColor = GetRGBString(Color.Black.ToArgb());
            }
            else
            {
                strColor = GetRGBString(Color.Red.ToArgb());
            }

            return strColor;
        }

        private void ReadACFile()
        {
            try
            {
                int tempcounter = 0;
                int i = 0;
                int CCpos = 0;
                int SFpos = 0;
                int intNumOfProj = Convert.ToInt16(txtSVNumOfProj.Text);
                string strPrimWellName = "";
                string line;
                string templine;
                string strBitDepth;
                string strTemp;
                string strPlanOrSurvey;
                string strACHeader;
                bool bIsSecondary = false;
                string[] arrMD = new string[intNumOfProj];
                string[] arrCC = new string[intNumOfProj];
                string[] arrSF = new string[intNumOfProj];

                string strLineHeaderFontFamily;
                string strLineHeaderFontSize;
                string strTableRowBackColor;

                strLineHeaderFontFamily = dtBodyTemplate.DefaultView[dtBodyTemplate.DefaultView.Count - 1]["LineHeaderFontName"].ToString();
                strLineHeaderFontSize = dtBodyTemplate.DefaultView[dtBodyTemplate.DefaultView.Count - 1]["LineHeaderFontSize"].ToString() + "pt";
                strTableRowBackColor = GetRGBString(Convert.ToInt32(dtBodyTemplate.DefaultView[dtBodyTemplate.DefaultView.Count - 1]["TableRowBackColor"]));

                if ((dSurveyMD == 0) || (!chkBottomLineChecker.Checked))
                {
                    for (int j = 0; j < dtVariables.Rows.Count; j++)
                    {
                        if (dtVariables.Rows[i]["Name"].ToString() == "Survey_MD")
                        {
                            dSurveyMD = Convert.ToDecimal(dtVariables.Rows[i]["Value"]);
                            break;
                        }
                    }
                }

                using (System.IO.StreamReader file = new System.IO.StreamReader(@"" + txtACFileName.Text))
                {
                    while ((line = file.ReadLine()) != null)
                    {
                        int iStaticIndent = 0;        
                        for (int j = dtBodyTemplate.DefaultView.Count - 1; j >= 0; j--)
                        {
                            if (Convert.ToInt16(dtBodyTemplate.DefaultView[j]["StaticIndent"]) > 0)
                            {
                                iStaticIndent = Convert.ToInt16(dtBodyTemplate.DefaultView[j]["StaticIndent"]);
                                break;
                            }
                        }

                        string strWellNameAddOn = "";
                        if (line.Contains("Primary Well:"))
                        {
                            bIsSecondary = false;
                            if (line.Contains("(s)(TVD"))
                            {
                                strWellNameAddOn = " (Surveys)";
                                i = line.IndexOf("(s)(TVD");
                            }
                            else if (line.Contains("(p)(TVD"))
                            {
                                strWellNameAddOn = " (Plan)";
                                i = line.IndexOf("(p)(TVD");
                            }
                            else
                            {
                                i = line.IndexOf("(TVD");
                            }
                            //strPrimWellName = line.Substring(0, i - 1) + strWellNameAddOn;
                            strPrimWellName = line.Substring("Primary Well: ".Length, i - 1 - "Primary Well: ".Length);
                        }
                        else if (line.Contains("Secondary Well:"))
                        {
                            line = line.Substring(16);
                            if (line.Contains("(s)(TVD"))
                            {
                                strWellNameAddOn = " (Surveys)";
                                i = line.IndexOf("(s)(TVD");
                            }
                            else if (line.Contains("(p)(TVD"))
                            {
                                strWellNameAddOn = " (Plan)";
                                i = line.IndexOf("(p)(TVD");
                            }
                            else
                            {
                                i = line.IndexOf("(TVD");
                            }

                            
                            strACHeader = "**Anti-Collision for " + line.Substring(0, i - 1) + strWellNameAddOn + "**";

                            if (strPrimWellName != line.Substring(0, i - 1).Trim(' '))
                            {
                                bIsSecondary = true;
                            }
                            else
                            {
                                bIsSecondary = false;
                            }

                            lstACHeaders.Add(strACHeader);

                            if (chkGenerateTable.Checked)
                            {
                                builder.Append("<br>");
                                builder.Append("<style>table { width: 500px; border-collapse: collapse; border: 5px solid black; } th, td { border-collapse: collapse; border: 1px solid black; } </style>");
                                //builder.Append("<style type='text/css'>.headerStyle { column-span:all; width:500px; text-align:center; }</style>");
                                builder.Append("<style type='text/css'>.columnoneStyle { width:300px; text-align:center; }</style>");
                                builder.Append("<style type='text/css'>.columntwoStyle { width:200px; text-align:center; }</style>");
                                builder.Append("<table style ='background-color:" + strTableRowBackColor + "'>");
                                builder.Append("<tr><th colspan='2' style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + ";'>" + strACHeader + "</th></tr>");
                            }
                            else
                            {
                                builder.Append("<span style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + ";'>" + strACHeader + "</span><br>");
                            }

                            //rtxtTemplate.Text = rtxtTemplate.Text + strACHeader +
                            //                    Environment.NewLine + Environment.NewLine;

                            strPlanOrSurvey = line.Substring(i + 1, 1).ToString();

                            line = file.ReadLine();
                            CCpos = GetColNum("CC", line);
                            SFpos = GetColNum("SF", line);

                            tempcounter = 0;
                            while (tempcounter < 2)
                            {
                                line = file.ReadLine();
                                tempcounter++;
                            }
                            templine = line;
                            int iK = 0;
                            for (int k = 0; k < intNumOfProj; k++)
                            {
                                //Check for (pr) in MD string and remove if so.  
                                if (templine.Substring(0, templine.IndexOf("\t")).Contains("(pr)"))
                                {
                                    arrMD[iK] = templine.Substring(0, templine.Substring(0, templine.IndexOf("\t")).Length - 5);
                                }
                                else
                                {
                                    arrMD[iK] = templine.Substring(0, templine.IndexOf("\t"));
                                }
                                arrCC[iK] = Convert.ToDouble(GetStringFromLine(CCpos, templine, "\t")).ToString("0.00");

                                arrSF[iK] = Convert.ToDouble(GetStringFromLine(SFpos, templine, "\t")).ToString("0.00");
                                if (k < (intNumOfProj - 1))
                                {
                                    line = file.ReadLine();
                                    templine = line;
                                }

                                if (Convert.ToDecimal(arrMD[iK]) < dSurveyMD)
                                {
                                    k = -1;
                                }
                                else
                                {
                                    iK = iK + 1;
                                }
                            }

                            strBitDepth = arrMD[1];

                            //rtxtTemplate.Text = rtxtTemplate.Text + strACHeader +
                            //                    Environment.NewLine + Environment.NewLine;




                            //"<tr>" +
                            //            "<td class='columnoneStyle'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; color:" + strLineHeaderColor + "; " +
                            //            "font-weight:" + strLineHeaderBold + "; text-decoration:" + strLineHeaderUnderlined + "; " +
                            //            "font-style:" + strLineHeaderItalic + "; background-color:" + strTableRowBackColor + "'>" + dtBodyTemplate.DefaultView[i]["LineHeader"].ToString() + "</td>" +
                            //            "<td class='columntwoStyle'; style=' font-family:" + strDataSourceFontFamily + "; font-size:" + strDataSourceFontSize + "; color:" + strDataSourceColor + "; " +
                            //            "font-weight:" + strDataSourceBold + "; text-decoration:" + strDataSourceUnderlined + "; " +
                            //            "font-style:" + strDataSourceItalic + "; background-color:" + strTableRowBackColor + "'>" + lstCurrentVariable[0] + "</td>" +

                            string strQCSF;
                            string strQCSFColor;

                            //CTC @ Distances. 
                            if (chkGenerateTable.Checked)
                            {
                                builder.Append("<tr><td class='columnoneStyle'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + ";'>CTC @ Survey (ft):</td>");
                                builder.Append("<td class='columntwoStyle'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + ";'>" +
                                                Convert.ToDouble(arrCC[0]).ToString("0.00") + "</td></tr>");
                                builder.Append("<tr><td class='columnoneStyle'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + ";'>CTC @ PTB(ft):</td>");
                                builder.Append("<td class='columntwoStyle'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + ";'>" +
                                                Convert.ToDouble(arrCC[1]).ToString("0.00") + "</td></tr>");

                                ////rtxtTemplate.Text = rtxtTemplate.Text + "CTC @ Survey (ft):" + SpaceGenerator("CTC @ Survey (ft):" + Convert.ToDouble(arrCC[0]).ToString("0.00"), iStaticIndent) +
                                ////                    Convert.ToDouble(arrCC[0]).ToString("0.00") + "'" + Environment.NewLine +
                                ////                    "CTC @ PTB (ft):" + SpaceGenerator("CTC @ PTB (ft):" + Convert.ToDouble(arrCC[1]).ToString("0.00"), iStaticIndent) + Convert.ToDouble(arrCC[1]).ToString("0.00") + "'" + Environment.NewLine;

                                for (int k = 2; k < intNumOfProj; k++)
                                {
                                    strTemp = "CTC @ " + GetSurveyDistance(strBitDepth, arrMD[k]) + "(ft):";
                                    builder.Append("<tr><td class='columnoneStyle'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + ";'>" + strTemp + "</td>");
                                    builder.Append("<td class='columntwoStyle'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + ";'>" +
                                                    Convert.ToDouble(arrCC[k]).ToString("0.00") + "</td></tr>");
                                    //rtxtTemplate.Text = rtxtTemplate.Text + strTemp + SpaceGenerator(strTemp + Convert.ToDouble(arrCC[k]).ToString("0.00"), iStaticIndent) + Convert.ToDouble(arrCC[k]).ToString("0.00") + "'" + Environment.NewLine;
                                }

                                //rtxtTemplate.Text = rtxtTemplate.Text + Environment.NewLine;

                                //SF @s. 
                                strQCSF = QCSF(arrSF[0], strPlanOrSurvey, bIsSecondary);
                                strQCSFColor = QCSFColor(strQCSF);

                                builder.Append("<tr><td class='columnoneStyle'; Style'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + ";'>SF @ Survey (ft):</td>");
                                builder.Append("<td class='columntwoStyle'; Style'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; color:" + strQCSFColor + ";'>" +
                                                strQCSF + "</td></tr>");

                                strQCSF = QCSF(arrSF[1], strPlanOrSurvey, bIsSecondary);
                                strQCSFColor = QCSFColor(strQCSF);

                                builder.Append("<tr><td class='columnoneStyle'; Style'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + ";'>SF @ PTB (ft):</td>");
                                builder.Append("<td class='columntwoStyle'; Style'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; color:" + strQCSFColor + ";'>" +
                                                strQCSF + "</td></tr>");

                                //rtxtTemplate.Text = rtxtTemplate.Text + "SF @ Survey (ft):" + SpaceGenerator("SF @ Survey (ft):" + arrSF[0], iStaticIndent) + QCSF(arrSF[0], strPlanOrSurvey, bIsSecondary) + Environment.NewLine +
                                //                                        "SF @ PTB (ft):" + SpaceGenerator("SF @ PTB (ft):" + arrSF[1], iStaticIndent) + QCSF(arrSF[1], strPlanOrSurvey, bIsSecondary) + Environment.NewLine;

                                for (int k = 2; k < intNumOfProj; k++)
                                {
                                    strQCSF = QCSF(arrSF[k], strPlanOrSurvey, bIsSecondary);
                                    strQCSFColor = QCSFColor(strQCSF);

                                    strTemp = "SF @ " + GetSurveyDistance(strBitDepth, arrMD[k]) + "(ft):";
                                    builder.Append("<tr><td class='columnoneStyle'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + ";'>" + strTemp + "</td>");
                                    builder.Append("<td class='columntwoStyle'; Style'; style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; color:" + strQCSFColor + ";'>" +
                                                    strQCSF + "</td></tr>");
                                    //rtxtTemplate.Text = rtxtTemplate.Text + strTemp + SpaceGenerator(strTemp + arrSF[k], iStaticIndent) + QCSF(arrSF[k], strPlanOrSurvey, bIsSecondary) + Environment.NewLine;
                                }

                                builder.Append("</table>");
                            }
                            else
                            {
                                string strSurveySpacing = HTMLSpaceGenerator("CTC @ Survey (ft):" + Convert.ToDouble(arrCC[0]).ToString("0.00"), iStaticIndent);
                                string strBitSpacing = HTMLSpaceGenerator("CTC @ PTB (ft):" + Convert.ToDouble(arrCC[1]).ToString("0.00"), iStaticIndent);
                                string strSFSurveySpacing = HTMLSpaceGenerator("SF @ Survey (ft):" + arrSF[0], iStaticIndent);
                                string strSFBitSpacing = HTMLSpaceGenerator("SF @ PTB (ft):" + arrSF[1], iStaticIndent);

                                builder.Append("<span style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + ";'>CTC @ Survey (ft):" +
                                                strSurveySpacing + Convert.ToDouble(arrCC[0]).ToString("0.00") + "<br>" + "CTC @ PTB (ft):" + strBitSpacing + Convert.ToDouble(arrCC[1]).ToString("0.00") + "<br></td>");
                                ////rtxtTemplate.Text = rtxtTemplate.Text + "CTC @ Survey (ft):" + SpaceGenerator("CTC @ Survey (ft):" + Convert.ToDouble(arrCC[0]).ToString("0.00"), iStaticIndent) +
                                ////                    Convert.ToDouble(arrCC[0]).ToString("0.00") + "'" + Environment.NewLine +
                                ////                    "CTC @ PTB (ft):" + SpaceGenerator("CTC @ PTB (ft):" + Convert.ToDouble(arrCC[1]).ToString("0.00"), iStaticIndent) + Convert.ToDouble(arrCC[1]).ToString("0.00") + "'" + Environment.NewLine;

                                for (int k = 2; k < intNumOfProj; k++)
                                {
                                    strTemp = "CTC @ " + GetSurveyDistance(strBitDepth, arrMD[k]) + "(ft):";
                                    string strTempSpacing = HTMLSpaceGenerator(strTemp + Convert.ToDouble(arrCC[k]).ToString("0.00"), iStaticIndent);
                                    builder.Append("<span style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + ";'>" + strTemp +
                                                    strTempSpacing + Convert.ToDouble(arrCC[k]).ToString("0.00") + "<br></td>");
                                    //rtxtTemplate.Text = rtxtTemplate.Text + strTemp + SpaceGenerator(strTemp + Convert.ToDouble(arrCC[k]).ToString("0.00"), iStaticIndent) + Convert.ToDouble(arrCC[k]).ToString("0.00") + "'" + Environment.NewLine;
                                }

                                builder.Append("<br>");
                                //rtxtTemplate.Text = rtxtTemplate.Text + Environment.NewLine;

                                //SF @s. 
                                strQCSF = QCSF(arrSF[0], strPlanOrSurvey, bIsSecondary);
                                strQCSFColor = QCSFColor(strQCSF);

                                builder.Append("<span style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; color:" + strQCSFColor + ";'>SF @ Survey (ft):" + strSFSurveySpacing +
                                                strQCSF + "<br>" + "SF @ PTB (ft):" + strSFBitSpacing + QCSF(arrSF[1], strPlanOrSurvey, bIsSecondary) + "<br></span>");
                                //rtxtTemplate.Text = rtxtTemplate.Text + "SF @ Survey (ft):" + SpaceGenerator("SF @ Survey (ft):" + arrSF[0], iStaticIndent) + QCSF(arrSF[0], strPlanOrSurvey, bIsSecondary) + Environment.NewLine +
                                //                                        "SF @ PTB (ft):" + SpaceGenerator("SF @ PTB (ft):" + arrSF[1], iStaticIndent) + QCSF(arrSF[1], strPlanOrSurvey, bIsSecondary) + Environment.NewLine;

                                for (int k = 2; k < intNumOfProj; k++)
                                {
                                    strQCSF = QCSF(arrSF[k], strPlanOrSurvey, bIsSecondary);
                                    strQCSFColor = QCSFColor(strQCSF);
                                    strTemp = "SF @ " + GetSurveyDistance(strBitDepth, arrMD[k]) + "(ft):";
                                    string strTempSpacing = HTMLSpaceGenerator(strTemp + arrSF[k], iStaticIndent);
                                    builder.Append("<span style=' font-family:" + strLineHeaderFontFamily + "; font-size:" + strLineHeaderFontSize + "; color:" + strQCSFColor + ";'>" + strTemp +
                                                    strTempSpacing + strQCSF + "<br></span>");
                                    //rtxtTemplate.Text = rtxtTemplate.Text + strTemp + SpaceGenerator(strTemp + arrSF[k], iStaticIndent) + QCSF(arrSF[k], strPlanOrSurvey, bIsSecondary) + Environment.NewLine;
                                }

                                builder.Append("<br>");
                            }
                            //rtxtTemplate.Text = rtxtTemplate.Text + Environment.NewLine;
                        }
                    }

                    if (chkGenerateTable.Checked)
                    {
                        builder.Append("<br>");
                    }

                    file.Close();
                    file.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString() + Environment.NewLine +
                                "There was an issue reading the DD AC Text File.");
            }

        }       

        private void txtJobNumber_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.JobNumber = txtJobNumber.Text;
            Properties.Settings.Default.Save();
        }

        private void txtClient_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Client = txtClient.Text;
            Properties.Settings.Default.Save();
        }

        private void txtWellName_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.WellName = txtWellName.Text;
            Properties.Settings.Default.Save();
        }

        private void txtRig_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Rig = txtRig.Text;
            Properties.Settings.Default.Save();
        }

        private void txtRPM_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.RPM = txtSurfaceRPM.Text;
            Properties.Settings.Default.Save();
        }

        private void txtFlowRate_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.FlowRate = txtFlowRate.Text;
            Properties.Settings.Default.Save();
        }

        private void txtGamma_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Gamma = txtGamma.Text;
            Properties.Settings.Default.Save();
        }

        private void txtWOB_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.WOB = txtWOB.Text;
            Properties.Settings.Default.Save();
        }

        private void txtDiff_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Diff = txtDifferentialPressure.Text;
            Properties.Settings.Default.Save();
        }

        private void txtSPP_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SPP = txtSPP.Text;
            Properties.Settings.Default.Save();
        }

        private void txtACFileName_TextChanged(object sender, EventArgs e)
        {
            if (Path.GetFileName(txtACFileName.Text) == "")
            {
                txtACFileName.Text = @"c:\ACReport.txt";
            }
            Properties.Settings.Default.ACFileName = txtACFileName.Text;
            Properties.Settings.Default.Save();


            strACPath = txtACFileName.Text;
        }

        private void txtACThreshold_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ACThreshold = txtACThreshold.Text;
            Properties.Settings.Default.Save();
        }

        private void txtSurveyFileName_TextChanged(object sender, EventArgs e)
        {
            if (Path.GetFileName(txtSurveyFileName.Text) == "")
            {
                txtSurveyFileName.Text = @"c:\SurveyReport.txt";
            }
            Properties.Settings.Default.SurveyFileName = txtSurveyFileName.Text;
            Properties.Settings.Default.Save();

            strSurveyPath = txtSurveyFileName.Text;
        }

        private void txtTargetInc_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.TargetInc = txtTargetInc.Text;
            Properties.Settings.Default.Save();
        }

        private void txtTargetAzm_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.TargetAzm = txtTargetAzm.Text;
            Properties.Settings.Default.Save();
        }

        private void txtROP_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ROPRotating = txtROPRotating.Text;
            Properties.Settings.Default.Save();
        }

        private void txtPlanNorthZeroReference_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.PlanNorthZeroReference = txtPlanNorthZeroReference.Text;
            Properties.Settings.Default.Save();
        }

        private void txtPlanEastZeroReference_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.PlanEastZeroReference = txtPlanEastZeroReference.Text;
            Properties.Settings.Default.Save();
                
        }

        private void txtToolTemp_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ToolTemp = txtSurveyTemp.Text;
            Properties.Settings.Default.Save();
        }

        private void chkChangePlanNSEWZeroReference_CheckedChanged(object sender, EventArgs e)
        {
            if ((chkThirdParty.Checked) && (chkChangePlanNSEWZeroReference.Checked))
            {
                MessageBox.Show("'AC' cannot be added while in 'Third Party Mode'.");
                chkChangePlanNSEWZeroReference.Checked = false;
            }

            Properties.Settings.Default.UsePlanZeroReference = chkChangePlanNSEWZeroReference.Checked;
            Properties.Settings.Default.Save();

            ScreenSetup();
        }

        private void chkWithoutAC_CheckedChanged(object sender, EventArgs e)
        {
            if ((chkThirdParty.Checked) && (!chkWithoutAC.Checked))
            {
                MessageBox.Show("'AC' cannot be added while in 'Third Party Mode'.");
                chkWithoutAC.Checked = true;
            }

            Properties.Settings.Default.WithoutAC = chkWithoutAC.Checked;
            Properties.Settings.Default.Save();

            if (chkWithoutAC.Checked)
            {
                txtSVNumOfProj.Text = "2";
            }
            else
            {
                txtSVNumOfProj.Text = "4";
            }

            if (chkWithoutAC.Checked)
            {
                grpAC.Hide();
            }
            else
            {
                grpAC.Show();
            }            
        }

        private void chkThirdParty_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ThirdParty = chkThirdParty.Checked;
            Properties.Settings.Default.Save();

            chkBottomLineChecker.Checked = false;
            chkManualOrientation.Checked = false;
            chkWithoutAC.Checked = true;
            chkChangePlanNSEWZeroReference.Checked = false;
            chkLateralTargetLine.Checked = false;

            ShowGrpBottomLine();

            ScreenSetup();
        }

        private void chkBottomLineChecker_CheckedChanged(object sender, EventArgs e)
        {
            if ((chkThirdParty.Checked) && (chkBottomLineChecker.Checked))
            {
                MessageBox.Show("The 'Bottom Line Checker' cannot be used in 'Third Party Mode'.");
                chkBottomLineChecker.Checked = false;                
            }

            ShowGrpBottomLine();

            Properties.Settings.Default.UseBottomLine = chkBottomLineChecker.Checked;
            Properties.Settings.Default.Save();
            
            
            //lblBottomLineChecked.Visible = chkBottomLineChecker.Checked;
        }

        private void chkManualOrientation_CheckedChanged(object sender, EventArgs e)
        {
            if (chkManualOrientation.Checked)
            {
                chkLateralTargetLine.Checked = false;
            }

            if ((chkThirdParty.Checked) && (chkManualOrientation.Checked))
            {
                MessageBox.Show("'Manual Orientation' cannot be used in 'Third Party Mode'.");
                chkManualOrientation.Checked = false;
            }

            Properties.Settings.Default.ManualOrientation = chkManualOrientation.Checked;
            Properties.Settings.Default.Save();

            txtBitAboveBelow.Visible = chkManualOrientation.Checked;

            ScreenSetup();
        }

        private void txtSVNumOfProj_Leave(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(txtSVNumOfProj.Text))
            {
                if (Convert.ToDouble(txtSVNumOfProj.Text) >= 2)
                {
                    Properties.Settings.Default.SVProjections = txtSVNumOfProj.Text;
                    Properties.Settings.Default.Save();
                }
                else
                {
                    MessageBox.Show("The number of projections must equal to or larger than 2");
                    txtSVNumOfProj.Text = "2";
                    Properties.Settings.Default.SVProjections = txtSVNumOfProj.Text;
                    Properties.Settings.Default.Save();
                }
            }
            else
            {
                txtSVNumOfProj.Text = "2";
                Properties.Settings.Default.SVProjections = txtSVNumOfProj.Text;
                Properties.Settings.Default.Save();
            }

            iSVNumOfProj = Convert.ToInt16(txtSVNumOfProj.Text);
        }

        private void chkLateralTargetLine_CheckedChanged(object sender, EventArgs e)
        {
            if ((chkThirdParty.Checked) && (chkLateralTargetLine.Checked))
            {
                MessageBox.Show("The 'Lateral Target Line' box cannot be used in 'Third Party Mode'.");
                chkLateralTargetLine.Checked = false;
            }

            if (chkLateralTargetLine.Checked)
            {
                chkManualOrientation.Checked = false;
                grpLateralTargetLine.Visible = true;
            }
            else
            {
                grpLateralTargetLine.Visible = false;
            }
            Properties.Settings.Default.LateralTargetLine = chkLateralTargetLine.Checked;
            Properties.Settings.Default.Save();

            grpLateralTargetLine.Enabled = chkLateralTargetLine.Checked;
        }

        private void ShowGrpBottomLine()
        {
            if ((chkThirdParty.Checked) || (chkBottomLineChecker.Checked))
            {
                grpBottomLine.Visible = true;                
            }
            else
            {
                grpBottomLine.Visible = false;
            }

            if (chkThirdParty.Checked)
            {
                grpDDSurvey.Visible = false;
            }
            else
            {
                grpDDSurvey.Visible = true;
            }
        }

        private void txtTargetLineBeginningVS_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.LateralTargetLineVS = txtTargetLineBeginningVS.Text;
            Properties.Settings.Default.Save();
        }

        private void txtTargetLineBeginningTVD_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.LateralTargetLineTVD = txtTargetLineBeginningTVD.Text;
            Properties.Settings.Default.Save();
        }

        private void txtTargetLineInc_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.LateralTargetLineInc = txtTargetLineInc.Text;
            Properties.Settings.Default.Save();
        }

        private void rtxtRecipientList_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.RecipientList = rtxtRecipientList.Text;
            Properties.Settings.Default.Save();
        }

        private void rtxtCCRecipientList_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.CCRecipientList = rtxtCCRecipientList.Text;
            Properties.Settings.Default.Save();
        }

        private void rtxtMailFiles_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.MailFiles = rtxtMailFiles.Text;
            Properties.Settings.Default.Save();
        }

        private void rtxtDaySignature_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.DaySignature = rtxtDaySignature.Text;
            Properties.Settings.Default.Save();
        }

        private void rtxtNightSignature_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.NightSignature = rtxtNightSignature.Text;
            Properties.Settings.Default.Save();
        }

        private void txtMailFolder_TextChanged(object sender, EventArgs e)
        {
            if (txtMailFolder.Text == "")
            {
                txtMailFolder.Text = @"c:\";
            }
            Properties.Settings.Default.MailFolder = txtMailFolder.Text;
            Properties.Settings.Default.Save();
        }

        private void txtTemplateFilePath_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.TemplateFile = txtTemplateFilePath.Text;
            Properties.Settings.Default.Save();
        }

        private void txtActivity_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Activity = txtActivity.Text;
            Properties.Settings.Default.Save();
        }

        private void txtCounty_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.County = txtCounty.Text;
            Properties.Settings.Default.Save();
        }

        private void txtPlanNumber_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.PlanNumber = txtPlanNumber.Text;
            Properties.Settings.Default.Save();
        }

        private void txtROPSliding_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ROPSliding = txtROPSliding.Text;
            Properties.Settings.Default.Save();
        }

        private void txtDLSNeeded_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.DLSNeeded= txtDLSNeeded.Text;
            Properties.Settings.Default.Save();
        }

        private void txtRotaryTorque_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.RotaryTorque = txtRotaryTorque.Text;
            Properties.Settings.Default.Save();
        }

        private void txtRPG_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.RPG = txtRPG.Text;
            Properties.Settings.Default.Save();
        }

        private void txtSurveyCourseLength_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SurveyCourseLength = txtSurveyCourseLength.Text;
            Properties.Settings.Default.Save();
        }

        private void txtSlideAhead_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SlideAhead = txtSlideAhead.Text;
            Properties.Settings.Default.Save();
        }

        private void txtSlideSeen_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SlideSeen = txtSlideSeen.Text;
            Properties.Settings.Default.Save();
        }

        private void txtRotateWeight_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.RotateWeight = txtRotateWeight.Text;
            Properties.Settings.Default.Save();
        }

        private void txtSlackOff_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SlackOff = txtSlackOff.Text;
            Properties.Settings.Default.Save();
        }

        private void txtPickUp_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.PickUp = txtPickUp.Text;
            Properties.Settings.Default.Save();
        }

        private void txtCMTemp_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.CMTemp = txtCMTemp.Text;
            Properties.Settings.Default.Save();
        }

        private void txtAboveBelow_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.BitAboveBelow = txtBitAboveBelow.Text;
            Properties.Settings.Default.Save();
        }

        private void txtRightLeft_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.BitRightLeft = txtBitRightLeft.Text;
            Properties.Settings.Default.Save();
        }

        private void txtMWDSurveyFileName_TextChanged(object sender, EventArgs e)
        {
            if (Path.GetFileName(txtMWDSurveyFileName.Text) == "")
            {
                txtMWDSurveyFileName.Text = @"c:\MWDSurvey.pdf";
            }
            Properties.Settings.Default.MWDSurveyFile = txtMWDSurveyFileName.Text;
            Properties.Settings.Default.Save();
        }

        private void rtxtComments_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Comments = rtxtComments.Text;
            Properties.Settings.Default.Save();
        }

        private void chkComments_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.CommentsIncluded = chkComments.Checked;
            Properties.Settings.Default.Save();
        }

        private void txtSurveyAboveBelow_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SurveyAboveBelow = txtSurveyAboveBelow.Text;
            Properties.Settings.Default.Save();
        }

        private void txtSurveyRightLeft_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SurveyRightLeft = txtSurveyRightLeft.Text;
            Properties.Settings.Default.Save();
        }

        private void txtSlideTF_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SlideTF = txtSlideTF.Text;
            Properties.Settings.Default.Save();
        }

        private void txtMudWeight_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.MudWeight = txtMudWeight.Text;
            Properties.Settings.Default.Save();
        }

        private void txtECD_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ECD = txtECD.Text;
            Properties.Settings.Default.Save();
        }

        private void txtESD_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ESD = txtESD.Text;
            Properties.Settings.Default.Save();
        }

        private void SelectedSignature_CheckedChanged(object sender, EventArgs e)
        {
            if (rbDay.Checked)
            {
                Properties.Settings.Default.SigToUse = "Day";                
                rtxtDaySignature.Show();
                rtxtNightSignature.Hide();
            }
            else if (rbNight.Checked)
            {
                Properties.Settings.Default.SigToUse = "Night";
                rtxtNightSignature.Show();
                rtxtDaySignature.Hide();
            }
            else if (rbAutomatic.Checked)
            {
                Properties.Settings.Default.SigToUse = "Automatic";

                if (GetTimeOfDay() == "Day")
                {
                    rtxtDaySignature.Show();
                    rtxtNightSignature.Hide();
                }
                else
                {
                    rtxtNightSignature.Show();
                    rtxtDaySignature.Hide();
                }
            }
            Properties.Settings.Default.Save();
        }

        private string GetTimeOfDay()
        {
            string strTimeOfDay = "Night";

            TimeSpan tsStart = TimeSpan.Parse("06:00");
            TimeSpan tsEnd = TimeSpan.Parse("18:00");
            TimeSpan tsNow = DateTime.Now.TimeOfDay;

            if ((tsNow >= tsStart) && (tsNow <= tsEnd))
            {
                strTimeOfDay = "Day";
            }
                    

            return strTimeOfDay;
        }

        private void btnMailFolderPath_Click(object sender, EventArgs e)
        {
            txtMailFolder.Text = GetFolderPath(txtMailFolder.Text, false);
        }

        private void btnChangeACPath_Click(object sender, EventArgs e)
        {
            txtACFileName.Text = GetFilePath(txtACFileName.Text, "Text Files|*.txt");
        }

        private void btnChangeSurveyPath_Click(object sender, EventArgs e)
        {
            txtSurveyFileName.Text = GetFilePath(txtSurveyFileName.Text, "Text Files|*.txt");
        }

        private void btnChangeBLPath_Click(object sender, EventArgs e)
        {
            if (!chkThirdParty.Checked)
            {
                txtMWDSurveyFileName.Text = GetFilePath(txtMWDSurveyFileName.Text, "PDF Files|*.pdf");
            }
            else
            {
                txtMWDSurveyFileName.Text = GetFilePath(txtMWDSurveyFileName.Text, "Excel Files|*.xlsx");
            }
        }

        private string GetFilePath(string strFilePath, string strFileType)
        {
            OpenFileDialog ofdGetFile = new OpenFileDialog();
            ofdGetFile.InitialDirectory = Path.GetPathRoot(strFilePath);
            ofdGetFile.Filter = strFileType;

            DialogResult drResult = new DialogResult();
            drResult = ofdGetFile.ShowDialog();

            string strFileName;

            if (drResult == DialogResult.OK)
            {
                strFileName = ofdGetFile.FileName;
                ofdGetFile.Dispose();
                return strFileName;
            }
            else
            {
                ofdGetFile.Dispose();
                return strFilePath;
            }            
        }

        private string GetFolderPath(string strFolderPath, bool bShowNewFolderButton)
        {
            FolderBrowserDialog fbdGetFolder = new FolderBrowserDialog();
            fbdGetFolder.ShowNewFolderButton = bShowNewFolderButton;
            fbdGetFolder.Description = "Select Folder to Open.";
            fbdGetFolder.SelectedPath = strFolderPath;                        

            DialogResult drResult = new DialogResult();
            drResult = fbdGetFolder.ShowDialog();

            if (drResult == DialogResult.OK)
            {
                strFolderPath = fbdGetFolder.SelectedPath;
                fbdGetFolder.Dispose();
                return strFolderPath;
            }
            else
            {
                fbdGetFolder.Dispose();
                return strFolderPath;
            }
        }

        private void btnEditMailFiles_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdGetFile = new OpenFileDialog();
            ofdGetFile.InitialDirectory = txtMailFolder.Text;
            ofdGetFile.Multiselect = true;

            DialogResult drResult = new DialogResult();
            drResult = ofdGetFile.ShowDialog();

            if (drResult == DialogResult.OK)
            {
                string[] arrFiles = ofdGetFile.FileNames;

                if (arrFiles.Length > 0)
                {
                    if (Path.GetDirectoryName(arrFiles[0]) != txtMailFolder.Text)
                    {
                        txtMailFolder.Text = Path.GetDirectoryName(arrFiles[0]);
                    }
                    rtxtMailFiles.Text = "";
                    for (int i = 0; i < arrFiles.Length; i++)
                    {
                        rtxtMailFiles.Text = rtxtMailFiles.Text + Path.GetFileName(arrFiles[i]);
                        if (i < (arrFiles.Length - 1))
                        {
                            rtxtMailFiles.Text = rtxtMailFiles.Text + Environment.NewLine;
                        }
                    }                    
                }
            }
        }

        private void txtSVNumOfProj_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSVNumOfProj.Text.Replace(" ", "")))
            {
                txtSVNumOfProj.Text = "2";                
            }

            if (chkThirdParty.Checked)
            {
                txtSVNumOfProj.Text = "2";
            }

            Properties.Settings.Default.SVProjections = txtSVNumOfProj.Text;
            Properties.Settings.Default.Save();
            iSVNumOfProj = Convert.ToInt16(txtSVNumOfProj.Text);
        }

        private void Doubles_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = KeyPress_CheckIsDouble(sender, e, true, true);
        }

        private void DoublesWithoutNegative_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = KeyPress_CheckIsDouble(sender, e, false, true);
        }

        private void IntsWithoutNegative_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = KeyPress_CheckIsInt(sender, e, false);
        }

        public bool KeyPress_CheckIsDouble(object sender, KeyPressEventArgs e, bool bAllowNegative, bool bAllowDecimal)
        {
            Control cntrl = (Control)sender;
            if (Char.IsControl(e.KeyChar))
            {
                return false;
            }
            else if (e.KeyChar == Convert.ToChar("-"))
            {
                if (bAllowNegative)
                {
                    if (cntrl.Text.Length > 1)
                    {
                        if (cntrl.Text.ToString().Contains("-"))
                        {
                            for (int i = 1; i < cntrl.Text.Length; i++)
                            {
                                if (cntrl.Text[i].ToString() == "-")
                                {
                                    cntrl.Text.Remove(i - 1, 1);
                                }
                            }
                        }
                        return false;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return true;
                }
            }
            else if (e.KeyChar == Convert.ToChar("."))
            {
                if (bAllowDecimal)
                {
                    if (cntrl.Text.Contains("."))
                    {
                        e.KeyChar = Convert.ToChar(" ");
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return true;
                }
            }
            else if (e.KeyChar == Convert.ToChar(" "))
            {
                cntrl.Text.Trim();
                if (cntrl.Text.ToString().Contains(" "))
                {
                    for (int i = 0; i < cntrl.Text.Length; i++)
                    {
                        if (cntrl.Text[i].ToString() == " ")
                        {
                            cntrl.Text.ToString().Remove(i - 1, 1);
                        }
                    }
                }
                return true;
            }
            else if ((e.KeyChar != Convert.ToChar(" ")) && (e.KeyChar != Convert.ToChar(".")) && (e.KeyChar != Convert.ToChar("-")))
            {
                double dResult;
                string strValue = e.KeyChar.ToString();
                if ((strValue != "-") && (strValue != ".") && (strValue != "-."))
                {
                    if (double.TryParse(strValue, out dResult))
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                e.KeyChar = Convert.ToChar("0");
            }
            return true;
        }

        public bool KeyPress_CheckIsInt(object sender, KeyPressEventArgs e, bool bAllowNegative)
        {
            Control cntrl = (Control)sender;
            if (Char.IsControl(e.KeyChar))
            {
                return false;
            }
            else if (e.KeyChar == Convert.ToChar("-"))
            {
                if (bAllowNegative)
                {
                    if (cntrl.Text.Length > 1)
                    {
                        if (cntrl.Text.ToString().Contains("-"))
                        {
                            for (int i = 1; i < cntrl.Text.Length; i++)
                            {
                                if (cntrl.Text[i].ToString() == "-")
                                {
                                    cntrl.Text.Remove(i - 1, 1);
                                }
                            }
                        }
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return true;
                }
            }
            else if (e.KeyChar == Convert.ToChar("."))
            {
                return true;                
            }
            else if (e.KeyChar == Convert.ToChar(" "))
            {
                cntrl.Text.Trim();
                if (cntrl.Text.ToString().Contains(" "))
                {
                    for (int i = 0; i < cntrl.Text.Length; i++)
                    {
                        if (cntrl.Text[i].ToString() == " ")
                        {
                            cntrl.Text.ToString().Remove(i - 1, 1);
                        }
                    }
                }
                return true;
            }
            else if ((e.KeyChar != Convert.ToChar(" ")) && (e.KeyChar != Convert.ToChar(".")) && (e.KeyChar != Convert.ToChar("-")))
            {
                int iResult;
                string strValue = e.KeyChar.ToString();
                if ((strValue != "-") && (strValue != ".") && (strValue != "-."))
                {
                    if (int.TryParse(strValue, out iResult))
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                e.KeyChar = Convert.ToChar("0");
            }
            return true;
        }

        //Used for ComboBoxes once a selection has been made. 
        public bool IsInt(string strValue)
        {
            int iResult;
            if (int.TryParse(strValue, out iResult))
            {
                return true;
            }
            else
            {
                MessageBox.Show("Only integer values allowed. No Decimal Values.");
                return false;
            }
        }

        private void btnTemplates_Click(object sender, EventArgs e)
        {
            frmTemplates frmOpenTemplates = new frmTemplates(cboTemplates.Text, txtTemplateFilePath.Text);

            frmOpenTemplates.ShowDialog();

            dsTemplates.Tables.Clear();
            dsTemplates = frmOpenTemplates.GetBodyTemplateDataSet;
            dtSubjectTemplate = null;
            dtBodyTemplate = null;
            dtSubjectTemplate = dsTemplates.Tables["SubjectTemplate"].Copy();
            dtBodyTemplate = dsTemplates.Tables["BodyTemplate"].Copy();     

            TemplateFileChanged();            

            frmOpenTemplates.Dispose();
        }

        private void btnTemplateFilePath_Click(object sender, EventArgs e)
        {
            string strPath = GetFilePath(txtTemplateFilePath.Text, "Text Files|*.txt");

            if (File.Exists(strPath))
            {
                try
                {
                    dsTemplates.Clear();
                    dsTemplates.ReadXml(strPath);
                    cboTemplates.Items.Clear();
                    cboTemplates.Text = "";

                    for (int i = 0; i < dtSubjectTemplate.Rows.Count; i++)
                    {
                        if (!cboTemplates.Items.Contains(dtSubjectTemplate.Rows[i]["Name"].ToString()))
                        {
                            cboTemplates.Items.Add(dtSubjectTemplate.Rows[i]["Name"].ToString());
                        }

                        if (i == 0)
                        {
                            cboTemplates.Text = dtSubjectTemplate.Rows[i]["Name"].ToString();
                        }
                    }

                    txtTemplateFilePath.Text = strPath;
                    ScreenSetup();
                }
                catch
                {
                    FileIsNotValid(strPath);
                }
            }
            else
            {
                FileDoesNotExist();
            }
        }

        private void cboTemplates_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(cboTemplates.Text.Replace(" ", "")))
            {
                dtSubjectTemplate.DefaultView.RowFilter = "Name = '" + cboTemplates.Text + "'";
                dtBodyTemplate.DefaultView.RowFilter = "Name = '" + cboTemplates.Text + "'";
                ScreenSetup();
            }

            Properties.Settings.Default.CurrentTemplate = cboTemplates.Text;
            Properties.Settings.Default.Save();
        }

        private void txtTFMode_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.TFMode = txtTFMode.Text;
            Properties.Settings.Default.Save();
        }

        private void ScreenSetup()
        {
            lblJobNumber.Visible = false;
            txtJobNumber.Visible = false;
            txtJobNumber.Enabled = false;
            lblCounty.Visible = false;
            txtCounty.Visible = false;
            txtCounty.Enabled = false;
            lblRig.Visible = false;
            txtRig.Visible = false;
            txtRig.Enabled = false;
            lblClient.Visible = false;
            txtClient.Visible = false;
            txtClient.Enabled = false;
            lblPlanNumber.Visible = false;
            txtPlanNumber.Visible = false;
            txtPlanNumber.Enabled = false;
            lblWellName.Visible = false;
            txtWellName.Visible = false;
            txtWellName.Enabled = false;
            lblActivity.Visible = false;
            txtActivity.Visible = false;
            txtActivity.Enabled = false;
            lblTargetInc.Visible = false;
            txtTargetInc.Visible = false;
            txtTargetInc.Enabled = false;
            lblTargetAzm.Visible = false;
            txtTargetAzm.Visible = false;
            txtTargetAzm.Enabled = false;
            lblGamma.Visible = false;
            txtGamma.Visible = false;
            txtGamma.Enabled = false;
            lblROPRotating.Visible = false;
            txtROPRotating.Visible = false;
            txtROPRotating.Enabled = false;
            lblROPSliding.Visible = false;
            txtROPSliding.Visible = false;
            txtROPSliding.Enabled = false;
            lblSurveyTemp.Visible = false;
            txtSurveyTemp.Visible = false;
            txtSurveyTemp.Enabled = false;
            lblCMTemp.Visible = false;
            txtCMTemp.Visible = false;
            txtCMTemp.Enabled = false;
            lblFlowRate.Visible = false;
            txtFlowRate.Visible = false;
            txtFlowRate.Enabled = false;
            lblMudWeight.Visible = false;
            txtMudWeight.Visible = false;
            txtMudWeight.Enabled = false;
            lblECD.Visible = false;
            txtECD.Visible = false;
            txtECD.Enabled = false;
            lblESD.Visible = false;
            txtESD.Visible = false;
            txtESD.Enabled = false;
            lblSPP.Visible = false;
            txtSPP.Visible = false;
            txtSPP.Enabled = false;
            lblSurfaceRPM.Visible = false;
            txtSurfaceRPM.Visible = false;
            txtSurfaceRPM.Enabled = false;
            lblDifferentialPressure.Visible = false;
            txtDifferentialPressure.Visible = false;
            txtDifferentialPressure.Enabled = false;
            lblWOB.Visible = false;
            txtWOB.Visible = false;
            txtWOB.Enabled = false;
            lblRotateWeight.Visible = false;
            txtRotateWeight.Visible = false;
            txtRotateWeight.Enabled = false;
            lblSlackOff.Visible = false;
            txtSlackOff.Visible = false;
            txtSlackOff.Enabled = false;
            lblPickUp.Visible = false;
            txtPickUp.Visible = false;
            txtPickUp.Enabled = false;
            lblDLSNeeded.Visible = false;
            txtDLSNeeded.Visible = false;
            txtDLSNeeded.Enabled = false;
            lblRotaryTorque.Visible = false;
            txtRotaryTorque.Visible = false;
            txtRotaryTorque.Enabled = false;
            lblSurveyCourseLength.Visible = false;
            txtSurveyCourseLength.Visible = false;
            txtSurveyCourseLength.Enabled = false;
            lblSlideSeen.Visible = false;
            txtSlideSeen.Visible = false;
            txtSlideSeen.Enabled = false;
            chkComments.Visible = false;
            rtxtComments.Visible = false;
            rtxtComments.Enabled = false;
            lblSlideAhead.Visible = false;
            txtSlideAhead.Visible = false;
            txtSlideAhead.Enabled = false;
            lblRPG.Visible = false;
            txtRPG.Visible = false;
            txtRPG.Enabled = false;
            lblBitAboveBelow.Visible = false;
            txtBitAboveBelow.Visible = false;
            txtBitAboveBelow.Enabled = false;
            lblBitRightLeft.Visible = false;
            txtBitRightLeft.Visible = false;
            txtBitRightLeft.Enabled = false;
            lblSurveyAboveBelow.Visible = false;
            txtSurveyAboveBelow.Visible = false;
            txtSurveyAboveBelow.Enabled = false;
            lblSurveyRightLeft.Visible = false;
            txtSurveyRightLeft.Visible = false;
            txtSurveyRightLeft.Enabled = false;
            lblSlideTF.Visible = false;
            txtSlideTF.Visible = false;
            txtSlideTF.Enabled = false;
            lblPlanNorthZeroReference.Visible = false;
            txtPlanNorthZeroReference.Visible = false;
            txtPlanNorthZeroReference.Enabled = false;
            lblPlanEastZeroReference.Visible = false;
            txtPlanEastZeroReference.Visible = false;
            txtPlanEastZeroReference.Enabled = false;
            lblTFMode.Visible = false;
            txtTFMode.Visible = false;
            txtTFMode.Enabled = false;
            chkComments.Visible = false;
            rtxtComments.Visible = false;
            rtxtComments.Enabled = false;
            //rtxtMailFiles.Height = 419;

            int iLoopCount = dtSubjectTemplate.DefaultView.Count;
            DataTable dtLoopTable = dtSubjectTemplate.DefaultView.ToTable();

            int iTabIndex = 1;
            int iLocationX = 15;
            int iLocationY = 15;
            Point pLoc = new Point(iLocationX, iLocationY);
            for (int j = 0; j < 2; j++)
            {
                for (int i = 0; i < iLoopCount; i++)
                {
                    if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Job_Number")
                    {
                        if (!txtJobNumber.Enabled)
                        {
                            lblJobNumber.Visible = true;
                            lblJobNumber.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtJobNumber.Visible = true;
                            txtJobNumber.Enabled = true;
                            txtJobNumber.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtJobNumber.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "County")
                    {
                        if (!txtCounty.Enabled)
                        {
                            lblCounty.Visible = true;
                            lblCounty.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtCounty.Visible = true;
                            txtCounty.Enabled = true;
                            txtCounty.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtCounty.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Client")
                    {
                        if (!txtClient.Enabled)
                        {
                            lblClient.Visible = true;
                            lblClient.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtClient.Visible = true;
                            txtClient.Enabled = true;
                            txtClient.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtClient.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Plan_Number")
                    {
                        if (!txtPlanNumber.Enabled)
                        {
                            lblPlanNumber.Visible = true;
                            lblPlanNumber.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtPlanNumber.Visible = true;
                            txtPlanNumber.Enabled = true;
                            txtPlanNumber.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtPlanNumber.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Well_Name")
                    {
                        if (!txtWellName.Enabled)
                        {
                            lblWellName.Visible = true;
                            lblWellName.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtWellName.Visible = true;
                            txtWellName.Enabled = true;
                            txtWellName.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtWellName.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Rig")
                    {
                        if (!txtRig.Enabled)
                        {
                            lblRig.Visible = true;
                            lblRig.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtRig.Visible = true;
                            txtRig.Enabled = true;
                            txtRig.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtRig.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Activity")
                    {
                        if (!txtActivity.Enabled)
                        {
                            lblActivity.Visible = true;
                            lblActivity.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtActivity.Visible = true;
                            txtActivity.Enabled = true;
                            txtActivity.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtActivity.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Target_Inc")
                    {
                        if (!txtTargetInc.Enabled)
                        {
                            lblTargetInc.Visible = true;
                            lblTargetInc.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtTargetInc.Visible = true;
                            txtTargetInc.Enabled = true;
                            txtTargetInc.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtTargetInc.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Target_Azm")
                    {
                        if (!txtTargetAzm.Enabled)
                        {
                            lblTargetAzm.Visible = true;
                            lblTargetAzm.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtTargetAzm.Visible = true;
                            txtTargetAzm.Enabled = true;
                            txtTargetAzm.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtTargetAzm.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Gamma")
                    {
                        if (!txtGamma.Enabled)
                        {
                            lblGamma.Visible = true;
                            lblGamma.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtGamma.Visible = true;
                            txtGamma.Enabled = true;
                            txtGamma.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtGamma.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "ROP_Rotating")
                    {
                        if (!txtROPRotating.Enabled)
                        {
                            lblROPRotating.Visible = true;
                            lblROPRotating.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtROPRotating.Visible = true;
                            txtROPRotating.Enabled = true;
                            txtROPRotating.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtROPRotating.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "ROP_Sliding")
                    {
                        if (!txtROPSliding.Enabled)
                        {
                            lblROPSliding.Visible = true;
                            lblROPSliding.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtROPSliding.Visible = true;
                            txtROPSliding.Enabled = true;
                            txtROPSliding.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtROPSliding.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Survey_Temp")
                    {
                        if (!txtSurveyTemp.Enabled)
                        {
                            lblSurveyTemp.Visible = true;
                            lblSurveyTemp.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtSurveyTemp.Visible = true;
                            txtSurveyTemp.Enabled = true;
                            txtSurveyTemp.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtSurveyTemp.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "CM_Temp")
                    {
                        if (!txtCMTemp.Enabled)
                        {
                            lblCMTemp.Visible = true;
                            lblCMTemp.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtCMTemp.Visible = true;
                            txtCMTemp.Enabled = true;
                            txtCMTemp.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtCMTemp.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Flow_Rate")
                    {
                        if (!txtFlowRate.Enabled)
                        {
                            lblFlowRate.Visible = true;
                            lblFlowRate.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtFlowRate.Visible = true;
                            txtFlowRate.Enabled = true;
                            txtFlowRate.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtFlowRate.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Mud_Weight")
                    {
                        if (!txtMudWeight.Enabled)
                        {
                            lblMudWeight.Visible = true;
                            lblMudWeight.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtMudWeight.Visible = true;
                            txtMudWeight.Enabled = true;
                            txtMudWeight.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtMudWeight.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "ECD")
                    {
                        if (!txtECD.Enabled)
                        {
                            lblECD.Visible = true;
                            lblECD.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtECD.Visible = true;
                            txtECD.Enabled = true;
                            txtECD.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtECD.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "ESD")
                    {
                        if (!txtESD.Enabled)
                        {
                            lblESD.Visible = true;
                            lblESD.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtESD.Visible = true;
                            txtESD.Enabled = true;
                            txtESD.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtESD.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "SPP")
                    {
                        if (!txtSPP.Enabled)
                        {
                            lblSPP.Visible = true;
                            lblSPP.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtSPP.Visible = true;
                            txtSPP.Enabled = true;
                            txtSPP.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtSPP.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Rev/Gal")
                    {
                        if (!txtRPG.Enabled)
                        {
                            lblRPG.Visible = true;
                            lblRPG.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtRPG.Visible = true;
                            txtRPG.Enabled = true;
                            txtRPG.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtRPG.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Surface_RPM")
                    {
                        if (!txtSurfaceRPM.Enabled)
                        {
                            lblSurfaceRPM.Visible = true;
                            lblSurfaceRPM.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtSurfaceRPM.Visible = true;
                            txtSurfaceRPM.Enabled = true;
                            txtSurfaceRPM.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtSurfaceRPM.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Motor_RPM")
                    {
                        if (!txtRPG.Enabled)
                        {
                            lblRPG.Visible = true;
                            lblRPG.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtRPG.Visible = true;
                            txtRPG.Enabled = true;
                            txtRPG.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtRPG.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }

                        if (!txtFlowRate.Enabled)
                        {
                            lblFlowRate.Visible = true;
                            lblFlowRate.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtFlowRate.Visible = true;
                            txtFlowRate.Enabled = true;
                            txtFlowRate.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtFlowRate.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Slide_Rotate_Footage")
                    {
                        if (!txtSlideSeen.Enabled)
                        {
                            lblSlideSeen.Visible = true;
                            lblSlideSeen.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtSlideSeen.Visible = true;
                            txtSlideSeen.Enabled = true;
                            txtSlideSeen.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtSlideSeen.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }

                        if (!txtSurveyCourseLength.Enabled)
                        {
                            lblSurveyCourseLength.Visible = true;
                            lblSurveyCourseLength.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtSurveyCourseLength.Visible = true;
                            txtSurveyCourseLength.Enabled = true;
                            txtSurveyCourseLength.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtSurveyCourseLength.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Bit_RPM")
                    {
                        if (!txtSurfaceRPM.Enabled)
                        {
                            lblSurfaceRPM.Visible = true;
                            lblSurfaceRPM.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtSurfaceRPM.Visible = true;
                            txtSurfaceRPM.Enabled = true;
                            txtSurfaceRPM.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtSurfaceRPM.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }

                        if (!txtRPG.Enabled)
                        {
                            lblRPG.Visible = true;
                            lblRPG.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtRPG.Visible = true;
                            txtRPG.Enabled = true;
                            txtRPG.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtRPG.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }

                        if (!txtFlowRate.Enabled)
                        {
                            lblFlowRate.Visible = true;
                            lblFlowRate.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtFlowRate.Visible = true;
                            txtFlowRate.Enabled = true;
                            txtFlowRate.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtFlowRate.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Differential_Pressure")
                    {
                        if (!txtDifferentialPressure.Enabled)
                        {
                            lblDifferentialPressure.Visible = true;
                            lblDifferentialPressure.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtDifferentialPressure.Visible = true;
                            txtDifferentialPressure.Enabled = true;
                            txtDifferentialPressure.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtDifferentialPressure.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "WOB")
                    {
                        if (!txtWOB.Enabled)
                        {
                            lblWOB.Visible = true;
                            lblWOB.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtWOB.Visible = true;
                            txtWOB.Enabled = true;
                            txtWOB.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtWOB.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Rotate_Weight")
                    {
                        if (!txtRotateWeight.Enabled)
                        {
                            lblRotateWeight.Visible = true;
                            lblRotateWeight.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtRotateWeight.Visible = true;
                            txtRotateWeight.Enabled = true;
                            txtRotateWeight.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtRotateWeight.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Pick_Up")
                    {
                        if (!txtPickUp.Enabled)
                        {
                            lblPickUp.Visible = true;
                            lblPickUp.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtPickUp.Visible = true;
                            txtPickUp.Enabled = true;
                            txtPickUp.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtPickUp.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Slack_Off")
                    {
                        if (!txtSlackOff.Enabled)
                        {
                            lblSlackOff.Visible = true;
                            lblSlackOff.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtSlackOff.Visible = true;
                            txtSlackOff.Enabled = true;
                            txtSlackOff.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtSlackOff.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Tool_Face_Mode")
                    {
                        if (!txtTFMode.Enabled)
                        {
                            lblTFMode.Visible = true;
                            lblTFMode.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtTFMode.Visible = true;
                            txtTFMode.Enabled = true;
                            txtTFMode.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtTFMode.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "DLS_Needed")
                    {
                        if (!txtDLSNeeded.Enabled)
                        {
                            lblDLSNeeded.Visible = true;
                            lblDLSNeeded.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtDLSNeeded.Visible = true;
                            txtDLSNeeded.Enabled = true;
                            txtDLSNeeded.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtDLSNeeded.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Rotary_Torque")
                    {
                        if (!txtRotaryTorque.Enabled)
                        {
                            lblRotaryTorque.Visible = true;
                            lblRotaryTorque.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtRotaryTorque.Visible = true;
                            txtRotaryTorque.Enabled = true;
                            txtRotaryTorque.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtRotaryTorque.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Motor_Yield")
                    {
                        if (!txtSurveyCourseLength.Enabled)
                        {
                            lblSurveyCourseLength.Visible = true;
                            lblSurveyCourseLength.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtSurveyCourseLength.Visible = true;
                            txtSurveyCourseLength.Enabled = true;
                            txtSurveyCourseLength.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtSurveyCourseLength.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }

                        if (!txtSlideSeen.Enabled)
                        {
                            lblSlideSeen.Visible = true;
                            lblSlideSeen.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtSlideSeen.Visible = true;
                            txtSlideSeen.Enabled = true;
                            txtSlideSeen.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtSlideSeen.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Slide_Ahead")
                    {
                        if (!txtSlideAhead.Enabled)
                        {
                            lblSlideAhead.Visible = true;
                            lblSlideAhead.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtSlideAhead.Visible = true;
                            txtSlideAhead.Enabled = true;
                            txtSlideAhead.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtSlideAhead.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Slide_Seen")
                    {
                        if (!txtSlideSeen.Enabled)
                        {
                            lblSlideSeen.Visible = true;
                            lblSlideSeen.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtSlideSeen.Visible = true;
                            txtSlideSeen.Enabled = true;
                            txtSlideSeen.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtSlideSeen.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Slide_Tool_Face")
                    {
                        if (!txtSlideTF.Enabled)
                        {
                            lblSlideTF.Visible = true;
                            lblSlideTF.Location = pLoc;
                            iLocationY = iLocationY + 15;
                            pLoc.Y = iLocationY;
                            txtSlideTF.Visible = true;
                            txtSlideTF.Enabled = true;
                            txtSlideTF.Location = pLoc;
                            iLocationY = iLocationY + 20;
                            txtSlideTF.TabIndex = iTabIndex;
                            iTabIndex = iTabIndex + 1;

                            iLocationX = CheckLocationX(iLocationX, iLocationY);
                            iLocationY = CheckLocationY(iLocationX, iLocationY);
                            pLoc.X = iLocationX;
                            pLoc.Y = iLocationY;
                        }
                    }
                    else if ((dtLoopTable.Rows[i]["DataSource"].ToString() == "Bit_A/B_Plan") &&
                             (chkThirdParty.Checked))
                    {
                        lblBitAboveBelow.Visible = true;
                        lblBitAboveBelow.Location = pLoc;
                        iLocationY = iLocationY + 15;
                        pLoc.Y = iLocationY;
                        txtBitAboveBelow.Visible = true;
                        txtBitAboveBelow.Enabled = true;
                        txtBitAboveBelow.Location = pLoc;
                        iLocationY = iLocationY + 20;
                        txtBitAboveBelow.TabIndex = iTabIndex;
                        iTabIndex = iTabIndex + 1;

                        iLocationX = CheckLocationX(iLocationX, iLocationY);
                        iLocationY = CheckLocationY(iLocationX, iLocationY);
                        pLoc.X = iLocationX;
                        pLoc.Y = iLocationY;
                    }
                    else if ((dtLoopTable.Rows[i]["DataSource"].ToString() == "Bit_R/L_Plan") &&
                             (chkThirdParty.Checked))
                    {
                        lblBitRightLeft.Visible = true;
                        lblBitRightLeft.Location = pLoc;
                        iLocationY = iLocationY + 15;
                        pLoc.Y = iLocationY;
                        txtBitRightLeft.Visible = true;
                        txtBitRightLeft.Enabled = true;
                        txtBitRightLeft.Location = pLoc;
                        iLocationY = iLocationY + 20;
                        txtBitRightLeft.TabIndex = iTabIndex;
                        iTabIndex = iTabIndex + 1;

                        iLocationX = CheckLocationX(iLocationX, iLocationY);
                        iLocationY = CheckLocationY(iLocationX, iLocationY);
                        pLoc.X = iLocationX;
                        pLoc.Y = iLocationY;
                    }

                    else if ((dtLoopTable.Rows[i]["DataSource"].ToString() == "Survey_A/B_Plan") &&
                             (chkThirdParty.Checked))
                    {
                        lblSurveyAboveBelow.Visible = true;
                        lblSurveyAboveBelow.Location = pLoc;
                        iLocationY = iLocationY + 15;
                        pLoc.Y = iLocationY;
                        txtSurveyAboveBelow.Visible = true;
                        txtSurveyAboveBelow.Enabled = true;
                        txtSurveyAboveBelow.Location = pLoc;
                        iLocationY = iLocationY + 20;
                        txtSurveyAboveBelow.TabIndex = iTabIndex;
                        iTabIndex = iTabIndex + 1;

                        iLocationX = CheckLocationX(iLocationX, iLocationY);
                        iLocationY = CheckLocationY(iLocationX, iLocationY);
                        pLoc.X = iLocationX;
                        pLoc.Y = iLocationY;
                    }
                    else if ((dtLoopTable.Rows[i]["DataSource"].ToString() == "Survey_R/L_Plan") &&
                             (chkThirdParty.Checked))
                    {
                        lblSurveyRightLeft.Visible = true;
                        lblSurveyRightLeft.Location = pLoc;
                        iLocationY = iLocationY + 15;
                        pLoc.Y = iLocationY;
                        txtSurveyRightLeft.Visible = true;
                        txtSurveyRightLeft.Enabled = true;
                        txtSurveyRightLeft.Location = pLoc;
                        iLocationY = iLocationY + 20;
                        txtSurveyRightLeft.TabIndex = iTabIndex;
                        iTabIndex = iTabIndex + 1;

                        iLocationX = CheckLocationX(iLocationX, iLocationY);
                        iLocationY = CheckLocationY(iLocationX, iLocationY);
                        pLoc.X = iLocationX;
                        pLoc.Y = iLocationY;
                    }
                    else if (dtLoopTable.Rows[i]["DataSource"].ToString() == "Comments")
                    {
                        chkComments.Visible = true;
                        rtxtComments.Visible = true;
                        rtxtComments.Enabled = true;
                        if (chkComments.Checked)
                        {
                            //rtxtMailFiles.Height = 322;
                        }
                    }
                }
                iLoopCount = dtBodyTemplate.DefaultView.Count;
                dtLoopTable.Clear();
                dtLoopTable = dtBodyTemplate.DefaultView.ToTable();
            }

            dtLoopTable.Dispose();

            if (chkManualOrientation.Checked)
            {
                lblBitAboveBelow.Visible = true;
                lblBitAboveBelow.Location = pLoc;
                iLocationY = iLocationY + 15;
                pLoc.Y = iLocationY;
                txtBitAboveBelow.Visible = true;
                txtBitAboveBelow.Enabled = true;
                txtBitAboveBelow.Location = pLoc;
                iLocationY = iLocationY + 20;
                txtBitAboveBelow.TabIndex = iTabIndex;
                iTabIndex = iTabIndex + 1;

                iLocationX = CheckLocationX(iLocationX, iLocationY);
                iLocationY = CheckLocationY(iLocationX, iLocationY);
                pLoc.X = iLocationX;
                pLoc.Y = iLocationY;
            }

            if (chkChangePlanNSEWZeroReference.Checked)
            {
                lblPlanNorthZeroReference.Visible = true;
                lblPlanNorthZeroReference.Location = pLoc;
                iLocationY = iLocationY + 15;
                pLoc.Y = iLocationY;
                txtPlanNorthZeroReference.Visible = true;
                txtPlanNorthZeroReference.Enabled = true;
                txtPlanNorthZeroReference.Location = pLoc;
                iLocationY = iLocationY + 20;
                txtPlanNorthZeroReference.TabIndex = iTabIndex;
                iTabIndex = iTabIndex + 1;

                iLocationX = CheckLocationX(iLocationX, iLocationY);
                iLocationY = CheckLocationY(iLocationX, iLocationY);
                pLoc.X = iLocationX;
                pLoc.Y = iLocationY;

                lblPlanEastZeroReference.Visible = true;
                lblPlanEastZeroReference.Location = pLoc;
                iLocationY = iLocationY + 15;
                pLoc.Y = iLocationY;
                txtPlanEastZeroReference.Visible = true;
                txtPlanEastZeroReference.Enabled = true;
                txtPlanEastZeroReference.Location = pLoc;
                iLocationY = iLocationY + 20;
                txtPlanEastZeroReference.TabIndex = iTabIndex;
                iTabIndex = iTabIndex + 1;

                iLocationX = CheckLocationX(iLocationX, iLocationY);
                iLocationY = CheckLocationY(iLocationX, iLocationY);
                pLoc.X = iLocationX;
                pLoc.Y = iLocationY;
            }
            
        }

        private int CheckLocationX(int iLocX, int iLocY)
        {
            if (iLocY >= 415)
            {
                iLocX = iLocX + 140; 
            }
            return iLocX;
        }

        private int CheckLocationY(int iLocX, int iLocY)
        {
            if (iLocY >= 415)
            {
                iLocY = 15;
            }
            return iLocY;
        }
    }
}
