using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MWD_Template_Creator
{
    public partial class frmBottomLineForm : Form
    {
        private bool bRedAlert;
        private bool bSurveysFound;
        private string strSurveyMessage;
        private string strMessage;

        public frmBottomLineForm(bool bSurveysFoundTemp, bool bRedAlertTemp, string strSurveyMessageTemp, string strMessageTemp)
        {
            InitializeComponent();

            bSurveysFound = bSurveysFoundTemp;
            bRedAlert = bRedAlertTemp;
            strSurveyMessage = strSurveyMessageTemp;
            strMessage = strMessageTemp;
        }

        private void frmBottomLineForm_Load(object sender, EventArgs e)
        {
            if (!bSurveysFound)
            {
                lblMessage.Text = strSurveyMessage;
            }
            else
            {
                lblMessage.Text = strMessage;
            }
            if ((bRedAlert) || (!bSurveysFound))
            {
                this.BackColor = Color.Red;
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {            
            if ((bRedAlert) && (!bSurveysFound))
            {
                lblMessage.Text = strMessage;
                bRedAlert = false;
                bSurveysFound = true;
            }
            else
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
