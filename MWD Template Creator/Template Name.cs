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
    public partial class frmTemplateName : Form
    {
        private bool bOKToAdd = false;

        public string GetTemplateName
        {
            get
            {
                return txtTemplateName.Text;
            }
        }

        public string GetTemplateToDuplicate
        {
            get
            {
                return cboTemplates.Text;
            }
        }

        public bool OKToAdd
        {
            get
            {
                return bOKToAdd;
            }
            set
            {
                bOKToAdd = value;
            }
        }

        public string LabelText
        {
            set
            {
                lblTemplateName.Text = value;
            }
        }

        public string FormText
        {
            set
            {
                this.Text = value;
            }
        }

        public frmTemplateName(List<string> lstTemplates)
        {
            InitializeComponent();

            btnOKDuplicate.Enabled = false;

            cboTemplates.Show();
            txtTemplateName.Hide();
            btnOKDuplicate.Text = "Duplicate";

            for (int i = 0; i < lstTemplates.Count; i++)
            {
                cboTemplates.Items.Add(lstTemplates[i]);
            }            
        }

        public frmTemplateName()
        {
            InitializeComponent();

            btnOKDuplicate.Enabled = false;

            txtTemplateName.Show();
            cboTemplates.Hide();
            btnOKDuplicate.Text = "Ok";
        }

        private void txtTemplateName_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtTemplateName.Text.Replace(" ", "")))
            {
                btnOKDuplicate.Enabled = true;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            OKToAdd = false;
            txtTemplateName.Text = "";
            this.Close();
        }

        private void btnOkDuplicate_Click(object sender, EventArgs e)
        {
            OKToAdd = true;
            this.Close();
        }

        private void cboTemplates_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(cboTemplates.Text.Replace(" ", "")))
            {
                btnOKDuplicate.Enabled = true;
            }
        }

        
    }
}
