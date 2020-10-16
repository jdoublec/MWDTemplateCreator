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
    public partial class frmBodySample : Form
    {

        

        public frmBodySample(string strSubject, string strTemp)
        {
            InitializeComponent();

            txtSubject.Text = strSubject;
            rtxtTemplate.Rtf = strTemp;
        }
    }
}
