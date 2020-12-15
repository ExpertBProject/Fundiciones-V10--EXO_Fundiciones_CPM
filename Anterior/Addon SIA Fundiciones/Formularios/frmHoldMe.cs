using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Addon_SIA
{
    public partial class frmHoldMe : Form
    {
        private SAPbouiCOM.Application mobjSBOApplication;        
        public SAPbouiCOM.Application AppObject
        {
            set
            {
                mobjSBOApplication = value;
            }
            get
            {
                return mobjSBOApplication;
            }
        }

        public frmHoldMe()
        {
            InitializeComponent();
            TimerAccountSetup.Interval = 400;
            TimerAccountSetup.Start();

            TimerManageGroups.Interval = 400;
            TimerManageGroups.Start();

            TimerManageGrpUsers.Interval = 400;
            TimerManageGrpUsers.Start();
            
            TimerManageReports.Interval = 400;
            TimerManageReports.Start();
            
            TimerCrystalReports.Interval = 400;
            TimerCrystalReports.Start();
            
            TimerCrystalViewer.Interval = 400;
            TimerCrystalViewer.Start();
        }

        private void frmHoldMe_FormClosing(object sender, FormClosingEventArgs e)
        {        
            e.Cancel = true;
        }

        private bool isFormExisting( string FormName )
        {
            bool bExistingFormFound = false;
            foreach (Form f in Application.OpenForms)
            {
                if (f.Name == FormName)
                {
                    bExistingFormFound = true;
                    break;
                }
            }
            return bExistingFormFound;
        }

        private void TimerCrystalViewer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (csVariablesGlobales.LanzarImpresionCrystal)
                {
                    csVariablesGlobales.LanzarImpresionCrystal = false;
                    if (isFormExisting("frmReportViewer") == false)
                    {
                        showCrystalViewerWindow();// this will show a dialog so that funciton will finish before going to line below.
                    }
                }
            }
            catch
            {
                
            }
        }

        private void showCrystalViewerWindow()
        {
            frmReportViewer formReportViewer = new frmReportViewer(true );                
            formReportViewer.Show();
        }
    }
}