using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Addon_SIA
{
    public partial class frmReportViewer : Form
    {        
        private bool bClose = false;
        private object _myParentForm = null;
        private bool _isB1 = false;
        public frmReportViewer(bool bCanClose, ref object frmParent)
        {
            InitializeComponent();
            bClose = bCanClose;
            _myParentForm = frmParent;
            _isB1 = false;            
        }
        public frmReportViewer(bool IsB1)
        {
            InitializeComponent();
            _isB1 = true;
        }
        private void frmReportViewer_Activated(object sender, EventArgs e)
        {
            if (bClose == true)
            {
                this.Close();
            }
        }

        private void frmReportViewer_Load(object sender, EventArgs e)
        {   
            this.Text = @"SAP Business One : Crystal Reports - Vista Previa";
            if (bClose == true)
            {
                this.Visible = false;
                this.BringToFront();
            }
            if (bClose != true)
            {
                ConfigureCrystalReports();
            }
            this.BringToFront();            
        }
        private void ConfigureCrystalReports()
        {
            //CrystalDecisions.CrystalReports.Engine.ReportDocument rd = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            //rd = csVariablesGlobales.crReport;
            //CrystalDecisions.Shared.TableLogOnInfo myLogin;
            //csUtilidades csUtilidades = new csUtilidades();
            //csUtilidades.LeerConexion(false);
            //foreach (CrystalDecisions.CrystalReports.Engine.Table myTable in rd.Database.Tables )
            //{
            //    myLogin = myTable.LogOnInfo;
            //    myLogin.ConnectionInfo.Password = csVariablesGlobales.DBPassword;
            //    myLogin.ConnectionInfo.UserID = csVariablesGlobales.oCompany.DbUserName;
            //    myLogin.ConnectionInfo.DatabaseName = csVariablesGlobales.oCompany.CompanyDB;
            //    myLogin.ConnectionInfo.ServerName = csVariablesGlobales.oCompany.Server;
            //    myTable.ApplyLogOnInfo(myLogin); 
            //}
            //rd.Refresh(); 
            //crystalReportViewer.ReportSource = rd;              
            //rd.Refresh();
            crystalReportViewer.ReportSource = csVariablesGlobales.crReport;
            csVariablesGlobales.crReport.Refresh();
        }

        private void frmReportViewer_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (!_isB1)
            {
                Form ParForm = null;
                ParForm = (Form)_myParentForm;
                ParForm.Visible = true;
            }
        }        
    }
}