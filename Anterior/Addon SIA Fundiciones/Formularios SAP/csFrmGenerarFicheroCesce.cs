using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Diagnostics;
using System.Data;
using System.Data.SqlClient;
using System.Threading;

namespace Addon_SIA
{
    class csFrmGenerarFicheroCesce
    {
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Item oItem = null;
        private SAPbouiCOM.Button oButton = null;
        private SAPbouiCOM.StaticText oStaticText = null;
        private SAPbouiCOM.EditText oEditText = null;
        private SAPbouiCOM.LinkedButton oLinkedButton = null;

        public void CargarFormulario()
        {
            CrearFormularioGenerarFicheroCesce();
            oForm.Visible = true;
            csUtilidades Utilidades = new csUtilidades();
            csUtilidades.SaveAsXml(oForm, "FrmGenerarFicheroCesce.xml", "");
            // events handled by SBO_Application_ItemEvent
            csVariablesGlobales.SboApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(FrmGenerarFicheroCesce_ItemEvent);
            // events handled by SBO_Application_AppEvent
            csVariablesGlobales.SboApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(FrmGenerarFicheroCesce_AppEvent);
        }

        private void CrearFormularioGenerarFicheroCesce()
        {
            int BaseLeft = 0;
            int BaseTop = 0;

            #region Formulario
            SAPbouiCOM.FormCreationParams oCreationParams = null;
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "FrmGenerarFicheroCesce";
            oCreationParams.FormType = "SIASL10004";

            oForm = csVariablesGlobales.SboApp.Forms.AddEx(oCreationParams);

            // set the form properties
            oForm.Title = "Generar Fichero CESCE";
            oForm.Left = 300;
            oForm.ClientWidth = 500;
            oForm.Top = 100;
            oForm.ClientHeight = 120;
            //oForm.EnableMenu("1293", true);
            #endregion

            #region Data Sources
            //  Adding a User Data Source
            oForm.DataSources.UserDataSources.Add("dsDesFec", SAPbouiCOM.BoDataType.dt_DATE, 254);
            oForm.DataSources.UserDataSources.Add("dsHasFec", SAPbouiCOM.BoDataType.dt_DATE, 254);
            oForm.DataSources.UserDataSources.Add("dsSelFich", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            #endregion

            #region Marcos
            #region Fichero Cesce
            BaseLeft = 5;
            BaseTop = 15;

            //****************************
            // Adding a Rectangle Item
            // for cosmetic purposes only
            //****************************
            oItem = oForm.Items.Add("Rect1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = BaseLeft;
            oItem.Width = 480;
            oItem.Top = BaseTop;
            oItem.Height = 60;

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblFicCes", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 75;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Fichero CESCE";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #endregion

            #region Botones
            //*****************************************
            // Adding Items to the form
            // and setting their properties
            //*****************************************
            #region btnCancelar
            //************************
            // Adding a Cancel button
            //***********************

            oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5;
            oItem.Width = 65;
            oItem.Top = 85;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";
            #endregion
            #region Botón Generar Fichero
            // /**********************
            // Adding an Ok button
            //*********************

            // We get automatic event handling for
            // the Ok and Cancel Buttons by setting
            // their UIDs to 1 and 2 respectively

            oItem = oForm.Items.Add("btnGenFic", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 80;
            oItem.Width = 105;
            oItem.Top = 85;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Generar Fichero";
            #endregion
            #region Selecciona Fichero
            //************************
            // Adding a Cancel button
            //***********************

            oItem = oForm.Items.Add("btnSelFich", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 440;
            oItem.Width = 20;
            oItem.Top = 45;
            oItem.Height = 19;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "...";
            #endregion
            #endregion

            #region Campos
            #region Desde Fecha
            BaseLeft = 10;
            BaseTop = 30;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtDesFec", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblDesFec";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsDesFec");

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblDesFec", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtDesFec";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Desde Fecha";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Cuenta Haber
            BaseLeft = 240;
            BaseTop = 30;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtHasFec", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblHasFec";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsHasFec");

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblHasFec", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtHasFec";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Hasta Fecha";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Selecciona Fichero
            //*************************
            // Adding a Text Edit item
            //*************************
            BaseLeft = 10;
            BaseTop = 50;

            oItem = oForm.Items.Add("txtSelFich", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 310;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsSelFich");

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblSelFich", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtSelFich";

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Selecciona Fichero";
            #endregion
            #endregion
        }

        public void FrmGenerarFicheroCesce_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            csUtilidades Utilidades = new csUtilidades();
            BubbleEvent = true;

            if (pVal.FormUID == "FrmGenerarFicheroCesce")
            {

//                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                        #region Generar Fichero
                        if (pVal.ItemUID == "btnGenFic" && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                        {
                            if (ValidarProceso())
                            {
                                GenerarFicheroCesce();
                            }
                        }
                        if (pVal.ItemUID == "btnSelFich" && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                        {
                            //SBOUtilyArch.cLanzarReport Archivo = new SBOUtilyArch.cLanzarReport();
                            //string dll;
                            //dll = Archivo.FNArchivo(@"C:\", "txt files (*.txt)|*.txt", "Fichero");
                            //oForm.DataSources.UserDataSources.Item("dsSelFich").ValueEx = dll;
                            csOpenFileDialog OpenFileDialog = new csOpenFileDialog();
                            OpenFileDialog.Filter = "All Files (*)|*|Dat (*.dat)|*.dat|Text Files (*.txt)|*.txt";
                            //OpenFileDialog.InitialDirectory =
                            //    Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                            OpenFileDialog.InitialDirectory = csUtilidades.DameValor("[@SIA_PARAM]", "U_Ruta", "U_TipoRuta = 'Destino Ficheros Varios'");
                            Thread threadGetExcelFile = new Thread(new ThreadStart(OpenFileDialog.GetFileName));
                            threadGetExcelFile.ApartmentState = ApartmentState.STA;
                            try
                            {
                                threadGetExcelFile.Start();
                                while (!threadGetExcelFile.IsAlive) ; // Wait for thread to get started
                                Thread.Sleep(1);  // Wait a sec more
                                threadGetExcelFile.Join();    // Wait for thread to end

                                // Use file name as you will here
                                string strValue = OpenFileDialog.FileName;
                                oForm.DataSources.UserDataSources.Item("dsSelFich").ValueEx = strValue;
                            }
                            catch (Exception ex)
                            {
                                csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "OK", "", "");
                            }
                            threadGetExcelFile = null;
                            OpenFileDialog = null;
                        }
                        #endregion
                        break;
                }
            }
        }

        private void FrmGenerarFicheroCesce_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {

            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:

                    //csVariablesGlobales.SboApp.MessageBox("A Shut Down Event has been caught" + Environment.NewLine + "Terminating 'Complex Form' Add On...", 1, "Ok", "", "");

                    ////**************************************************************
                    ////
                    //// Take care of terminating your AddOn application
                    ////
                    ////**************************************************************

                    //System.Windows.Forms.Application.Exit();

                    break;
            }
        }

        private bool ValidarProceso()
        {
            try
            {
                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtDesFec").Specific;
                if (oEditText.String == "")
                {
                    oEditText.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("La fecha 'desde' debe existir", 1, "Ok", "", "");
                    return false;
                }
                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtHasFec").Specific;
                if (oEditText.String == "")
                {
                    oEditText.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("La fecha 'hasta' debe existir", 1, "Ok", "", "");
                    return false;
                }
                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtSelFich").Specific;
                if (oEditText.String == "")
                {
                    oEditText.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("El fichero destino debe existir", 1, "Ok", "", "");
                    return false;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void GenerarFicheroCesce()
        {
            SAPbobsCOM.Documents oInvoice;
            SAPbobsCOM.BusinessPartners oBusinessPartners;
            SAPbobsCOM.Series oSeries;
            csUtilidades Utilidades = new csUtilidades();
            string DesdeFecha;
            string HastaFecha;
            string StrLinea;
            string StrSQL;

            oInvoice = (SAPbobsCOM.Documents)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices));
            oBusinessPartners = (SAPbobsCOM.BusinessPartners)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners));
            //oSeries = (SAPbobsCOM.Series)(csVariablesGlobales.Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSeri));

            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtDesFec").Specific;
            DesdeFecha = oEditText.Value;
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtHasFec").Specific;
            HastaFecha = oEditText.Value;
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtSelFich").Specific;
            System.IO.StreamWriter SW = new System.IO.StreamWriter(oEditText.String, false);

            StrSQL = "SELECT A.DocEntry " +
                     "FROM OINV A, OCRD B " +
                     "WHERE TaxDate BETWEEN '" + DesdeFecha + "' AND '" + HastaFecha + "' " +
                     "AND A.CardCode = B.CardCode AND B.U_SIA_cesce1 <> '' AND B.U_SIA_cesce2 <> ''";

            SqlDataAdapter daFactura = new SqlDataAdapter(StrSQL, csVariablesGlobales.conAddon);
            System.Data.DataTable dtFactura = new System.Data.DataTable();
            DataRow drFactura;
            daFactura.Fill(dtFactura);
            if (dtFactura.Rows.Count > 0)
            {
                for (int i = 0; i < dtFactura.Rows.Count; i++)
                {
                    drFactura = dtFactura.Rows[i];
                    oInvoice.GetByKey(Convert.ToInt32(drFactura["DocEntry"].ToString()));
                    oBusinessPartners.GetByKey(oInvoice.CardCode.ToString());

                    StrLinea = " ,     " +
                               oBusinessPartners.UserFields.Fields.Item("U_SIA_cesce2").Value.ToString().PadRight(10).ToString() +
                               ",     " +
                               oBusinessPartners.FederalTaxID.ToString().PadRight(19).ToString() +
                               ",                   " +
                               ",      " +
                               (csUtilidades.DameValor("NNM1", "SeriesName", "ObjectCode ='13' AND Series ='" + oInvoice.Series.ToString() + "'") + oInvoice.DocNum.ToString()).ToString().PadRight(17).ToString() +
                               ",      " +
                               oInvoice.DocDueDate.Day.ToString().PadLeft(2, '0').ToString() + "/" + oInvoice.DocDueDate.Month.ToString().PadLeft(2, '0').ToString() + "/" + oInvoice.DocDueDate.Year.ToString().Substring(0, 2).ToString() +
                               "  ," +
                               oInvoice.DocTotal.ToString().Replace('.', ',').PadLeft(19).ToString() +
                               "  ," +
                               "     0    ," +
                               "     " +
                               oInvoice.DocDate.Day.ToString().PadLeft(2, '0').ToString() + "/" + oInvoice.DocDate.Month.ToString().PadLeft(2, '0').ToString() + "/" + oInvoice.DocDate.Year.ToString().Substring(0, 3).ToString() +
                               "  ,    ";
                    //oInvoice.DocNum.ToString() + ", " +
                    //oInvoice.DocDueDate.ToShortDateString() + ", " +
                    //oInvoice.DocTotal.ToString().Replace(',', '.') + ", 0, " + oInvoice.TaxDate.ToShortDateString();

                    SW.WriteLine(StrLinea);
                }
                SW.Close();
                csVariablesGlobales.SboApp.MessageBox("Fichero creado correctamente", 1, "", "", "");
            }
        }
    }
}
