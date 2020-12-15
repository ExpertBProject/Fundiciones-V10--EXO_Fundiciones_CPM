using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Diagnostics;
using System.Threading;

namespace Addon_SIA
{
    class csFrmImportacionesSIA
    {
        private SAPbouiCOM.Form oForm;
        SAPbouiCOM.Item oItem;
        SAPbouiCOM.Button oButton;
        SAPbouiCOM.StaticText oStaticText;
        SAPbouiCOM.EditText oEditText;
        SAPbouiCOM.FormCreationParams oCreationParams;

        public void CargarFormulario()
        {
            CrearFormularioImportacionesSIA();
            oForm.Visible = true;
            //csUtilidades csUtilidades = new csUtilidades();
            csUtilidades.SaveAsXml(oForm, "FrmImportacionesSIA.xml", "");
            // events handled by SBO_Application_ItemEvent
            csVariablesGlobales.SboApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(FrmImportacionesSIA_ItemEvent);
            // events handled by SBO_Application_AppEvent
            csVariablesGlobales.SboApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(FrmImportacionesSIA_AppEvent);
        }

        private void CrearFormularioImportacionesSIA()
        {
            int BaseLeft = 0;
            int BaseTop = 0;            

            #region Formulario
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "FrmImportacionesSIA";
            oCreationParams.FormType = "SIASL00001";

            oForm = csVariablesGlobales.SboApp.Forms.AddEx(oCreationParams);
            // set the form properties
            oForm.Title = "Importaciones SIA";
            oForm.Left = 300;
            oForm.ClientWidth = 500;
            oForm.Top = 100;
            oForm.ClientHeight = 120;
            #endregion

            #region Data Sources
            //  Adding a User Data Source
            oForm.DataSources.UserDataSources.Add("dsSelFiCo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsSelFiRe", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            #endregion

            #region Marcos
            #region Importaciones SIA
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
            oItem = oForm.Items.Add("lblImpInt", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 100;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Importaciones SIA";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #endregion

            #region Botones
            //*****************************************
            // Adding Items to the form
            // and setting their properties
            //*****************************************
            #region Selecciona Fichero Consultas
            //************************
            // Adding a Cancel button
            //***********************

            oItem = oForm.Items.Add("btnSelFiCo", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 440;
            oItem.Width = 20;
            oItem.Top = 25;
            oItem.Height = 19;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "...";
            #endregion
            #region Selecciona Fichero Reports
            //************************
            // Adding a Cancel button
            //***********************

            oItem = oForm.Items.Add("btnSelFiRe", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 440;
            oItem.Width = 20;
            oItem.Top = 45;
            oItem.Height = 19;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "...";
            #endregion
            #region btnCancelar
            //************************
            // Adding a Cancel button
            //***********************

            oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 75;
            oItem.Width = 65;
            oItem.Top = 85;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";
            #endregion
            #region Importar
            // /**********************
            // Adding an Procesos
            //*********************

            // We get automatic event handling for
            // the Ok and Cancel Buttons by setting
            // their UIDs to 1 and 2 respectively

            oItem = oForm.Items.Add("btnImp", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5;
            oItem.Width = 65;
            oItem.Top = 85;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Importar";
            #endregion
            #endregion

            #region Campos
            #region Selecciona Fichero Consultas
            //*************************
            // Adding a Text Edit item
            //*************************
            BaseLeft = 10;
            BaseTop = 30;

            oItem = oForm.Items.Add("txtSelFiCo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 150;
            oItem.Width = 280;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.Enabled = false;
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsSelFiCo");

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblSelFiCo", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 140;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtSelFiCo";

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Selecciona Fichero Consultas";
            #endregion
            #region Selecciona Fichero Reports
            //*************************
            // Adding a Text Edit item
            //*************************
            BaseLeft = 10;
            BaseTop = 50;

            oItem = oForm.Items.Add("txtSelFiRe", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 150;
            oItem.Width = 280;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.Enabled = false;
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsSelFiRe");

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblSelFiRe", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 140;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtSelFiRe";

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Selecciona Fichero Reports";
            #endregion
            #endregion
        }

        public void FrmImportacionesSIA_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //csUtilidades csUtilidades = new csUtilidades();
            csCrearConsultas CrearConsultas = new csCrearConsultas();
            if (FormUID == "FrmImportacionesSIA")
            {
                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == "btnSelFiCo" || pVal.ItemUID == "btnSelFiRe")
                        {
                            if (pVal.ItemUID.ToString().Substring(0, 8) == "btnSelFi" && !pVal.BeforeAction && pVal.ActionSuccess)
                            {
                                csOpenFileDialog OpenFileDialog = new csOpenFileDialog();
                                OpenFileDialog.Filter = "All Files (*)|*|Dat (*.dat)|*.dat|Text Files (*.txt)|*.txt";
                                OpenFileDialog.InitialDirectory = csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_PARAM]", "U_Ruta", "U_TipoRuta = 'Destino Ficheros SIA'");
                                Thread threadGetFile = new Thread(new ThreadStart(OpenFileDialog.GetFileName));
                                threadGetFile.TrySetApartmentState(ApartmentState.STA);
                                try
                                {
                                    threadGetFile.Start();
                                    while (!threadGetFile.IsAlive) ; // Wait for thread to get started
                                    Thread.Sleep(1);  // Wait a sec more
                                    threadGetFile.Join();    // Wait for thread to end

                                    // Use file name as you will here
                                    string strValue = OpenFileDialog.FileName;
                                    oForm.DataSources.UserDataSources.Item("dsSelFi" + pVal.ItemUID.ToString().Substring(8, 2)).ValueEx = strValue;
                                }
                                catch (Exception ex)
                                {
                                    csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "OK", "", "");
                                }
                                threadGetFile.Abort();
                                threadGetFile = null;
                                OpenFileDialog.InitialDirectory = "";
                                OpenFileDialog = null;
                            }
                        }
                        if (pVal.ItemUID == "btnImp" && !pVal.BeforeAction && pVal.ActionSuccess)
                        {
                            try
                            {
                                #region Fichero Consultas
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtSelFiCo").Specific;
                                if (oEditText.String != "")
                                {
                                    csVariablesGlobales.SboApp.SetStatusBarMessage("Leyendo fichero de importación de consultas SIA", BoMessageTime.bmt_Short, false);
                                    char[] CaracteresDelimitadores = { '|' };
                                    System.IO.StreamReader srLinea = new System.IO.StreamReader(oEditText.String,
                                                                                                System.Text.Encoding.Default,
                                                                                                true);
                                    while (srLinea.Peek() != -1)
                                    {
                                        string Linea = srLinea.ReadLine();
                                        string[] Columnas = Linea.Split(CaracteresDelimitadores);
                                        string Categoria = Columnas.GetValue(0).ToString();
                                        string Descripcion = Columnas.GetValue(1).ToString();
                                        string Consulta = Columnas.GetValue(2).ToString();
                                        CrearConsultas.CrearConsulta(Categoria, Consulta, Descripcion);
                                    }
                                    srLinea.Close();
                                    csVariablesGlobales.SboApp.SetStatusBarMessage("Fichero leído completamente. Los posibles cambios se han realizado", BoMessageTime.bmt_Short, false);
                                    csVariablesGlobales.SboApp.MessageBox("Proceso importación de consultas finalizado", 1, "OK", "", "");
                                }
                                #endregion
                                #region Fichero Reports
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtSelFiRe").Specific;
                                if (oEditText.String != "")
                                {
                                    csVariablesGlobales.SboApp.SetStatusBarMessage("Leyendo fichero de importación de reports SIA", BoMessageTime.bmt_Short, false);
                                    char[] CaracteresDelimitadores = { '|' };
                                    System.IO.StreamReader srLinea = new System.IO.StreamReader(oEditText.String,
                                                                                                System.Text.Encoding.Default,
                                                                                                true);
                                    while (srLinea.Peek() != -1)
                                    {
                                        string Linea = srLinea.ReadLine();
                                        string[] Columnas = Linea.Split(CaracteresDelimitadores);
                                        string NombreReport = Columnas.GetValue(0).ToString();
                                        string Descripcion = Columnas.GetValue(1).ToString();
                                        string TipoDocumento = Columnas.GetValue(2).ToString();
                                        string Borrador = Columnas.GetValue(3).ToString();
                                        csUtilidades.InsertaRegistroTablaReport("" + csVariablesGlobales.Prefijo + "_REPORT", NombreReport, Descripcion, TipoDocumento, Borrador);
                                    }
                                    srLinea.Close();
                                    csVariablesGlobales.SboApp.SetStatusBarMessage("Fichero leído completamente. Los posibles cambios se han realizado", BoMessageTime.bmt_Short, false);
                                    csVariablesGlobales.SboApp.MessageBox("Proceso importación de reports finalizado", 1, "OK", "", "");
                                }
                                #endregion
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                            }

                        }
                        break;
                }
            }
        }

        private void FrmImportacionesSIA_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //System.Windows.Forms.Application.Exit();
                    break;
            }
        }
    }
}
