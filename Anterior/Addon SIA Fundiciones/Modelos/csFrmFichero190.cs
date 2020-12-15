using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Diagnostics;
using System.Data;
using System.Data.SqlClient;

namespace Addon_SIA
{
    class csFrmModelo190
    {
        private SAPbouiCOM.Form oForm;

        public void CargarFormulario()
        {
            CrearFormulario190();
            oForm.Visible = true;
            //csUtilidades Utilidades = new csUtilidades();
            csUtilidades.SaveAsXml(oForm, "FrmModelo190.xml", "");
            // events handled by SBO_Application_ItemEvent
            csVariablesGlobales.SboApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(FrmModelo190_ItemEvent);
            // events handled by SBO_Application_AppEvent
            csVariablesGlobales.SboApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(FrmModelo190_AppEvent);
        }

        private void CrearFormulario190()
        {
            int BaseLeft = 0;
            int BaseTop = 0;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;
            SAPbouiCOM.CheckBox oCheckBox = null;
            SAPbouiCOM.StaticText oStaticText = null;
            SAPbouiCOM.EditText oEditText = null;

            #region Formulario
            SAPbouiCOM.FormCreationParams oCreationParams = null;
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "FrmModelo190";
            oCreationParams.FormType = "2001060008";

            oForm = csVariablesGlobales.SboApp.Forms.AddEx(oCreationParams);

            // set the form properties
            oForm.Title = "Fichero 190";
            oForm.Left = 300;
            oForm.ClientWidth = 500;
            oForm.Top = 100;
            oForm.ClientHeight = 400;
            #endregion

            #region Data Sources
            //  Adding a User Data Source
            oForm.DataSources.UserDataSources.Add("dsDeud", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("dsDesdeFec", SAPbouiCOM.BoDataType.dt_DATE, 254);
            oForm.DataSources.UserDataSources.Add("dsHastaFec", SAPbouiCOM.BoDataType.dt_DATE, 254);
            oForm.DataSources.UserDataSources.Add("dsSelFich", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            #endregion

            #region Marcos
            #region Rango IC
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
            oItem.Height = 100;

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblRangoIC", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 50;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Rango IC";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #endregion

            #region Botones
            //*****************************************
            // Adding Items to the form
            // and setting their properties
            //*****************************************
            #region btnOk
            // /**********************
            // Adding an Ok button
            //*********************

            // We get automatic event handling for
            // the Ok and Cancel Buttons by setting
            // their UIDs to 1 and 2 respectively

            oItem = oForm.Items.Add("btnGene", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5;
            oItem.Width = 65;
            oItem.Top = 365;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Generar";
            #endregion
            #region btnCancelar
            //************************
            // Adding a Cancel button
            //***********************

            oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 75;
            oItem.Width = 65;
            oItem.Top = 365;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";
            #endregion
            #region Selecciona Fichero
            //************************
            // Adding a Cancel button
            //***********************

            oItem = oForm.Items.Add("btnSelFich", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 440;
            oItem.Width = 20;
            oItem.Top = 275;
            oItem.Height = 19;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "...";
            #endregion
            #endregion

            #region Campos
            #region Desde Fecha
            BaseLeft = 10;
            BaseTop = 150;

            // Date Picker
            //__________________________________________________________________________________________
            ////*************************
            //// Adding a Text Edit item
            ////*************************

            oItem = oForm.Items.Add("txtDesFec", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            //// bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsDesdeFec");

            ////**********************************
            //// Adding Static Text item
            ////**********************************

            oItem = oForm.Items.Add("lblDesFec", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtDesFec";
            oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            oStaticText.Caption = "Desde Fecha";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Hasta Fecha
            BaseLeft = 10;
            BaseTop = 170;

            // Date Picker
            //__________________________________________________________________________________________
            ////*************************
            //// Adding a Text Edit item
            ////*************************

            oItem = oForm.Items.Add("txtHasFec", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            //// bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsHastaFec");

            ////**********************************
            //// Adding Static Text item
            ////**********************************

            oItem = oForm.Items.Add("lblHasFec", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtHasFec";
            oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            oStaticText.Caption = "Hasta Fecha";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Selecciona Fichero
            //*************************
            // Adding a Text Edit item
            //*************************
            BaseLeft = 10;
            BaseTop = 280;

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

            #region CheckBox
            #region CheckBox Deudores
            BaseLeft = 30;
            BaseTop = 30;
            oItem = oForm.Items.Add("chkDeud", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 19;
            oCheckBox = ((SAPbouiCOM.CheckBox)(oItem.Specific));
            oCheckBox.Caption = "Deudores";
            // binding the Check box with a data source
            oCheckBox.DataBind.SetBound(true, "", "dsDeud");
            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #endregion
        }

        private void FrmModelo190_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {

            BubbleEvent = true;

            if (FormUID == "FrmModelo190")
            {

                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                        //System.Windows.Forms.Application.Exit();
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        try
                        {
                            if (pVal.ItemUID == "1" && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                            {
                                //Genera347();
                                BubbleEvent = false;
                            }
                            //if (b == false && pVal.ItemUID == "btnSelFich" && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                            if (pVal.ItemUID == "btnSelFich" && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                            {
                                //SBOUtilyArch.cLanzarReport Archivo = new SBOUtilyArch.cLanzarReport();
                                string dll;
                                //dll = Archivo.FNArchivo(@"C:\", "txt files (*.txt)|*.txt", "Fichero");
                                //oForm.DataSources.UserDataSources.Item("dsSelFich").ValueEx = dll;
                                //b = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message == "Se ha seleccionado Cancelar.")
                            {
                                oForm.DataSources.UserDataSources.Item("dsSelFich").ValueEx = "";
                            }
                            else
                            {
                                MessageBox.Show(ex.Message.ToString());
                            }
                        }
                        break;
                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                        string sCFL_ID = null;
                        sCFL_ID = oCFLEvento.ChooseFromListUID;
                        SAPbouiCOM.ChooseFromList oCFL = null;
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                        if (oCFLEvento.BeforeAction == false)
                        {
                            SAPbouiCOM.DataTable oDataTable = null;
                            oDataTable = oCFLEvento.SelectedObjects;
                            string val = null;
                            try
                            {
                                val = System.Convert.ToString(oDataTable.GetValue(0, 0));
                            }
                            catch
                            {
                                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                            }
                            if ((pVal.ItemUID == "txtDesdeIC") | (pVal.ItemUID == "btnSelDIC"))
                            {
                                oForm.DataSources.UserDataSources.Item("dsDesdeIC").ValueEx = val;
                            }
                            if ((pVal.ItemUID == "txtHastaIC") | (pVal.ItemUID == "btnSelHIC"))
                            {
                                oForm.DataSources.UserDataSources.Item("dsHastaIC").ValueEx = val;
                            }
                        }
                        break;
                }
            }
        }

        private void FrmModelo190_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {

            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:

                    csVariablesGlobales.SboApp.MessageBox("Se canceló la ejecución de SAP" + Environment.NewLine + "Se cerrará el Addon SIA", 1, "Ok", "", "");
                    //System.Windows.Forms.Application.Exit();
                    break;
            }
        }
    }
}
