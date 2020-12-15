using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Data.SqlClient;

namespace Addon_SIA
{
    class csFrmAsientoGastos
    {
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Item oItem = null;
        private SAPbouiCOM.Button oButton = null;
        private SAPbouiCOM.StaticText oStaticText = null;
        private SAPbouiCOM.EditText oEditText = null;
        private SAPbouiCOM.LinkedButton oLinkedButton = null;
        private SAPbouiCOM.ComboBox oComboBox = null;


        public void CargarFormulario()
        {
            CrearFormularioAsientoGastos();
            oForm.Visible = true;
            //csUtilidades csUtilidades = new csUtilidades();
            csUtilidades.SaveAsXml(oForm, "FrmAsientoGastos.xml", "");
            csVariablesGlobales.SboApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(FrmAsientoGastos_ItemEvent);
            csVariablesGlobales.SboApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(FrmAsientoGastos_AppEvent);
        }

        private void CrearFormularioAsientoGastos()
        {
            int BaseLeft = 0;
            int BaseTop = 0;

            #region Formulario
            SAPbouiCOM.FormCreationParams oCreationParams = null;
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "FrmAsientoGastos";
            oCreationParams.FormType = "SIASL00003";

            oForm = csVariablesGlobales.SboApp.Forms.AddEx(oCreationParams);

            // set the form properties
            oForm.Title = "Asiento Gastos";
            oForm.Left = 300;
            oForm.ClientWidth = 500;
            oForm.Top = 100;
            oForm.ClientHeight = 120;
            //oForm.EnableMenu("1293", true);
            #endregion

            #region Data Sources
            //  Adding a User Data Source
            oForm.DataSources.UserDataSources.Add("dsImp", SAPbouiCOM.BoDataType.dt_SUM, 254);
            oForm.DataSources.UserDataSources.Add("dsDeb", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("dsHab", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("dsMon", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            #endregion

            #region Marcos
            #region Asiento
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
            oItem = oForm.Items.Add("lblAsiento", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 50;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Asiento";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #endregion

            AddChooseFromList();

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
            #region Botón Crear Asiento
            // /**********************
            // Adding an Ok button
            //*********************

            // We get automatic event handling for
            // the Ok and Cancel Buttons by setting
            // their UIDs to 1 and 2 respectively

            oItem = oForm.Items.Add("btnCreAsi", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 80;
            oItem.Width = 105;
            oItem.Top = 85;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Crear Asiento Gastos";
            #endregion
            #endregion

            #region Campos
            #region Cuenta Debe
            BaseLeft = 10;
            BaseTop = 30;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtDeb", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblDeb";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsDeb");
            oEditText.ChooseFromListUID = "CFLCueDeb";
            oEditText.ChooseFromListAlias = "AcctCode";

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblDeb", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtDeb";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Cuenta Debe";

            // Link the column to the BP master data system form
            oItem = oForm.Items.Add("lkbDeb", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            //Set the properties
            oItem.Top = BaseTop;
            oItem.Left = BaseLeft + 105;
            oItem.LinkTo = "txtDeb";
            oLinkedButton = (SAPbouiCOM.LinkedButton)oItem.Specific;
            oLinkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts;

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Cuenta Haber
            BaseLeft = 240;
            BaseTop = 30;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtHab", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblHab";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsHab");
            oEditText.ChooseFromListUID = "CFLCueHab";
            oEditText.ChooseFromListAlias = "AcctCode";

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblHab", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtHab";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Cuenta Haber";

            // Link the column to the BP master data system form
            oItem = oForm.Items.Add("lkbHab", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            //Set the properties
            oItem.Top = BaseTop;
            oItem.Left = BaseLeft + 105;
            oItem.LinkTo = "txtHab";
            oLinkedButton = (SAPbouiCOM.LinkedButton)oItem.Specific;
            oLinkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts;

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Importe
            // Calc Picker
            //__________________________________________________________________________________________
            ////*************************
            //// Adding a Text Edit item
            ////*************************
            BaseLeft = 10;
            BaseTop = 50;

            oItem = oForm.Items.Add("txtImp", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            //// bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsImp");

            ////**********************************
            //// Adding Static Text item
            ////**********************************

            oItem = oForm.Items.Add("lblImp", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 80;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtImp";
            oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            oStaticText.Caption = "Importe";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Moneda
            BaseLeft = 240;
            BaseTop = 50;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtMon", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblMon";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsMon");
            oEditText.ChooseFromListUID = "CFLMon";
            oEditText.ChooseFromListAlias = "CurrCode";

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblMon", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtMon";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Moneda";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #endregion
        }

        public void FrmAsientoGastos_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            //csUtilidades csUtilidades = new csUtilidades();
            BubbleEvent = true;

            if (pVal.FormUID == "FrmAsientoGastos")
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                        #region Crear Asiento
                        if (pVal.ItemUID == "btnCreAsi" && !pVal.BeforeAction && pVal.ActionSuccess)
                        {
                            if (ValidarProceso())
                            {
                                int RetVal;
                                string Serie;
                                string RefDate;
                                string DueDate;
                                string TaxDate;
                                string Ref2;
                                SAPbobsCOM.JournalEntries oJournalEntries;
                                SAPbobsCOM.JournalEntries_Lines oJournalEntriesLineas;
                                oJournalEntries = (SAPbobsCOM.JournalEntries)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries));
                                oJournalEntries.GetByKey(csVariablesGlobales.NumeroAsiento);
                                Serie = oJournalEntries.Series.ToString();
                                RefDate = oJournalEntries.ReferenceDate.ToString();
                                DueDate = oJournalEntries.DueDate.ToString();
                                TaxDate = oJournalEntries.TaxDate.ToString();
                                Ref2 = oJournalEntries.Reference2.ToString();
                                oJournalEntries = null;
                                oJournalEntries = (SAPbobsCOM.JournalEntries)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries));
                                oJournalEntries.Series = Convert.ToInt32(Serie);
                                oJournalEntries.ReferenceDate = Convert.ToDateTime(RefDate);
                                oJournalEntries.DueDate = Convert.ToDateTime(DueDate);
                                oJournalEntries.TaxDate = Convert.ToDateTime(TaxDate);
                                oJournalEntries.Reference2 = Ref2;
                                oJournalEntriesLineas = oJournalEntries.Lines;
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtDeb").Specific;
                                oJournalEntriesLineas.ShortName = oEditText.String;
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtMon").Specific;
                                if (oEditText.String == "EUR" || oEditText.String == "")
                                {
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtImp").Specific;
                                    //                                    oJournalEntriesLineas.Debit = Convert.ToDouble(oEditText.String);
                                    oJournalEntriesLineas.Debit = csUtilidades.TextoADouble(oEditText.String);
                                }
                                else
                                {
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtMon").Specific;
                                    oJournalEntriesLineas.FCCurrency = oEditText.Value;
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtImp").Specific;
                                    oJournalEntriesLineas.FCDebit = csUtilidades.TextoADouble(oEditText.String);
                                }
                                oJournalEntriesLineas.Add();
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtHab").Specific;
                                oJournalEntriesLineas.ShortName = oEditText.String;
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtMon").Specific;
                                if (oEditText.String == "EUR" || oEditText.String == "")
                                {
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtImp").Specific;
                                    oJournalEntriesLineas.Credit = csUtilidades.TextoADouble(oEditText.String);
                                }
                                else
                                {
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtMon").Specific;
                                    oJournalEntriesLineas.FCCurrency = oEditText.Value;
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtImp").Specific;
                                    oJournalEntriesLineas.FCCredit = csUtilidades.TextoADouble(oEditText.String);
                                }
                                oJournalEntriesLineas.AdditionalReference = csVariablesGlobales.NumeroAsiento.ToString();
                                RetVal = oJournalEntries.Add();
                                if (RetVal != 0)
                                {
                                    csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                                }
                                else
                                {
                                    csVariablesGlobales.SboApp.MessageBox("Asiento generado correctamente", 1, "", "", "");
                                    csVariablesGlobales.FormularioAsientoGastosAbierto = false;
                                    oForm.Close();
                                }
                            }
                        }
                        #endregion
                        break;

                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                        string sCFL_ID = null;
                        sCFL_ID = oCFLEvento.ChooseFromListUID;
                        SAPbouiCOM.ChooseFromList oCFL = null;
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                        if (!oCFLEvento.BeforeAction)
                        {
                            SAPbouiCOM.DataTable oDataTable = null;
                            oDataTable = oCFLEvento.SelectedObjects;
                            string Codigo = null;
                            try
                            {
                                Codigo = System.Convert.ToString(oDataTable.GetValue(0, 0));
                                if (pVal.ItemUID == "txtDeb")
                                {
                                    oForm.DataSources.UserDataSources.Item("dsDeb").ValueEx = Codigo;
                                }
                                if (pVal.ItemUID == "txtHab")
                                {
                                    oForm.DataSources.UserDataSources.Item("dsHab").ValueEx = Codigo;
                                }
                                if (pVal.ItemUID == "txtMon")
                                {
                                    oForm.DataSources.UserDataSources.Item("dsMon").ValueEx = Codigo;
                                }
                            }
                            catch
                            {
                                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                            }

                        }
                        break;
                }
            }
        }

        private void FrmAsientoGastos_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
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

        private void AddChooseFromList()
        {
            try
            {

                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                SAPbouiCOM.Conditions oCons = null;
                SAPbouiCOM.Condition oCon = null;

                oCFLs = oForm.ChooseFromLists;

                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                //  Adding CFL para Cuentas Debe
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "1";
                oCFLCreationParams.UniqueID = "CFLCueDeb";
                oCFL = oCFLs.Add(oCFLCreationParams);

                oCons = oCFL.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "Postable";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "Y";
                oCFL.SetConditions(oCons);

                //  Adding CFL para Cuentas Habe
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "1";
                oCFLCreationParams.UniqueID = "CFLCueHab";
                oCFL = oCFLs.Add(oCFLCreationParams);

                oCons = oCFL.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "Postable";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "Y";
                oCFL.SetConditions(oCons);

                //  Adding CFL para Cuentas Habe
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "37";
                oCFLCreationParams.UniqueID = "CFLMon";
                oCFL = oCFLs.Add(oCFLCreationParams);
            }
            catch
            {
                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
            }
        }

        public void CargarDatosInicialesDePantalla()
        {
            SAPbobsCOM.JournalEntries oJournalEntries;
            SAPbobsCOM.JournalEntries_Lines oJournalEntriesLineas;
            oJournalEntries = (SAPbobsCOM.JournalEntries)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries));
            oForm = csVariablesGlobales.SboApp.Forms.Item("FrmAsientoGastos");
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtMon").Specific;
            oJournalEntries.GetByKey(csVariablesGlobales.NumeroAsiento);
            oJournalEntriesLineas = oJournalEntries.Lines;
            oJournalEntriesLineas.SetCurrentLine(0);
            oEditText.String = oJournalEntriesLineas.FCCurrency.ToString();
        }

        private bool ValidarProceso()
        {
            try
            {
                //csUtilidades csUtilidades = new csUtilidades();
                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtDeb").Specific;
                if (oEditText.String == "")
                {
                    oEditText.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("La cuenta 'debe' debe existir", 1, "Ok", "", "");
                    return false;
                }
                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtHab").Specific;
                if (oEditText.String == "")
                {
                    oEditText.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("La cuenta 'haber' debe existir", 1, "Ok", "", "");
                    return false;
                }
                //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtImp").Specific;
                //if (csUtilidades.TextoADouble(oEditText.String) <= 0.00)
                //{
                //    oEditText.Active = true;
                //    csVariablesGlobales.SboApp.MessageBox("El importe debe ser mayor que 0", 1, "Ok", "", "");
                //    return false;
                //}
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
