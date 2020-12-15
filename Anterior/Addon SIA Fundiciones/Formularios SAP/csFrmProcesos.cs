using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Data;
using System.Data.SqlClient;



namespace Addon_SIA
{
    class csFrmProcesos
    {
        //const string Estructura = "[a-z]";
        //static readonly Regex Estructura_Regex = new Regex(Estructura);

        SAPbouiCOM.UserDataSource oUserDataSourceArticulo = null;
        SAPbouiCOM.UserDataSource oUserDataSourceArticuloDescripcion = null;
        SAPbouiCOM.UserDataSource oUserDataSourceCantidad = null;
        SAPbouiCOM.UserDataSource oUserDataSourceUbicacionComponente = null;
        SAPbouiCOM.UserDataSource oUserDataSourceAlmacenComponente = null;
        SAPbouiCOM.UserDataSource oUserDataSourceBaseLine = null;
        private SAPbouiCOM.Form oForm;

        //static bool TodasMayusculas(string Cadena)
        //{
        //    if (Estructura_Regex.IsMatch(Cadena))
        //    {
        //        return false;
        //    }
        //    return true;
        //}

        public void CargarFormulario()
        {
            CrearFormularioProcesos();
            oForm.Visible = true;
            csUtilidades Utilidades = new csUtilidades();
            csUtilidades.SaveAsXml(oForm, "FrmProcesos.xml", "");
            // events handled by SBO_Application_ItemEvent
            csVariablesGlobales.SboApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(FrmProcesos_ItemEvent);
            // events handled by SBO_Application_AppEvent
            csVariablesGlobales.SboApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(FrmProcesos_AppEvent);
        }

        private void CrearFormularioProcesos()
        {
            int BaseLeft = 0;
            int BaseTop = 0;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;
            SAPbouiCOM.CheckBox oCheckBox = null;
            SAPbouiCOM.StaticText oStaticText = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.Columns oColumns = null;
            SAPbouiCOM.Column oColumn = null;

            #region Formulario
            SAPbouiCOM.FormCreationParams oCreationParams = null;
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "FrmProcesos";
            oCreationParams.FormType = "SIASL10002";

            oForm = csVariablesGlobales.SboApp.Forms.AddEx(oCreationParams);

            //oForm.DataSources.UserDataSources.Add("FolderDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            //oForm.DataSources.UserDataSources.Add("OpBtnDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            //oForm.DataSources.UserDataSources.Add("CheckDS1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            //oForm.DataSources.UserDataSources.Add("CheckDS2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            //oForm.DataSources.UserDataSources.Add("CheckDS3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

            // set the form properties
            oForm.Title = "Procesos";
            oForm.Left = 300;
            oForm.ClientWidth = 500;
            oForm.Top = 100;
            oForm.ClientHeight = 325;
            #endregion

            #region Data Sources
            //  Adding a User Data Source
            oForm.DataSources.UserDataSources.Add("dsNumOrd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsUbiArt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsAlmArt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsSerie", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("dsFecPed", SAPbouiCOM.BoDataType.dt_DATE, 254);
            oForm.DataSources.UserDataSources.Add("dsTipo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsRef2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsCant", SAPbouiCOM.BoDataType.dt_QUANTITY, 20);
            oForm.DataSources.UserDataSources.Add("dsArt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsColAC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            oForm.DataSources.UserDataSources.Add("dsColUC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            oForm.DataSources.UserDataSources.Add("dsColArt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsColAD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsColCan", SAPbouiCOM.BoDataType.dt_QUANTITY, 20);
            oForm.DataSources.UserDataSources.Add("dsDPC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("dsBasLin", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10);
            #endregion

            #region Marcos
            #region Procesos
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
            oItem = oForm.Items.Add("lblDesmon", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 60;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Procesos";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Detalle
            BaseLeft = 5;
            BaseTop = 130;

            //****************************
            // Adding a Rectangle Item
            // for cosmetic purposes only
            //****************************
            oItem = oForm.Items.Add("Rect2", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = BaseLeft;
            oItem.Width = 480;
            oItem.Top = BaseTop;
            oItem.Height = 150;

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblLin", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 80;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Líneas Proceso";

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
            oItem.Top = 290;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";
            #endregion
            #region Botón Procesos
            // /**********************
            // Adding an Procesos
            //*********************

            // We get automatic event handling for
            // the Ok and Cancel Buttons by setting
            // their UIDs to 1 and 2 respectively

            oItem = oForm.Items.Add("btnProc", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 80;
            oItem.Width = 105;
            oItem.Top = 290;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Procesar Orden";
            #endregion
            #endregion

            #region Labels
            #region Almacen
            BaseTop = 0;
            BaseLeft = 0;
            oItem = oForm.Items.Add("lblAlm", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 0;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "";

            BaseTop = 0;
            BaseLeft = 0;
            #endregion
            #endregion

            #region Campos
            #region Número Orden
            BaseLeft = 10;
            BaseTop = 30;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtNumOrd", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblNumOrd";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsNumOrd");
            //oEditText.String = "Hasta IC";
            oEditText.ChooseFromListUID = "CFL1Ord";
            oEditText.ChooseFromListAlias = "DocNum";

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblNumOrd", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtNumOrd";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Número Orden";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Ubicación Artículo
            BaseLeft = 240;
            BaseTop = 70;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtUbiArt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblUbiArt";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsUbiArt");

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblUbiArt", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtUbiArt";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Ubicación Artículo";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Almacén Artículo
            BaseLeft = 10;
            BaseTop = 70;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtAlmArt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblAlmArt";
            oItem.Enabled = false;
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsAlmArt");

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblAlmArt", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 110;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtAlmArt";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Almacén Artículo";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Cantidad
            // Calc Picker
            //__________________________________________________________________________________________
            ////*************************
            //// Adding a Text Edit item
            ////*************************
            BaseLeft = 10;
            BaseTop = 90;

            oItem = oForm.Items.Add("txtCant", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.Enabled = false;
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            //// bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsCant");

            ////**********************************
            //// Adding Static Text item
            ////**********************************

            oItem = oForm.Items.Add("lblCant", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 80;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtCant";
            oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            oStaticText.Caption = "Cantidad";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Artículo
            ////*************************
            //// Adding a Text Edit item
            ////*************************
            BaseLeft = 240;
            BaseTop = 50;

            oItem = oForm.Items.Add("txtArt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.Enabled = false;
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            //// bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsArt");

            ////**********************************
            //// Adding Static Text item
            ////**********************************

            oItem = oForm.Items.Add("lblArt", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 80;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtArt";
            oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            oStaticText.Caption = "Artículo";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Fecha Pedido
            BaseLeft = 240;
            BaseTop = 30;

            // Date Picker
            //__________________________________________________________________________________________
            ////*************************
            //// Adding a Text Edit item
            ////*************************

            oItem = oForm.Items.Add("txtFecPed", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.Enabled = false;
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            //// bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsFecPed");
            //oEditText.String = DateTime.Today.Day.ToString() + "/" + DateTime.Today.Month.ToString() + "/" + DateTime.Today.Year.ToString();

            ////**********************************
            //// Adding Static Text item
            ////**********************************

            oItem = oForm.Items.Add("lblFecPed", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtFecPed";
            oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            oStaticText.Caption = "Fecha Pedido";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Tipo
            BaseLeft = 10;
            BaseTop = 50;

            // Date Picker
            //__________________________________________________________________________________________
            ////*************************
            //// Adding a Text Edit item
            ////*************************

            oItem = oForm.Items.Add("txtTipo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.Enabled = false;
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            //// bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsTipo");

            ////**********************************
            //// Adding Static Text item
            ////**********************************

            oItem = oForm.Items.Add("lblTipo", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtTipo";
            oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            oStaticText.Caption = "Tipo";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #endregion

            #region Combos
            #region Combo Series
            ////*************************
            //// Adding a Combo Box item
            ////*************************
            //BaseLeft = 10;
            //BaseTop = 50;

            //oItem = oForm.Items.Add("cmbSerie", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            //oItem.Left = BaseLeft + 120;
            //oItem.Width = 100;
            //oItem.Top = BaseTop;
            //oItem.Height = 14;

            //oItem.DisplayDesc = true;

            //oComboBox = ((SAPbouiCOM.ComboBox)(oItem.Specific));

            //// bind the Combo Box item to the defined used data source
            //oComboBox.DataBind.SetBound(true, "", "dsSerie");

            //oComboBox.ValidValues.LoadSeries("202", SAPbouiCOM.BoSeriesMode.sf_Add); //Orden Fabricación

            ////***************************
            //// Adding a Static Text item
            ////***************************
            //oItem = oForm.Items.Add("lblSerie", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            //oItem.Left = BaseLeft;
            //oItem.Width = 100;
            //oItem.Top = BaseTop;
            //oItem.Height = 14;
            //oItem.LinkTo = "cmbSerie";
            //oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            //oStaticText.Caption = "Serie";

            //BaseLeft = 0;
            //BaseTop = 0;
            #endregion
            #endregion

            #region Matrix
            BaseTop = 135;
            BaseLeft = 15;

            //***************************
            // Adding a Matrix item
            //***************************

            oItem = oForm.Items.Add("matDet", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oItem.Left = BaseLeft;
            oItem.Width = 465;
            oItem.Top = BaseTop;
            oItem.Height = 140;

            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oColumns = oMatrix.Columns;

            //***********************************
            // Adding Culomn items to the matrix
            //***********************************
            #region Columna #
            oColumn = oColumns.Add("col#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;
            #endregion
            #region Columna Código Artículo
            // Add a column for Codigo Articulo
            oColumn = oColumns.Add("colArtic", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Número Artículo";
            oColumn.Width = 80;
            oColumn.Editable = false;

            // bind the text edit item to the defined used data source
            oColumn.DataBind.SetBound(true, "", "dsColArt");
            #endregion
            #region Columna Descripción Artículo
            // Add a column for Item Name
            oColumn = oColumns.Add("colDesArt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Descripción de Artículo";
            oColumn.Width = 100;
            oColumn.Editable = false;

            // bind the text edit item to the defined used data source
            oColumn.DataBind.SetBound(true, "", "dsColAD");
            #endregion
            #region Columna Cantidad
            // Add a column for Cantidad
            oColumn = oColumns.Add("colCant", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Cantidad";
            oColumn.Width = 60;
            oColumn.Editable = false;

            // bind the text edit item to the defined used data source
            oColumn.DataBind.SetBound(true, "", "dsColCan");
            #endregion
            #region Columna Almacén Componente
            // Add a column for Ubicación Destino
            oColumn = oColumns.Add("colAlmCom", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Alm. Com.";
            oColumn.Width = 80;
            oColumn.Editable = false;

            // bind the text edit item to the defined used data source
            oColumn.DataBind.SetBound(true, "", "dsColAC");
            #endregion
            #region Columna Ubicación Componentes
            // Add a column for Ubicación Destino
            oColumn = oColumns.Add("colUbiCom", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ubic. Com.";
            oColumn.Width = 80;
            oColumn.Editable = true;

            // bind the text edit item to the defined used data source
            oColumn.DataBind.SetBound(true, "", "dsColUC");
            #endregion
            #region Columna Base Line
            // Add a column for Ubicación Destino
            oColumn = oColumns.Add("colBasLin", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "B.L.";
            oColumn.Width = 80;
            oColumn.Visible = false;

            // bind the text edit item to the defined used data source
            oColumn.DataBind.SetBound(true, "", "dsBasLin");
            #endregion
            #endregion

            #region CheckBox
            #region Dar Por Cerrado
            BaseLeft = 240;
            BaseTop = 90;
            oItem = oForm.Items.Add("chkDPC", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 19;
            oCheckBox = ((SAPbouiCOM.CheckBox)(oItem.Specific));
            oCheckBox.Caption = "Cerrar Orden";
            // binding the Check box with a data source
            oCheckBox.DataBind.SetBound(true, "", "dsDPC");
            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #endregion
        }

        public void FrmProcesos_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            csUtilidades Utilidades = new csUtilidades();
            SAPbouiCOM.EditText oEditText;
            SAPbouiCOM.CheckBox oCheckBox;
            BubbleEvent = true;

            if (FormUID == "FrmProcesos")
            {
                //oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                SAPbouiCOM.Matrix oMatrix;

                switch (pVal.EventType)
                {
                    //case BoEventTypes.et_KEY_DOWN:
                    //    #region Ubicación Artículo
                    //    if (pVal.ItemUID == "txtUbiArt" & pVal.BeforeAction == false & pVal.ActionSuccess == true)
                    //    {
                    //        string a;
                    //        char b;
                    //        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtUbiArt").Specific;
                    //        b = (char)pVal.CharPressed;
                    //        a = (b).ToString().ToUpper();
                    //        pVal.CharPressed = a;
                    //        oEditText.String = oEditText.String.ToUpper();
                    //    }
                    //    #endregion
                    //    break;
                    case BoEventTypes.et_CLICK:
                        #region Procesar Orden
                        if (pVal.ItemUID == "btnProc" && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                        {
                            try
                            {
                                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                                csVariablesGlobales.oCompany.StartTransaction();
                                SAPbobsCOM.Documents oInvGenEntry;
                                SAPbobsCOM.Document_Lines oInvGenEntryLineas;
                                SAPbobsCOM.Documents oInvGenExit;
                                SAPbobsCOM.Document_Lines oInvGenExitLineas;
                                SAPbobsCOM.ProductionOrders oProductionOrders;
                                int RetValSal = 1;
                                int RetValEnt = 1;
                                oInvGenExit = (SAPbobsCOM.Documents)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit));
                                oInvGenEntry = (SAPbobsCOM.Documents)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry));
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matDet").Specific;
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtTipo").Specific;
                                if (ValidarProceso())
                                {
                                    if (oEditText.String == "Desmontaje")
                                    {
                                        #region Desmontaje
                                        #region Salida Artículo
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecPed").Specific;
                                        oInvGenExit.DocDate = Convert.ToDateTime(oEditText.String);
                                        oInvGenExitLineas = oInvGenExit.Lines;
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtCant").Specific;
                                        oInvGenExitLineas.Quantity = Convert.ToDouble(oEditText.String);
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmArt").Specific;
                                        oInvGenExitLineas.WarehouseCode = oEditText.String;
                                        //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtNumOrd").Specific;
                                        //oInvGenExitLineas.BaseEntry = Convert.ToInt32(oEditText.String);
                                        oInvGenExitLineas.BaseEntry = Convert.ToInt32(csVariablesGlobales.NumeroProduccion);
                                        oInvGenExitLineas.BaseType = 202;
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtArt").Specific;
                                        oInvGenExitLineas.Price = Convert.ToDouble(csUtilidades.DameValor("OITM", "LstEvlPric", "ItemCode ='" + oEditText.String + "'"));
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtUbiArt").Specific;
                                        oInvGenExitLineas.BatchNumbers.BatchNumber = oEditText.String;
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtCant").Specific;
                                        oInvGenExitLineas.BatchNumbers.Quantity = Convert.ToDouble(oEditText.String);
                                        oInvGenExitLineas.BatchNumbers.Add();
                                        RetValSal = oInvGenExit.Add();
                                        if (RetValSal != 0)
                                        {
                                            csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                                            csVariablesGlobales.oCompany.StartTransaction();
                                        }
                                        #endregion
                                        #region Entrada Componentes
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecPed").Specific;
                                        oInvGenEntry.DocDate = Convert.ToDateTime(oEditText.String);
                                        oInvGenEntryLineas = oInvGenEntry.Lines;
                                        for (int i = 1; i <= oMatrix.RowCount; i++)
                                        {
                                            if (i >= 2)
                                            {
                                                oInvGenEntryLineas.Add();
                                            }
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colCant").Cells.Item(i).Specific;
                                            oInvGenEntryLineas.Quantity = Convert.ToDouble(oEditText.String);
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colAlmCom").Cells.Item(i).Specific;
                                            oInvGenEntryLineas.WarehouseCode = oEditText.String;
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colBasLin").Cells.Item(i).Specific;
                                            oInvGenEntryLineas.BaseLine = Convert.ToInt32(oEditText.String);
                                            //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtNumOrd").Specific;
                                            //oInvGenEntryLineas.BaseEntry = Convert.ToInt32(oEditText.String);
                                            oInvGenEntryLineas.BaseEntry = Convert.ToInt32(csVariablesGlobales.NumeroProduccion);
                                            oInvGenEntryLineas.BaseType = 202;
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colArtic").Cells.Item(i).Specific;
                                            oInvGenEntryLineas.Price = Convert.ToDouble(csUtilidades.DameValor("OITM", "LstEvlPric", "ItemCode ='" + oEditText.String + "'"));
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colUbiCom").Cells.Item(i).Specific;
                                            oInvGenEntryLineas.BatchNumbers.BatchNumber = oEditText.String;
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colCant").Cells.Item(i).Specific;
                                            oInvGenEntryLineas.BatchNumbers.Quantity = Convert.ToDouble(oEditText.String);
                                            oInvGenEntryLineas.BatchNumbers.Add();
                                        }
                                        RetValEnt = oInvGenEntry.Add();
                                        if (RetValEnt != 0)
                                        {
                                            csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                                        }
                                        #endregion
                                        #endregion
                                    }
                                    if (oEditText.String == "Estándar")
                                    {
                                        #region Montaje
                                        #region Salida Componentes
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecPed").Specific;
                                        oInvGenExit.DocDate = Convert.ToDateTime(oEditText.String);
                                        oInvGenExitLineas = oInvGenExit.Lines;
                                        for (int i = 1; i <= oMatrix.RowCount; i++)
                                        {
                                            if (i >= 2)
                                            {
                                                oInvGenExitLineas.Add();
                                            }
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colCant").Cells.Item(i).Specific;
                                            oInvGenExitLineas.Quantity = Convert.ToDouble(oEditText.String);
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colArtic").Cells.Item(i).Specific;
                                            oInvGenExitLineas.Price = Convert.ToDouble(csUtilidades.DameValor("OITM", "LstEvlPric", "ItemCode ='" + oEditText.String + "'"));
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colAlmCom").Cells.Item(i).Specific;
                                            oInvGenExitLineas.WarehouseCode = oEditText.String;
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colBasLin").Cells.Item(i).Specific;
                                            oInvGenExitLineas.BaseLine = Convert.ToInt32(oEditText.String);
                                            //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtNumOrd").Specific;
                                            //oInvGenExitLineas.BaseEntry = Convert.ToInt32(oEditText.String);
                                            oInvGenExitLineas.BaseEntry = Convert.ToInt32(csVariablesGlobales.NumeroProduccion);
                                            oInvGenExitLineas.BaseType = 202;
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colUbiCom").Cells.Item(i).Specific;
                                            oInvGenExitLineas.BatchNumbers.BatchNumber = oEditText.String;
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colCant").Cells.Item(i).Specific;
                                            oInvGenExitLineas.BatchNumbers.Quantity = Convert.ToDouble(oEditText.String);
                                            oInvGenExitLineas.BatchNumbers.Add();
                                        }
                                        RetValSal = oInvGenExit.Add();
                                        if (RetValSal != 0)
                                        {
                                            csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                                            csVariablesGlobales.oCompany.StartTransaction();
                                        }
                                        #endregion
                                        #region Entrada Artículo
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecPed").Specific;
                                        oInvGenEntry.DocDate = Convert.ToDateTime(oEditText.String);
                                        oInvGenEntryLineas = oInvGenEntry.Lines;
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtArt").Specific;
                                        oInvGenEntryLineas.Price = Convert.ToDouble(csUtilidades.DameValor("OITM", "LstEvlPric", "ItemCode ='" + oEditText.String + "'"));
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtCant").Specific;
                                        oInvGenEntryLineas.Quantity = Convert.ToDouble(oEditText.String);
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmArt").Specific;
                                        oInvGenEntryLineas.WarehouseCode = oEditText.String;
                                        //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtNumOrd").Specific;
                                        //oInvGenEntryLineas.BaseEntry = Convert.ToInt32(oEditText.String);
                                        oInvGenEntryLineas.BaseEntry = Convert.ToInt32(csVariablesGlobales.NumeroProduccion);
                                        oInvGenEntryLineas.BaseType = 202;
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtUbiArt").Specific;
                                        oInvGenEntryLineas.BatchNumbers.BatchNumber = oEditText.String;
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtCant").Specific;
                                        oInvGenEntryLineas.BatchNumbers.Quantity = Convert.ToDouble(oEditText.String);
                                        oInvGenEntryLineas.BatchNumbers.Add();
                                        RetValEnt = oInvGenEntry.Add();
                                        if (RetValEnt != 0)
                                        {
                                            csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                                        }
                                        #endregion
                                        #endregion
                                    }
                                    if (RetValEnt == 0 && RetValSal == 0)
                                    {
                                        csVariablesGlobales.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                                        csVariablesGlobales.SboApp.MessageBox("El proceso se ha realizado correctamente", 1, "", "", "");
                                        oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkDPC").Specific;
                                        if (oCheckBox.Checked)
                                        {
                                            oProductionOrders = (SAPbobsCOM.ProductionOrders)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders));
                                            //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtNumOrd").Specific;
                                            //oProductionOrders.GetByKey(Convert.ToInt32(oEditText.String));
                                            oProductionOrders.GetByKey(Convert.ToInt32(csVariablesGlobales.NumeroProduccion));
                                            oProductionOrders.ProductionOrderStatus = BoProductionOrderStatusEnum.boposClosed;
                                            oProductionOrders.Update();
                                            csVariablesGlobales.SboApp.MessageBox("La orden de producción se ha dado por cerrada", 1, "", "", "");
                                        }
                                        #region Limpiar Formulario
                                        csVariablesGlobales.NumeroOrden = "";
                                        csVariablesGlobales.NumeroProduccion = "";
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtNumOrd").Specific;
                                        oEditText.String = "";
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecPed").Specific;
                                        oEditText.String = "";
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtCant").Specific;
                                        oEditText.String = "";
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmArt").Specific;
                                        oEditText.String = "";
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtUbiArt").Specific;
                                        oEditText.String = "";
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtTipo").Specific;
                                        oEditText.String = "";
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtArt").Specific;
                                        oEditText.String = "";
                                        oMatrix.Clear();
                                        oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkDPC").Specific;
                                        oCheckBox.Checked = true;
                                        #endregion
                                    }
                                    else
                                    {
                                        if (csVariablesGlobales.oCompany.InTransaction)
                                        {
                                            csVariablesGlobales.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                        }
                                    }
                                }
                                else
                                {
                                    if (csVariablesGlobales.oCompany.InTransaction)
                                    {
                                        csVariablesGlobales.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    }
                                }
                            }
                            catch
                            {
                                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                                if (csVariablesGlobales.oCompany.InTransaction)
                                {
                                    csVariablesGlobales.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                }
                            }
                        }

                        #endregion
                        break;
                    //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                        //System.Windows.Forms.Application.Exit();
                        //break;

                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                        string sCFL_ID = null;
                        sCFL_ID = oCFLEvento.ChooseFromListUID;
                        SAPbouiCOM.ChooseFromList oCFL = null;
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                        if (oCFLEvento.BeforeAction == false && oCFLEvento.ActionSuccess == true)
                        {
                            SAPbouiCOM.DataTable oDataTable = null;
                            oDataTable = oCFLEvento.SelectedObjects;
                            try
                            {
                                if (pVal.ItemUID == "txtNumOrd")
                                {
                                    csVariablesGlobales.NumeroProduccion = System.Convert.ToString(oDataTable.GetValue(0, 0));
                                    oForm.DataSources.UserDataSources.Item("dsNumOrd").ValueEx = System.Convert.ToString(oDataTable.GetValue(1, 0));
                                }
                            }
                            catch
                            {
                                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                            }

                        }
                        break;
                    case BoEventTypes.et_LOST_FOCUS:
                        oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                        #region Ubicación Artículo
                        if (pVal.ItemUID == "txtUbiArt" & pVal.BeforeAction == false & pVal.ActionSuccess == true)
                        {
                            SAPbouiCOM.EditText oTxt;
                            string Ubicacion;
                            string Almacen;
                            oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlm" + pVal.ItemUID.Substring(pVal.ItemUID.Length - 3, 3)).Specific;
                            Almacen = oTxt.Value;
                            oTxt = (SAPbouiCOM.EditText)oForm.Items.Item(pVal.ItemUID).Specific;
                            if (!csUtilidades.TodasMayusculas(oTxt.String))
                            {
                                oTxt.String = oTxt.String.ToUpper();
                            }
                            Ubicacion = oTxt.Value;
                            if (Ubicacion == "")
                            {
                                break;
                            }

                            if (csUtilidades.DameValor("[@SIA_UBIC]", "U_cod_ubic", "U_cod_ubic='" + Ubicacion + "' AND U_almacen='" + Almacen + "'") == "")
                            {
                                csVariablesGlobales.SboApp.MessageBox("No existe ninguna ubicación con ese código para ese almacén", 1, "Ok", "", "");
                                oTxt.Value = "";
                                oTxt.Active = true;
                                break;
                            }
                        }
                        #endregion
                        #region Número Orden
                        if (pVal.ItemUID == "txtNumOrd" & pVal.BeforeAction == false & pVal.ActionSuccess == true)
                        {
                            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtNumOrd").Specific;
                            if (oEditText.String != "" && oEditText.String != csVariablesGlobales.NumeroOrden)
                            {
                                SAPbobsCOM.Items oItems;
                                SAPbobsCOM.ProductionOrders oProductionOrders;
                                SAPbobsCOM.ProductionOrders_Lines oProductionOrdersLines;
                                oItems = (SAPbobsCOM.Items)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.oItems);
                                oProductionOrders = (SAPbobsCOM.ProductionOrders)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);
                                //oProductionOrders.GetByKey(Convert.ToInt32(oEditText.String));
                                oProductionOrders.GetByKey(Convert.ToInt32(csVariablesGlobales.NumeroProduccion));
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtCant").Specific;
                                oEditText.String = oProductionOrders.PlannedQuantity.ToString();
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecPed").Specific;
                                oEditText.String = oProductionOrders.DueDate.ToShortDateString();
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtTipo").Specific;
                                switch (oProductionOrders.ProductionOrderType.ToString())
                                {
                                    case "bopotDisassembly":
                                        oEditText.String = "Desmontaje";
                                        break;
                                    case "bopotSpecial":
                                        oEditText.String = "Especial";
                                        break;
                                    case "bopotStandard":
                                        oEditText.String = "Estándar";
                                        break;
                                }
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmArt").Specific;
                                oEditText.String = oProductionOrders.Warehouse.ToString();
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtArt").Specific;
                                oEditText.String = oProductionOrders.ItemNo.ToString();
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matDet").Specific;
                                #region UserDataSource - Los usaré para pasar valores a la matriz
                                oUserDataSourceArticulo = oForm.DataSources.UserDataSources.Item("dsColArt");
                                oUserDataSourceArticuloDescripcion = oForm.DataSources.UserDataSources.Item("dsColAD");
                                oUserDataSourceCantidad = oForm.DataSources.UserDataSources.Item("dsColCan");
                                oUserDataSourceUbicacionComponente = oForm.DataSources.UserDataSources.Item("dsColUC");
                                oUserDataSourceAlmacenComponente = oForm.DataSources.UserDataSources.Item("dsColAC");
                                oUserDataSourceBaseLine = oForm.DataSources.UserDataSources.Item("dsBasLin");
                                #endregion
                                oProductionOrdersLines = oProductionOrders.Lines;
                                oMatrix.Clear();
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtUbiArt").Specific;
                                oEditText.String = "";
                                for (int i = 0; i < oProductionOrders.Lines.Count; i++)
                                {
                                    oProductionOrdersLines.SetCurrentLine(i);
                                    oUserDataSourceArticulo.ValueEx = oProductionOrdersLines.ItemNo.ToString();
                                    oUserDataSourceCantidad.ValueEx = oProductionOrdersLines.PlannedQuantity.ToString();
                                    oUserDataSourceAlmacenComponente.ValueEx = oProductionOrdersLines.Warehouse.ToString();
                                    oItems.GetByKey(oProductionOrdersLines.ItemNo.ToString());
                                    oUserDataSourceArticuloDescripcion.ValueEx = oItems.ItemName.ToString();
                                    oUserDataSourceBaseLine.ValueEx = oProductionOrdersLines.LineNumber.ToString();
                                    oMatrix.AddRow(1, oMatrix.RowCount + 1);
                                }
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtNumOrd").Specific;
                                csVariablesGlobales.NumeroOrden = oEditText.String;
                            }
                            else
                            {
                                if (oEditText.String == "")
                                {
                                    #region Limpiar Formulario
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtCant").Specific;
                                    oEditText.String = "";
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecPed").Specific;
                                    oEditText.String = "";
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtTipo").Specific;
                                    oEditText.String = "";
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmArt").Specific;
                                    oEditText.String = "";
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtArt").Specific;
                                    oEditText.String = "";
                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matDet").Specific;
                                    oMatrix.Clear();
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtUbiArt").Specific;
                                    oEditText.String = "";
                                    oEditText.Active = true;
                                    csVariablesGlobales.NumeroOrden = "";
                                    #endregion
                                }
                            }
                        }
                        #endregion
                        #region Ubicación Componentes
                        if (pVal.ColUID == "colUbiCom" & pVal.BeforeAction == false & pVal.ActionSuccess == true)
                        {
                            //SAPbouiCOM.EditText oTxt;
                            string Ubicacion;
                            string Almacen;
                            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matDet").Specific;
                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colAlmCom").Cells.Item(pVal.Row).Specific;
                            //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlm" + pVal.ItemUID.Substring(pVal.ItemUID.Length - 3, 3)).Specific;
                            Almacen = oEditText.Value;
                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific;
                            if (!csUtilidades.TodasMayusculas(oEditText.String))
                            {
                                oEditText.String = oEditText.String.ToUpper();
                            }
                            Ubicacion = oEditText.Value;
                            if (Ubicacion == "")
                            {
                                break;
                            }

                            if (csUtilidades.DameValor("[@SIA_UBIC]", "U_cod_ubic", "U_cod_ubic='" + Ubicacion + "' AND U_almacen='" + Almacen + "'") == "")
                            {
                                csVariablesGlobales.SboApp.MessageBox("No existe ninguna ubicación con ese código para ese almacén", 1, "Ok", "", "");
                                oEditText.Value = "";
                                break;
                            }
                        }
                        #endregion
                        break;


                    case BoEventTypes.et_GOT_FOCUS:
                        oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                        if (pVal.ItemUID == "txtNumOrd" & pVal.BeforeAction == false & pVal.ActionSuccess == true)
                        {
                            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtNumOrd").Specific;
                            csVariablesGlobales.NumeroOrden = oEditText.String;
                        }
                        break;
                }
            }
        }

        private void FrmProcesos_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
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

                //  Adding CFL para Orden Producción
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "202";
                oCFLCreationParams.UniqueID = "CFL1Ord";
                oCFL = oCFLs.Add(oCFLCreationParams);

                oCons = oCFL.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "Status";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "R";
                oCFL.SetConditions(oCons);
            }
            catch
            {
                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
            }
        }

        private bool ValidarProceso()
        {
            try
            {
                csUtilidades Utilidades = new csUtilidades();
                SAPbouiCOM.EditText oTxt = null;
                SAPbouiCOM.Matrix oMatrix = null;
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtNumOrd").Specific;
                if (oTxt.String == "")
                {
                    csVariablesGlobales.SboApp.MessageBox("El Nº de Orden no puede estar vacío", 1, "", "", "");
                    return false;
                }
                string Ubicacion;
                string Almacen;
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmArt").Specific;
                Almacen = oTxt.Value;
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtUbiArt").Specific;
                if (!csUtilidades.TodasMayusculas(oTxt.String))
                {
                    oTxt.String = oTxt.String.ToUpper();
                }
                Ubicacion = oTxt.Value;
                if (Ubicacion == "")
                {
                    csVariablesGlobales.SboApp.MessageBox("La ubicación del artículo no puede estar vacío", 1, "", "", "");
                    return false;
                }
                if (csUtilidades.DameValor("[@SIA_UBIC]", "U_cod_ubic", "U_cod_ubic='" + Ubicacion + "' AND U_almacen='" + Almacen + "'") == "")
                {
                    csVariablesGlobales.SboApp.MessageBox("No existe ninguna ubicación del artículo con ese código para ese almacén", 1, "Ok", "", "");
                    oTxt.Value = "";
                    return false;
                }
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matDet").Specific;
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colAlmCom").Cells.Item(i).Specific;
                    //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlm" + pVal.ItemUID.Substring(pVal.ItemUID.Length - 3, 3)).Specific;
                    Almacen = oTxt.Value;
                    oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colUbiCom").Cells.Item(i).Specific;
                    if (!csUtilidades.TodasMayusculas(oTxt.String))
                    {
                        oTxt.String = oTxt.String.ToUpper();
                    }
                    Ubicacion = oTxt.Value;
                    if (Ubicacion == "")
                    {
                        csVariablesGlobales.SboApp.MessageBox("La ubicación del componente no puede estar vacío", 1, "", "", "");
                        return false;
                    }

                    if (csUtilidades.DameValor("[@SIA_UBIC]", "U_cod_ubic", "U_cod_ubic='" + Ubicacion + "' AND U_almacen='" + Almacen + "'") == "")
                    {
                        csVariablesGlobales.SboApp.MessageBox("No existe ninguna ubicación del componente con ese código para ese almacén", 1, "Ok", "", "");
                        oTxt.Value = "";
                        return false;
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void CargarDatosInicialesDePantalla()
        {
            SAPbouiCOM.CheckBox oCheckBox;
            oForm = csVariablesGlobales.SboApp.Forms.Item("FrmProcesos");
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkDPC").Specific;
            oCheckBox.Checked = true;
        }
    }
}
