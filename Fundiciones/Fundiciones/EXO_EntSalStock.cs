﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Data.SqlClient;

namespace Addon_SIA
{
    class csFrmEntSalStock
    {
        SAPbouiCOM.UserDataSource oUserDataSourceArticulo = null;
        SAPbouiCOM.UserDataSource oUserDataSourceArticuloDescripcion = null;
        SAPbouiCOM.UserDataSource oUserDataSourceCantidad = null;
        SAPbouiCOM.UserDataSource oUserDataSourceUbicacionOrigen = null;
        SAPbouiCOM.UserDataSource oUserDataSourceUbicacionDestino = null;
        SAPbouiCOM.UserDataSource oUserDataSourceAlmacenOrigen = null;
        SAPbouiCOM.UserDataSource oUserDataSourceAlmacenDestino = null;
        private SAPbouiCOM.Form oForm;

        public void CargarFormulario()
        {
            CrearFormularioEntSalStock();
            oForm.Visible = true;
            //csUtilidades Utilidades = new csUtilidades();
            csUtilidades.SaveAsXml(oForm, "FrmEntSalStock.xml", "");
            // events handled by SBO_Application_ItemEvent
            csVariablesGlobales.SboApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(FrmEntSalStock_ItemEvent);
            // events handled by SBO_Application_AppEvent
            csVariablesGlobales.SboApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(FrmEntSalStock_AppEvent);
        }

        private void CrearFormularioEntSalStock()
        {
            int BaseLeft = 0;
            int BaseTop = 0;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;
            SAPbouiCOM.ComboBox oComboBox = null;
            SAPbouiCOM.StaticText oStaticText = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.Columns oColumns = null;
            SAPbouiCOM.Column oColumn = null;

            #region Formulario
            SAPbouiCOM.FormCreationParams oCreationParams = null;
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "FrmEntSalStock";
            oCreationParams.FormType = "SIASL10001";

            oForm = csVariablesGlobales.SboApp.Forms.AddEx(oCreationParams);

            // set the form properties
            oForm.Title = "Entradas Salidas Stock";
            oForm.Left = 300;
            oForm.ClientWidth = 500;
            oForm.Top = 100;
            oForm.ClientHeight = 400;
            oForm.EnableMenu("1293", true);
            #endregion

            #region Data Sources
            //  Adding a User Data Source
            oForm.DataSources.UserDataSources.Add("dsColArt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsSerEnt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("dsSerSal", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("dsLisPre", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            oForm.DataSources.UserDataSources.Add("dsCant", SAPbouiCOM.BoDataType.dt_QUANTITY, 254);
            oForm.DataSources.UserDataSources.Add("dsFecCon", SAPbouiCOM.BoDataType.dt_DATE, 254);
            oForm.DataSources.UserDataSources.Add("dsFecDoc", SAPbouiCOM.BoDataType.dt_DATE, 254);
            oForm.DataSources.UserDataSources.Add("dsRef2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsUbiOri", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            oForm.DataSources.UserDataSources.Add("dsUbiDes", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            oForm.DataSources.UserDataSources.Add("dsArt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsArtDes", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsColUO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            oForm.DataSources.UserDataSources.Add("dsColUD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            oForm.DataSources.UserDataSources.Add("dsColAD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsColCan", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsArtic", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsAlmOri", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsAlmDes", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsColAlO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            oForm.DataSources.UserDataSources.Add("dsColAlD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            #endregion

            #region Marcos
            #region Traspaso
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
            oItem.Height = 80;

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblTrasp", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 50;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Traspaso";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Detalle
            BaseLeft = 5;
            BaseTop = 105;

            //****************************
            // Adding a Rectangle Item
            // for cosmetic purposes only
            //****************************
            oItem = oForm.Items.Add("Rect2", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = BaseLeft;
            oItem.Width = 480;
            oItem.Top = BaseTop;
            oItem.Height = 250;

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblLin", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 80;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Líneas Traspaso";

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
            oItem.Top = 365;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";
            #endregion
            #region Botón Traspaso Ubicación
            // /**********************
            // Adding an Ok button
            //*********************

            // We get automatic event handling for
            // the Ok and Cancel Buttons by setting
            // their UIDs to 1 and 2 respectively

            oItem = oForm.Items.Add("btnTrasUbi", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 80;
            oItem.Width = 105;
            oItem.Top = 365;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Traspasar Ubicación";
            #endregion
            #region Botón Añadir
            // /**********************
            // Adding an Ok button
            //*********************

            // We get automatic event handling for
            // the Ok and Cancel Buttons by setting
            // their UIDs to 1 and 2 respectively

            oItem = oForm.Items.Add("btnAnadir", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 360;
            oItem.Width = 100;
            oItem.Top = 180;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Añadir";
            #endregion
            #endregion

            #region Combos
            #region Combo Series Entrada
            //*************************
            // Adding a Combo Box item
            //*************************
            BaseLeft = 10;
            BaseTop = 30;

            oItem = oForm.Items.Add("cmbSerEnt", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.DisplayDesc = true;

            oComboBox = ((SAPbouiCOM.ComboBox)(oItem.Specific));

            // bind the Combo Box item to the defined used data source
            oComboBox.DataBind.SetBound(true, "", "dsSerEnt");

            //oComboBox.ValidValues.LoadSeries("59", SAPbouiCOM.BoSeriesMode.sf_Add); //Entradas
            //oComboBox.Select("0", BoSearchKey.psk_Index);

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblSerEnt", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "cmbSerEnt";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Serie Entrada";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Combo Series Salida
            //*************************
            // Adding a Combo Box item
            //*************************
            BaseLeft = 240;
            BaseTop = 30;

            oItem = oForm.Items.Add("cmbSerSal", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.DisplayDesc = true;

            oComboBox = ((SAPbouiCOM.ComboBox)(oItem.Specific));

            // bind the Combo Box item to the defined used data source
            oComboBox.DataBind.SetBound(true, "", "dsSerSal");

            //oComboBox.ValidValues.LoadSeries("60", SAPbouiCOM.BoSeriesMode.sf_Add); //Salidas
            //oComboBox.Select("0", BoSearchKey.psk_Index);

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblSerSal", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "cmbSerSal";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Serie Salida";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Combo Lista Precios
            //*************************
            // Adding a Combo Box item
            //*************************
            BaseLeft = 10;
            BaseTop = 70;

            oItem = oForm.Items.Add("cmbLisPre", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.DisplayDesc = true;

            oComboBox = ((SAPbouiCOM.ComboBox)(oItem.Specific));

            //#region Añadir Listas de Precios
            //oComboBox.ValidValues.Add("-2", "Último precio evaluado");
            //oComboBox.ValidValues.Add("-1", "Último precio de compra");
            //System.Data.DataTable RstDato = new System.Data.DataTable();
            //SqlCommand SelectCMD;
            //SelectCMD = new SqlCommand("SELECT ListNum, ListName FROM OPLN", csVariablesGlobales.conAddon);
            //SqlDataAdapter SQlDAGrid = new SqlDataAdapter();
            //SQlDAGrid.SelectCommand = SelectCMD;
            //SQlDAGrid.Fill(RstDato);

            //if (RstDato.Rows.Count > 0)
            //{
            //    for (int i = 0; i < RstDato.Rows.Count; i++)
            //    {
            //        oComboBox.ValidValues.Add(RstDato.Rows[i][0].ToString(), RstDato.Rows[i][1].ToString());
            //    }
            //}
            //#endregion
            //oComboBox.Select("-2", BoSearchKey.psk_ByValue);

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblLisPre", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "cmbLisPre";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Lista de Precios";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #endregion

            #region Labels
            #region Articulo Descripción
            BaseTop = 0;
            BaseLeft = 0;
            oItem = oForm.Items.Add("lblArtDes", SAPbouiCOM.BoFormItemTypes.it_STATIC);
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
            #region Campo Referencia 2
            BaseLeft = 240;
            BaseTop = 70;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtRef2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblRef2";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsRef2");


            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblRef2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtRef2";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Referencia 2";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion Campo Desde IC
            #region Fecha Contable
            BaseLeft = 10;
            BaseTop = 50;

            // Date Picker
            //__________________________________________________________________________________________
            ////*************************
            //// Adding a Text Edit item
            ////*************************

            oItem = oForm.Items.Add("txtFecCon", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            //// bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsFecCon");

            ////**********************************
            //// Adding Static Text item
            ////**********************************

            oItem = oForm.Items.Add("lblFecCon", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtFecCon";
            oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            oStaticText.Caption = "Fecha Contable";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Fecha Documento
            BaseLeft = 240;
            BaseTop = 50;

            // Date Picker
            //__________________________________________________________________________________________
            ////*************************
            //// Adding a Text Edit item
            ////*************************

            oItem = oForm.Items.Add("txtFecDoc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            //// bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsFecDoc");

            ////**********************************
            //// Adding Static Text item
            ////**********************************

            oItem = oForm.Items.Add("lblFecDoc", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtFecDoc";
            oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            oStaticText.Caption = "Fecha Documento";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Articulo
            BaseLeft = 10;
            BaseTop = 120;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtArt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblArt";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsArt");
            oEditText.ChooseFromListUID = "CFL1Art";
            oEditText.ChooseFromListAlias = "ItemCode";

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblArt", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtArt";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Artículo";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Cantidad
            // Calc Picker
            //__________________________________________________________________________________________
            ////*************************
            //// Adding a Text Edit item
            ////*************************
            BaseLeft = 240;
            BaseTop = 120;

            oItem = oForm.Items.Add("txtCant", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

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
            #region Almacén Origen
            BaseLeft = 10;
            BaseTop = 140;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtAlmOri", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblAlmOri";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsAlmOri");
            oEditText.ChooseFromListUID = "CFL1AO";
            oEditText.ChooseFromListAlias = "WhsCode";

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblAlmOri", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtAlmOri";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Almacén Origen";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Almacén Destino
            BaseLeft = 240;
            BaseTop = 140;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtAlmDes", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblAlmDes";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsAlmDes");
            oEditText.ChooseFromListUID = "CFL1AD";
            oEditText.ChooseFromListAlias = "WhsCode";

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblAlmDes", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtAlmDes";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Almacén Destino";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Ubicación Origen
            BaseLeft = 10;
            BaseTop = 160;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtUbiOri", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblUbiOri";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsUbiOri");
            SAPbobsCOM.FormattedSearches oFormattedSearches;
            oFormattedSearches = (SAPbobsCOM.FormattedSearches)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.oFormattedSearches);


            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblUbiOri", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtUbiOri";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Ubicación Origen";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Ubicación Destino
            BaseLeft = 240;
            BaseTop = 160;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtUbiDes", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblUbiDes";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsUbiDes");

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblUbiDes", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 110;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtUbiDes";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Ubicación Destino";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion

            #endregion

            #region Matrix
            BaseTop = 210;
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
            #region Columna Ubicación Origen
            // Add a column for Ubicación Origen
            oColumn = oColumns.Add("colUbiOri", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ubic. Ori.";
            oColumn.Width = 80;
            oColumn.Editable = false;

            // bind the text edit item to the defined used data source
            oColumn.DataBind.SetBound(true, "", "dsColUO");
            #endregion
            #region Columna Ubicación Destino
            // Add a column for Ubicación Destino
            oColumn = oColumns.Add("colUbiDes", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ubic. Des.";
            oColumn.Width = 80;
            oColumn.Editable = false;

            // bind the text edit item to the defined used data source
            oColumn.DataBind.SetBound(true, "", "dsColUD");
            #endregion
            #region Columna Almacén Origen
            // Add a column for Ubicación Origen
            oColumn = oColumns.Add("colAlmOri", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Alm. Ori.";
            oColumn.Width = 80;
            oColumn.Editable = false;

            // bind the text edit item to the defined used data source
            oColumn.DataBind.SetBound(true, "", "dsColAlO");
            #endregion
            #region Columna Almacén Destino
            // Add a column for Ubicación Destino
            oColumn = oColumns.Add("colAlmDes", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Alm. Des.";
            oColumn.Width = 80;
            oColumn.Editable = false;

            // bind the text edit item to the defined used data source
            oColumn.DataBind.SetBound(true, "", "dsColAlD");
            #endregion
            #endregion

            #region UserDataSource - Los usaré para pasar valores a la matriz
            oUserDataSourceArticulo = oForm.DataSources.UserDataSources.Item("dsColArt");
            oUserDataSourceArticuloDescripcion = oForm.DataSources.UserDataSources.Item("dsColAD");
            oUserDataSourceCantidad = oForm.DataSources.UserDataSources.Item("dsColCan");
            oUserDataSourceUbicacionOrigen = oForm.DataSources.UserDataSources.Item("dsColUO");
            oUserDataSourceUbicacionDestino = oForm.DataSources.UserDataSources.Item("dsColUD");
            oUserDataSourceAlmacenOrigen = oForm.DataSources.UserDataSources.Item("dsColAlO");
            oUserDataSourceAlmacenDestino = oForm.DataSources.UserDataSources.Item("dsColAlD");
            #endregion
        }

        public void FrmEntSalStock_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            csUtilidades Utilidades = new csUtilidades();
            BubbleEvent = true;

            if (pVal.FormUID == "FrmEntSalStock")
            {

                //oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                        #region Botón Traspasar Ubicación
                        if (pVal.ItemUID == "btnTrasUbi" && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                        {
                            SAPbobsCOM.Documents oInvGenEntry;
                            SAPbobsCOM.Document_Lines oInvGenEntryLineas;
                            SAPbobsCOM.Documents oInvGenExit;
                            SAPbobsCOM.Document_Lines oInvGenExitLineas;
                            SAPbouiCOM.EditText oEditText;
                            SAPbouiCOM.ComboBox oComboBox;
                            SAPbouiCOM.Matrix oMatrix;
                            int RetValEnt = 1;
                            int RetValSal = 1;
                            try
                            {
                                csVariablesGlobales.oCompany.StartTransaction();
                                if (ValidarTraspaso())
                                {
                                    int ValorInt;
                                    string ValorStr;
                                    double ValorDbl;

                                    #region Salida de Stock
                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matDet").Specific;
                                    oInvGenExit = (SAPbobsCOM.Documents)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit));

                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecCon").Specific;
                                    oInvGenExit.DocDate = Convert.ToDateTime(oEditText.String);

                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecDoc").Specific;
                                    oInvGenExit.TaxDate = Convert.ToDateTime(oEditText.String);

                                    oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbLisPre").Specific;
                                    ValorInt = Convert.ToInt32(oComboBox.Selected.Value);

                                    oInvGenExit.PaymentGroupCode = ValorInt;

                                    //oInvGenExit.DocCurrency = "EUR";
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtRef2").Specific;
                                    oInvGenExit.Reference2 = oEditText.String;

                                    oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbSerSal").Specific;
                                    ValorInt = Convert.ToInt32(oComboBox.Selected.Value);
                                    oInvGenExit.Series = ValorInt;

                                    oInvGenExitLineas = oInvGenExit.Lines;
                                    for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                                    {
                                        if (i >= 2)
                                        {
                                            oInvGenExitLineas.Add();
                                        }
                                        
                                        ValorStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colArtic").Cells.Item(i).Specific).String;
                                        oInvGenExitLineas.ItemCode = ValorStr;
                                        oInvGenExitLineas.Price = Convert.ToDouble(csUtilidades.DameValor("OITM", "LstEvlPric", "ItemCode ='" + ValorStr + "'"));

                                                                                                                      
                                        ValorStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colDesArt").Cells.Item(i).Specific).String;
                                        oInvGenExitLineas.ItemDescription = oEditText.String;

                                        ValorStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colAlmOri").Cells.Item(i).Specific).String;
                                        oInvGenExitLineas.WarehouseCode = ValorStr;
                                        
                                        ValorDbl = csUtilidades.ConvertirCantidad(((SAPbouiCOM.EditText)oMatrix.Columns.Item("colCant").Cells.Item(i).Specific).String);
                                        oInvGenExitLineas.Quantity = ValorDbl;
                                        oInvGenExitLineas.BatchNumbers.Quantity = ValorDbl;

                                        ValorStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colUbiOri").Cells.Item(i).Specific).String;
                                        oInvGenExitLineas.BatchNumbers.BatchNumber = ValorStr;
                                        oInvGenExitLineas.BatchNumbers.Add();
                                    }
                                    RetValSal = oInvGenExit.Add();
                                    if (RetValSal != 0)
                                    {
                                        csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                                        csVariablesGlobales.oCompany.StartTransaction();
                                    }
                                    #endregion

                                    #region Entrada de Stock

                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matDet").Specific;
                                    oInvGenEntry = (SAPbobsCOM.Documents)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry));

                                    
                                    oInvGenEntry.DocDate = Convert.ToDateTime(((SAPbouiCOM.EditText)oForm.Items.Item("txtFecCon").Specific).String);
                                    oInvGenEntry.TaxDate = Convert.ToDateTime(((SAPbouiCOM.EditText)oForm.Items.Item("txtFecCon").Specific).String);                                    

                                    oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbLisPre").Specific;
                                    ValorInt = Convert.ToInt32(oComboBox.Selected.Value);
                                    oInvGenEntry.PaymentGroupCode = ValorInt;
                                    
                                    oInvGenEntry.Reference2 = ((SAPbouiCOM.EditText)oForm.Items.Item("txtRef2").Specific).String;
                                    oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbSerEnt").Specific;
                                    ValorInt = Convert.ToInt32(oComboBox.Selected.Value);
                                    oInvGenEntry.Series = ValorInt;
                                    oInvGenEntryLineas = oInvGenEntry.Lines;
                                    for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                                    {
                                        if (i >= 2)
                                        {
                                            oInvGenEntryLineas.Add();
                                        }                                        
                                        ValorStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colArtic").Cells.Item(i).Specific).String;
                                        oInvGenEntryLineas.ItemCode = ValorStr;
                                        oInvGenEntryLineas.Price = Convert.ToDouble(csUtilidades.DameValor("OITM", "LstEvlPric", "ItemCode ='" + ValorStr + "'"));

                                        ValorStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colDesArt").Cells.Item(i).Specific).String;
                                        oInvGenEntryLineas.ItemDescription = ValorStr;

                                        ValorStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colAlmDes").Cells.Item(i).Specific).String;
                                        oInvGenEntryLineas.WarehouseCode = ValorStr;
                                        
                                        ValorDbl = csUtilidades.ConvertirCantidad(((SAPbouiCOM.EditText)oMatrix.Columns.Item("colCant").Cells.Item(i).Specific).String);                                        
                                        oInvGenEntryLineas.Quantity = ValorDbl;
                                        oInvGenEntryLineas.BatchNumbers.Quantity = ValorDbl;

                                        ValorStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colUbiDes").Cells.Item(i).Specific).String;
                                        oInvGenEntryLineas.BatchNumbers.BatchNumber = ValorStr;
                                        
                                        oInvGenEntryLineas.BatchNumbers.Add();
                                    }
                                    RetValEnt = oInvGenEntry.Add();
                                    if (RetValEnt != 0)
                                    {
                                        csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                                    }
                                    #endregion
                                    #region Limpiar Formulario
                                    if (RetValSal == 0 && RetValEnt == 0)
                                    {
                                        //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecCon").Specific;
                                        //oEditText.String = "";
                                        //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecDoc").Specific;
                                        //oEditText.String = "";
                                        oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtRef2").Specific;
                                        oEditText.String = "";
                                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matDet").Specific;
                                        oMatrix.Clear();
                                        csVariablesGlobales.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                                        csVariablesGlobales.SboApp.MessageBox("El traspaso se ha realizado correctamente", 1, "", "", "");
                                    }
                                    else
                                    {
                                        if (csVariablesGlobales.oCompany.InTransaction)
                                        {
                                            csVariablesGlobales.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    csVariablesGlobales.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
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
                        #region Botón Añadir
                        if (pVal.ItemUID == "btnAnadir" && pVal.ActionSuccess == true && pVal.BeforeAction == false)
                        {
                            try
                            {
                                SAPbouiCOM.Matrix oMatrix;
                                SAPbouiCOM.EditText otxtDato;
                                SAPbouiCOM.StaticText oStaticTextDato;
                                #region UserDataSource - Los usaré para pasar valores a la matriz
                                oUserDataSourceArticulo = oForm.DataSources.UserDataSources.Item("dsColArt");
                                oUserDataSourceArticuloDescripcion = oForm.DataSources.UserDataSources.Item("dsColAD");
                                oUserDataSourceCantidad = oForm.DataSources.UserDataSources.Item("dsColCan");
                                oUserDataSourceUbicacionOrigen = oForm.DataSources.UserDataSources.Item("dsColUO");
                                oUserDataSourceUbicacionDestino = oForm.DataSources.UserDataSources.Item("dsColUD");
                                oUserDataSourceAlmacenOrigen = oForm.DataSources.UserDataSources.Item("dsColAlO");
                                oUserDataSourceAlmacenDestino = oForm.DataSources.UserDataSources.Item("dsColAlD");
                                #endregion
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matDet").Specific;
                                otxtDato = (SAPbouiCOM.EditText)oForm.Items.Item("txtArt").Specific;
                                //if (otxtDato.String != "")
                                if (ValidarInsercionEnMatrix())
                                {
                                    oUserDataSourceArticulo.ValueEx = otxtDato.String;
                                    oStaticTextDato = (SAPbouiCOM.StaticText)oForm.Items.Item("lblArtDes").Specific;
                                    oUserDataSourceArticuloDescripcion.Value = oStaticTextDato.Caption;
                                    otxtDato = (SAPbouiCOM.EditText)oForm.Items.Item("txtCant").Specific;
                                    oUserDataSourceCantidad.Value = otxtDato.String;
                                    otxtDato.String = "0";
                                    otxtDato = (SAPbouiCOM.EditText)oForm.Items.Item("txtUbiOri").Specific;
                                    oUserDataSourceUbicacionOrigen.Value = otxtDato.String;
                                    otxtDato.String = "";
                                    otxtDato = (SAPbouiCOM.EditText)oForm.Items.Item("txtUbiDes").Specific;
                                    oUserDataSourceUbicacionDestino.Value = otxtDato.String;
                                    otxtDato.String = "";
                                    otxtDato = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmOri").Specific;
                                    oUserDataSourceAlmacenOrigen.Value = otxtDato.String;
                                    otxtDato.String = "01";
                                    otxtDato = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmDes").Specific;
                                    oUserDataSourceAlmacenDestino.Value = otxtDato.String;
                                    otxtDato.String = "01";
                                    oMatrix.AddRow(1, oMatrix.RowCount + 1);
                                    otxtDato = (SAPbouiCOM.EditText)oForm.Items.Item("txtArt").Specific;
                                    otxtDato.String = "";
                                    oStaticTextDato = (SAPbouiCOM.StaticText)oForm.Items.Item("lblArtDes").Specific;
                                    oStaticTextDato.Caption = "";
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.ToString());
                                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                            }
                        }
                        #endregion
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                        //System.Windows.Forms.Application.Exit();
                        break;
                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
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
                            string Codigo = null;
                            try
                            {
                                Codigo = System.Convert.ToString(oDataTable.GetValue(0, 0));
                                if (pVal.ItemUID == "txtArt")
                                {
                                    SAPbouiCOM.StaticText oStaticText;
                                    //DescripcionArticulo = System.Convert.ToString(oDataTable.GetValue(1, 0));
                                    oForm.DataSources.UserDataSources.Item("dsArt").ValueEx = Codigo;
                                    //oForm.DataSources.UserDataSources.Item("dsArtDes").ValueEx = DescripcionArticulo;
                                    oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblArtDes").Specific;
                                    oStaticText.Caption = System.Convert.ToString(oDataTable.GetValue(1, 0));
                                }
                                if (pVal.ItemUID == "txtAlmOri")
                                {
                                    oForm.DataSources.UserDataSources.Item("dsAlmOri").ValueEx = System.Convert.ToString(oDataTable.GetValue(0, 0));
                                }
                                if (pVal.ItemUID == "txtAlmDes")
                                {
                                    oForm.DataSources.UserDataSources.Item("dsAlmDes").ValueEx = System.Convert.ToString(oDataTable.GetValue(0, 0));
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
                        if ((pVal.ItemUID == "txtUbiOri" | pVal.ItemUID == "txtUbiDes") & pVal.BeforeAction == false & pVal.ActionSuccess == true)
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
                                break;
                            }

                        }
                        break;
                }
            }
        }

        private void FrmEntSalStock_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
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

                oCFLs = oForm.ChooseFromLists;

                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                //  Adding CFL para Artículo
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "4";
                oCFLCreationParams.UniqueID = "CFL1Art";
                oCFL = oCFLs.Add(oCFLCreationParams);

                //  Adding CFL para Almacén Origen
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "64";
                oCFLCreationParams.UniqueID = "CFL1AO";
                oCFL = oCFLs.Add(oCFLCreationParams);

                //  Adding CFL para Almacén Destino
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "64";
                oCFLCreationParams.UniqueID = "CFL1AD";
                oCFL = oCFLs.Add(oCFLCreationParams);
            }
            catch
            {
                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
            }
        }

        private bool ValidarInsercionEnMatrix()
        {
            try
            {
                csUtilidades Utilidades = new csUtilidades();
                string Ubicacion;
                string Almacen;
                string Articulo;
                double Cantidad;
                string UbicacionColumna;
                string AlmacenColumna;
                string ArticuloColumna;
                SAPbouiCOM.EditText oTxt = null;
                SAPbouiCOM.Matrix oMatrix;
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtArt").Specific;
                if (oTxt.String == "")
                {
                    oTxt.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("El artículo debe existir", 1, "Ok", "", "");
                    return false;
                }
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtCant").Specific;
                if (oTxt.String == "")
                {
                    oTxt.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("La cantidad debe ser mayor que 0", 1, "Ok", "", "");
                    return false;
                }
                if (Convert.ToDouble(oTxt.String) <= 0.0)
                {
                    oTxt.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("La cantidad debe ser mayor que 0", 1, "Ok", "", "");
                    return false;
                }
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmOri").Specific;
                if (oTxt.String == "")
                {
                    oTxt.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("El almacén origen debe existir", 1, "Ok", "", "");
                    return false;
                }
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmDes").Specific;
                if (oTxt.String == "")
                {
                    oTxt.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("El almacén destino debe existir", 1, "Ok", "", "");
                    return false;
                }
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmOri").Specific;
                Almacen = oTxt.Value;
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtUbiOri").Specific;
                if (!csUtilidades.TodasMayusculas(oTxt.String))
                {
                    oTxt.String = oTxt.String.ToUpper();
                }
                Ubicacion = oTxt.Value;
                if (Ubicacion == "")
                {
                    csVariablesGlobales.SboApp.MessageBox("La ubicación origen no puede estar vacío", 1, "", "", "");
                    return false;
                }
                if (csUtilidades.DameValor("[@SIA_UBIC]", "U_cod_ubic", "U_cod_ubic='" + Ubicacion + "' AND U_almacen='" + Almacen + "'") == "")
                {
                    csVariablesGlobales.SboApp.MessageBox("No existe ninguna ubicación origen con ese código para ese almacén", 1, "Ok", "", "");
                    oTxt.Value = "";
                    return false;
                }
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmDes").Specific;
                Almacen = oTxt.Value;
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtUbiDes").Specific;
                if (!csUtilidades.TodasMayusculas(oTxt.String))
                {
                    oTxt.String = oTxt.String.ToUpper();
                }
                Ubicacion = oTxt.Value;
                if (Ubicacion == "")
                {
                    csVariablesGlobales.SboApp.MessageBox("La ubicación destino no puede estar vacío", 1, "", "", "");
                    return false;
                }
                if (csUtilidades.DameValor("[@SIA_UBIC]", "U_cod_ubic", "U_cod_ubic='" + Ubicacion + "' AND U_almacen='" + Almacen + "'") == "")
                {
                    csVariablesGlobales.SboApp.MessageBox("No existe ninguna ubicación destino con ese código para ese almacén", 1, "Ok", "", "");
                    oTxt.Value = "";
                    return false;
                }
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmOri").Specific;
                Almacen = oTxt.Value;
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtUbiOri").Specific;
                Ubicacion = oTxt.Value;
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtArt").Specific;
                Articulo = oTxt.Value;
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtCant").Specific;
                Cantidad = Convert.ToDouble(oTxt.Value.ToString().Replace('.', ','));
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matDet").Specific;
                for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                {
                    oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colArtic").Cells.Item(i).Specific;
                    ArticuloColumna = oTxt.String;
                    oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colAlmOri").Cells.Item(i).Specific;
                    AlmacenColumna = oTxt.String;
                    oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colUbiOri").Cells.Item(i).Specific;
                    UbicacionColumna = oTxt.String;
                    if (Articulo == ArticuloColumna && Almacen == AlmacenColumna && Ubicacion == UbicacionColumna)
                    {
                        oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colCant").Cells.Item(i).Specific;
                        Cantidad = Cantidad + Convert.ToDouble(oTxt.Value);
                    }
                }
                if (csUtilidades.DameValor("OIBT", "Quantity", "ItemCode = '" + Articulo + "' AND BatchNum = '" + Ubicacion + "' AND WhsCode = '" + Almacen + "' AND Quantity >= " + Cantidad) == "")
                {
                    csVariablesGlobales.SboApp.MessageBox("No existe cantidad suficiente en esa ubicación de origen para ese código de artículo en ese almacén", 1, "Ok", "", "");
                    return false;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool ValidarTraspaso()
        {
            try
            {
                SAPbouiCOM.EditText oTxt = null;
                SAPbouiCOM.Matrix oMatrix = null;
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecCon").Specific;
                if (oTxt.String == "")
                {
                    oTxt.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("La fecha contable debe existir", 1, "Ok", "", "");
                    return false;
                }
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecDoc").Specific;
                if (oTxt.String == "")
                {
                    oTxt.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("La fecha de documento debe existir", 1, "Ok", "", "");
                    return false;
                }
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matDet").Specific;
                if (oMatrix.VisualRowCount == 0)
                {
                    csVariablesGlobales.SboApp.MessageBox("No hay ninguna línea que traspasar", 1, "Ok", "", "");
                    return false;
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
            csUtilidades Utilidades = new csUtilidades();
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oComboBox = null;
            string Serie;
            oForm = csVariablesGlobales.SboApp.Forms.Item("FrmEntSalStock");
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecCon").Specific;
            oEditText.String = DateTime.Today.Day.ToString() + "/" + DateTime.Today.Month.ToString() + "/" + DateTime.Today.Year.ToString();
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecDoc").Specific;
            oEditText.String = DateTime.Today.Day.ToString() + "/" + DateTime.Today.Month.ToString() + "/" + DateTime.Today.Year.ToString();
            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbSerEnt").Specific;
            Serie = csUtilidades.DameValor("ONNM a, NNM1 b", "b.SeriesName", "a.ObjectCode = b.ObjectCode AND b.Series = a.DfltSeries AND a.ObjectCode = '59'");
            oComboBox.ValidValues.LoadSeries("59", SAPbouiCOM.BoSeriesMode.sf_Add); //Entradas
            //oComboBox.Select("0", BoSearchKey.psk_Index);
            oComboBox.Select(Serie, BoSearchKey.psk_ByDescription);
            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbSerSal").Specific;
            Serie = csUtilidades.DameValor("ONNM a, NNM1 b", "b.SeriesName", "a.ObjectCode = b.ObjectCode AND b.Series = a.DfltSeries AND a.ObjectCode = '60'");
            oComboBox.ValidValues.LoadSeries("60", SAPbouiCOM.BoSeriesMode.sf_Add); //Salidas
            //oComboBox.Select("0", BoSearchKey.psk_Index);
            oComboBox.Select(Serie, BoSearchKey.psk_ByDescription);
            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbLisPre").Specific;
            #region Añadir Listas de Precios
            oComboBox.ValidValues.Add("-2", "Último precio evaluado");
            oComboBox.ValidValues.Add("-1", "Último precio de compra");
            System.Data.DataTable RstDato = new System.Data.DataTable();
            SqlCommand SelectCMD;
            SelectCMD = new SqlCommand("SELECT ListNum, ListName FROM OPLN", csVariablesGlobales.conAddon);
            SqlDataAdapter SQlDAGrid = new SqlDataAdapter();
            SQlDAGrid.SelectCommand = SelectCMD;
            SQlDAGrid.Fill(RstDato);

            if (RstDato.Rows.Count > 0)
            {
                for (int i = 0; i < RstDato.Rows.Count; i++)
                {
                    oComboBox.ValidValues.Add(RstDato.Rows[i][0].ToString(), RstDato.Rows[i][1].ToString());
                }
            }
            #endregion
            oComboBox.Select("-2", BoSearchKey.psk_ByValue);
            oComboBox.Select("-2", BoSearchKey.psk_ByValue);
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmOri").Specific;
            oEditText.String = "01";
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtAlmDes").Specific;
            oEditText.String = "01";
        }
    }
}
