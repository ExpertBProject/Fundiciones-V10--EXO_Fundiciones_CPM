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
    class csFrmModelo347
    {
        private SAPbouiCOM.Form oForm;

        public void CargarFormulario()
        {
            CrearFormulario347();
            oForm.Visible = true;
            //csUtilidades Utilidades = new csUtilidades();
            csUtilidades.SaveAsXml(oForm, "FrmModelo347.xml", "");
            // events handled by SBO_Application_ItemEvent
            csVariablesGlobales.SboApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(FrmModelo347_ItemEvent);
            // events handled by SBO_Application_AppEvent
            csVariablesGlobales.SboApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(FrmModelo347_AppEvent);
        }

        private void CrearFormulario347()
        {
            int BaseLeft = 0;
            int BaseTop = 0;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;
            SAPbouiCOM.OptionBtn oOptionBtn = null;
            SAPbouiCOM.CheckBox oCheckBox = null;
            SAPbouiCOM.StaticText oStaticText = null;
            SAPbouiCOM.EditText oEditText = null;

            #region Formulario
            SAPbouiCOM.FormCreationParams oCreationParams = null;
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "FrmModelo347";
            oCreationParams.FormType = "SIASL10005";

            oForm = csVariablesGlobales.SboApp.Forms.AddEx(oCreationParams);

            // set the form properties
            oForm.Title = "Modelo 347";
            oForm.Left = 300;
            oForm.ClientWidth = 500;
            oForm.Top = 100;
            oForm.ClientHeight = 460;
            //oForm.EnableMenu("519", true);
            //oForm.EnableMenu("520", true);
            #endregion

            #region Data Sources
            //  Adding a User Data Source
            oForm.DataSources.UserDataSources.Add("dsDesdeIC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsHastaIC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsDeud", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("dsAcre", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("dsDesdeFec", SAPbouiCOM.BoDataType.dt_DATE, 254);
            oForm.DataSources.UserDataSources.Add("dsHastaFec", SAPbouiCOM.BoDataType.dt_DATE, 254);
            oForm.DataSources.UserDataSources.Add("dsInExt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("dsInICRet", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("dsMaxImp", SAPbouiCOM.BoDataType.dt_PRICE, 254);
            oForm.DataSources.UserDataSources.Add("dsEjer", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("dsSelFich", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsFecCar", SAPbouiCOM.BoDataType.dt_DATE, 254);
            oForm.DataSources.UserDataSources.Add("dsFir", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsOptDes", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
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
            oItem.Height = 60;

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
            #region Rango Fechas
            BaseLeft = 5;
            BaseTop = 85;

            //****************************
            // Adding a Rectangle Item
            // for cosmetic purposes only
            //****************************
            oItem = oForm.Items.Add("Rect2", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = BaseLeft;
            oItem.Width = 235;
            oItem.Top = BaseTop;
            oItem.Height = 100;

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblRgFec", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 70;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Rango Fecha";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Opciones
            BaseLeft = 250;
            BaseTop = 85;

            //****************************
            // Adding a Rectangle Item
            // for cosmetic purposes only
            //****************************
            oItem = oForm.Items.Add("Rect3", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = BaseLeft;
            oItem.Width = 235;
            oItem.Top = BaseTop;
            oItem.Height = 100;

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblOpc", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 50;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Opciones";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Opciones Exportación
            BaseLeft = 5;
            BaseTop = 205;

            //****************************
            // Adding a Rectangle Item
            // for cosmetic purposes only
            //****************************
            oItem = oForm.Items.Add("Rect4", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = BaseLeft;
            oItem.Width = 480;
            oItem.Top = BaseTop;
            oItem.Height = 120;

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblOpcExp", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 110;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Opciones Exportación";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Destino Impreso
            BaseLeft = 5;
            BaseTop = 345;

            //****************************
            // Adding a Rectangle Item
            // for cosmetic purposes only
            //****************************
            oItem = oForm.Items.Add("Rect5", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = BaseLeft;
            oItem.Width = 480;
            oItem.Top = BaseTop;
            oItem.Height = 60;

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblDesImp", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 110;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Destino Impreso";

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
            #region btnGenerar
            // /**********************
            // Adding an Ok button
            //*********************

            oItem = oForm.Items.Add("btnGen", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5;
            oItem.Width = 65;
            oItem.Top = 425;
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
            oItem.Top = 425;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";
            #endregion
            #region btnInforme
            //************************
            // Adding a Cancel button
            //***********************

            oItem = oForm.Items.Add("btnInf", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 165;
            oItem.Width = 65;
            oItem.Top = 425;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Informe";
            #endregion
            #region btnCarta
            //************************
            // Adding a Cancel button
            //***********************

            oItem = oForm.Items.Add("btnCar", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 235;
            oItem.Width = 65;
            oItem.Top = 425;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Carta";
            #endregion
            #region Selecciona Fichero
            //************************
            // Adding a Cancel button
            //***********************

            oItem = oForm.Items.Add("btnSelFich", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 440;
            oItem.Width = 20;
            oItem.Top = 235;
            oItem.Height = 19;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "...";
            #endregion
            #endregion

            #region Campos
            #region Campo Desde IC
            BaseLeft = 250;
            BaseTop = 30;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtDesdeIC", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblDesdeIC";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsDesdeIC");
            //oEditText.String = "Hasta IC";
            oEditText.ChooseFromListUID = "CFL1DesIC";
            oEditText.ChooseFromListAlias = "CardCode";

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblDesdeIC", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtDesdeIC";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Desde IC";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion Campo Desde IC
            #region Campo Hasta IC
            BaseLeft = 250;
            BaseTop = 50;

            //*************************
            // Adding a Text Edit item
            //*************************
            oItem = oForm.Items.Add("txtHastaIC", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "lblHastaIC";
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsHastaIC");
            //oEditText.String = "Hasta IC";
            oEditText.ChooseFromListUID = "CFL1HasIC";
            oEditText.ChooseFromListAlias = "CardCode";

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblHastaIC", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.LinkTo = "txtHastaIC";
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Hasta IC";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion Campo Desde IC
            #region Desde Fecha
            BaseLeft = 10;
            BaseTop = 110;

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
            BaseTop = 130;

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
            #region Maximo Importe
            // Calc Picker
            //__________________________________________________________________________________________
            ////*************************
            //// Adding a Text Edit item
            ////*************************
            BaseLeft = 260;
            BaseTop = 147;

            oItem = oForm.Items.Add("txtMaxImp", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 100;
            oItem.Width = 80;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            //// bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsMaxImp");
            oEditText.String = "3005,06";

            ////**********************************
            //// Adding Static Text item
            ////**********************************

            oItem = oForm.Items.Add("lblMaxImp", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 80;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtMaxImp";
            oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            oStaticText.Caption = "Importe Máximo";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Ejercicio
            //*************************
            // Adding a Text Edit item
            //*************************
            BaseLeft = 10;
            BaseTop = 220;

            oItem = oForm.Items.Add("txtEjer", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsEjer");

            oEditText.String = Convert.ToString(DateTime.Today.Year - 1);

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblEjer", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtEjer";

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Ejercicio";
            #endregion
            #region Selecciona Fichero
            //*************************
            // Adding a Text Edit item
            //*************************
            BaseLeft = 10;
            BaseTop = 240;

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
            #region Fecha Carta
            BaseLeft = 10;
            BaseTop = 260;

            // Date Picker
            //__________________________________________________________________________________________
            ////*************************
            //// Adding a Text Edit item
            ////*************************

            oItem = oForm.Items.Add("txtFecCar", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            //// bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsFecCar");

            ////**********************************
            //// Adding Static Text item
            ////**********************************

            oItem = oForm.Items.Add("lblFecCar", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtFecCar";
            oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            oStaticText.Caption = "Fecha Carta";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region Firmante
            BaseLeft = 10;
            BaseTop = 280;

            // Date Picker
            //__________________________________________________________________________________________
            ////*************************
            //// Adding a Text Edit item
            ////*************************

            oItem = oForm.Items.Add("txtFir", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 120;
            oItem.Width = 300;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            //// bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsFir");

            ////**********************************
            //// Adding Static Text item
            ////**********************************

            oItem = oForm.Items.Add("lblFir", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtFir";
            oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            oStaticText.Caption = "Firmante";

            BaseLeft = 0;
            BaseTop = 0;
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
            #region CheckBox Acreedores
            BaseLeft = 30;
            BaseTop = 47;
            oItem = oForm.Items.Add("chkAcre", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 19;
            oCheckBox = ((SAPbouiCOM.CheckBox)(oItem.Specific));
            oCheckBox.Caption = "Acreedores";
            oCheckBox.DataBind.SetBound(true, "", "dsAcre");
            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region CheckBox Inc Extranjeros
            BaseLeft = 260;
            BaseTop = 110;

            oItem = oForm.Items.Add("chkExtr", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oItem.Left = BaseLeft;
            oItem.Width = 200;
            oItem.Top = BaseTop;
            oItem.Height = 19;
            oCheckBox = ((SAPbouiCOM.CheckBox)(oItem.Specific));
            oCheckBox.Caption = "Incorporar extranjeros";
            // binding the Check box with a data source
            oCheckBox.DataBind.SetBound(true, "", "dsInExt");

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #region CheckBox Inc IC con Retención
            BaseLeft = 260;
            BaseTop = 127;

            oItem = oForm.Items.Add("chkICRet", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oItem.Left = BaseLeft;
            oItem.Width = 200;
            oItem.Top = BaseTop;
            oItem.Height = 19;
            oCheckBox = ((SAPbouiCOM.CheckBox)(oItem.Specific));
            oCheckBox.Caption = "Incorporar IC con Retención";
            // binding the Check box with a data source
            oCheckBox.DataBind.SetBound(true, "", "dsInICRet");

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #endregion

            #region Labels
            BaseTop = 310;
            BaseLeft = 60;
            oItem = oForm.Items.Add("lblProces", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 250;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "";
            #endregion

            #region Option Button
            #region Option Pantalla
            BaseTop = 360;
            BaseLeft = 30;

            oItem = oForm.Items.Add("optPorPan", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 19;

            oOptionBtn = ((SAPbouiCOM.OptionBtn)(oItem.Specific));

            oOptionBtn.Caption = "Por Pantalla";
            oOptionBtn.DataBind.SetBound(true, "", "dsOptDes");
            #endregion
            #region Option Impresora
            oItem = oForm.Items.Add("optPorImp", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop + 19;
            oItem.Height = 19;

            oOptionBtn = ((SAPbouiCOM.OptionBtn)(oItem.Specific));

            oOptionBtn.Caption = "Por Impresora";
            oOptionBtn.GroupWith("optPorPan");
            oOptionBtn.DataBind.SetBound(true, "", "dsOptDes");

            BaseTop = 0;
            BaseLeft = 0;
            #endregion
            #endregion

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

                //  Adding 2 CFL, one for the button and one for the edit text.
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "2";
                oCFLCreationParams.UniqueID = "CFL1DesIC";

                oCFL = oCFLs.Add(oCFLCreationParams);

                oCFLCreationParams.UniqueID = "CFL2DesIC";
                oCFL = oCFLs.Add(oCFLCreationParams);



                //  Adding 2 CFL, one for the button and one for the edit text.
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "2";
                oCFLCreationParams.UniqueID = "CFL1HasIC";

                oCFL = oCFLs.Add(oCFLCreationParams);

                oCFLCreationParams.UniqueID = "CFL2HasIC";
                oCFL = oCFLs.Add(oCFLCreationParams);

            }
            catch
            {

            }
        }

        public void FrmModelo347_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            csUtilidades Utilidades = new csUtilidades();
            BubbleEvent = true;
            SAPbouiCOM.StaticText oStaticText;
            SAPbouiCOM.EditText oEditText;
            SAPbouiCOM.OptionBtn oOptionBtn;

            if (pVal.FormUID == "FrmModelo347")
            {

//                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                        //System.Windows.Forms.Application.Exit();
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                        try
                        {
                            if (pVal.ItemUID == "btnGen" && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                            {
                                if (ValidarGeneracion())
                                {
                                    Genera347();
                                    oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
                                    oStaticText.Caption = "Filtrando registros aptos para el 347";
                                    BorrarDatosInferioresAImporte();
                                    GenerarFichero();
                                    oStaticText.Caption = "Proceso Terminado";
                                    oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optPorPan").Specific;
                                    if (oOptionBtn.Selected)
                                    {
                                        Imprimir347(csVariablesGlobales.MenuImprimirPorPantalla, "Informe347.rpt");
                                    }
                                    else
                                    {
                                        Imprimir347(csVariablesGlobales.MenuImprimirPorImpresora, "Informe347.rpt");
                                    }
                                    BubbleEvent = false;
                                }
                            }
                            if (pVal.ItemUID == "btnInf" && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                            {
                                oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optPorPan").Specific;
                                if (oOptionBtn.Selected)
                                {
                                    Imprimir347(csVariablesGlobales.MenuImprimirPorPantalla, "Informe347.rpt");
                                }
                                else
                                {
                                    Imprimir347(csVariablesGlobales.MenuImprimirPorImpresora, "Informe347.rpt");
                                }
                                BubbleEvent = false;
                            }
                            if (pVal.ItemUID == "btnCar" && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                            {
                                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecCar").Specific;
                                if (oEditText.String != "")
                                {
                                    oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optPorPan").Specific;
                                    if (oOptionBtn.Selected)
                                    {
                                        Imprimir347(csVariablesGlobales.MenuImprimirPorPantalla, "Carta347.rpt");
                                    }
                                    else
                                    {
                                        Imprimir347(csVariablesGlobales.MenuImprimirPorImpresora, "Carta347.rpt");
                                    }
                                }
                                else
                                {
                                    oEditText.Active = true;
                                    csVariablesGlobales.SboApp.MessageBox("La fecha de la carta es obligatoria", 1, "Ok", "", "");
                                }
                                BubbleEvent = false;
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
                                OpenFileDialog.InitialDirectory = csUtilidades.DameValor("[@SIA_PARAM]", "U_Ruta", "U_TipoRuta = 'Destino Ficheros Modelos'");
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

        private void FrmModelo347_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {

            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:

                    csVariablesGlobales.SboApp.MessageBox("Se canceló la ejecución de SAP" + Environment.NewLine + "Se cerrará el Addon SIA", 1, "Ok", "", "");
                    //System.Windows.Forms.Application.Exit();
                    break;
            }
        }

        public void CargarDatosInicialesDePantalla()
        {
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.OptionBtn oOptionBtn = null;
            csUtilidades Utilidades = new csUtilidades();
            oForm = csVariablesGlobales.SboApp.Forms.Item("FrmModelo347");
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtMaxImp").Specific;
            oEditText.String = csUtilidades.DameValor("OADM", "MinAmnt347", "");
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtEjer").Specific;
            oEditText.String = Convert.ToString(DateTime.Today.Year - 1);
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optPorPan").Specific;
            oOptionBtn.Selected = true;
        }

        private void Genera347()
        {
            #region Declaraciones
            Recordset oRecordSet;
            SAPbouiCOM.CheckBox oCheckBox = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbobsCOM.UserTable pReTabla = null;
            csUtilidades Utilidades = new csUtilidades();
            string StrRetencion;
            string StrSQL = "";
            string StrRetCom;
            string StrCodMax;
            string DesdeFecha = "";
            string HastaFecha = "";
            string DesdeIC = "";
            string HastaIC = "";
            bool ValorDeudor;
            bool ValorAcreedor;
            #endregion
            #region Asignación Variables
            pReTabla = csVariablesGlobales.oCompany.UserTables.Item("SIA_MOD347");
            StrRetencion = csUtilidades.DameValor("OADM", "SHandleWT", "");
            StrRetCom = csUtilidades.DameValor("OADM", "pHandleWT", "");
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkDeud").Specific;
            ValorDeudor = oCheckBox.Checked;
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkAcre").Specific;
            ValorAcreedor = oCheckBox.Checked;
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtDesFec").Specific;
            DesdeFecha = oEditText.String;
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtHasFec").Specific;
            HastaFecha = oEditText.String;
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtDesdeIC").Specific;
            DesdeIC = oEditText.String;
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtHastaIC").Specific;
            HastaIC = oEditText.String;
            #endregion

            //Limpiar tabla Modelo 347
            StrSQL = "DELETE FROM [@SIA_MOD347]";
            SqlCommand cmdBorrar = new SqlCommand(StrSQL, csVariablesGlobales.conAddon);
            cmdBorrar.ExecuteNonQuery();

            #region CARGO FACTURA
            oRecordSet = (SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            StrSQL = "SELECT T3.CardCode AS CodIC, T3.CardName AS NomIC, " +
                     "T3.CardType AS TipIC, T3.LicTradNum AS NifIC, " +
                     "SUM(T0.DocTotal - T0.VatSum) AS Base, " +
                     "T5.Name AS Prov, T4.Name AS Pais, " +
                     "T0.[Indicator] AS Indic, T3.ZipCode AS CP, " +
                     "MIN(T4.ReportCode) AS CodInf, " +
                     "SUM(T0.VatSum - T0.EquVatSum) AS Iva, " +
                     "SUM(T0.DocTotal) AS Total " +
                     "FROM OCST T5 RIGHT OUTER JOIN " +
                     "OCRY T4 ON T5.Country = T4.Code RIGHT OUTER JOIN " +
                     "OINV T0 INNER JOIN " +
                     "OCRD T3 ON T3.CardCode = T0.CardCode ON T4.Code = T3.Country AND T5.Code = T3.State1 " +
                     "WHERE (T0.DOCDATE >= '" + DesdeFecha.Substring(6, 4) + "/" + DesdeFecha.Substring(3, 2) + "/" + DesdeFecha.Substring(0, 2) + "' " +
                     "AND T0.DOCDATE    <= '" + HastaFecha.Substring(6, 4) + "/" + HastaFecha.Substring(3, 2) + "/" + HastaFecha.Substring(0, 2) + "' " +
                     "AND T3.CARDCODE   >= '" + DesdeIC + "' " +
                     "AND T3.CARDCODE   <= '" + HastaIC + "') ";
            if (ValorAcreedor && ValorDeudor)
            {
                StrSQL = StrSQL + "AND (T3.CardType = 'C' OR T3.CardType = 'S') ";
            }
            else
            {
                if (ValorDeudor)
                {
                    StrSQL = StrSQL + "AND T3.CardType = 'C' ";
                }
                if (ValorAcreedor)
                {
                    StrSQL = StrSQL + "AND T3.CardType = 'S' ";
                }
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkExtr").Specific;
            if (oCheckBox.Checked == false) //sólo españoles
            {
                StrSQL = StrSQL + "AND T3.Country = 'ES' ";
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkICRet").Specific;
            if (oCheckBox.Checked == false && StrRetencion == "Y") //sólo incluimos a los que no tienen retencion, si el sistema asi lo tiene
            {
                StrSQL = StrSQL + "AND (T3.WTLiable='N' OR T3.WTLiable='') ";
            }
            StrSQL = StrSQL + "GROUP BY T3.CardCode, T0.[Indicator], T3.CardName, T3.CardType, " +
                     " T3.LicTradNum, T5.Name, T4.Name, T3.ZipCode";

            SqlDataAdapter daSQLFactura = new SqlDataAdapter(StrSQL, csVariablesGlobales.conAddon);
            System.Data.DataTable dtSQLFactura = new System.Data.DataTable();
            DataRow drFactura;
            daSQLFactura.Fill(dtSQLFactura);
            SAPbouiCOM.StaticText oStaticText;
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            if (dtSQLFactura.Rows.Count > 0)
            {
                for (int i = 0; i < dtSQLFactura.Rows.Count; i++)
                {
                    oStaticText.Caption = "Cargando Factura " + Convert.ToUInt32(i + 1) + " de " + dtSQLFactura.Rows.Count;
                    System.Windows.Forms.Application.DoEvents();
                    drFactura = dtSQLFactura.Rows[i];
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = csUtilidades.DameValor("[@SIA_MOD347]", "Max(Code)", "");
                    if (StrCodMax == "")
                    {
                        StrCodMax = "0";
                    }
                    StrCodMax = Convert.ToString(Convert.ToInt32(StrCodMax) + 1);
                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    pReTabla.Code = StrCodMax;
                    pReTabla.Name = StrCodMax;
                    pReTabla.UserFields.Fields.Item("U_CodIC").Value = drFactura["CodIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NomIC").Value = drFactura["NomIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_TipIC").Value = drFactura["TipIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NifIC").Value = drFactura["NifIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(drFactura["Base"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Prov").Value = drFactura["Prov"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Pais").Value = drFactura["Pais"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Indic").Value = drFactura["Indic"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CP").Value = drFactura["CP"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CodInf").Value = drFactura["CodInf"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Iva").Value = Convert.ToDouble(drFactura["Iva"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Total").Value = Convert.ToDouble(drFactura["Total"].ToString());
                    pReTabla.UserFields.Fields.Item("U_TipDoc").Value = "FACTURA";

                    if (pReTabla.Add() != 0)
                    {
                        csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                        return;
                    }
                }
            }
            else
            {
                oStaticText.Caption = "No hay facturas de cargo que tratar.";
            }
            #endregion
            #region CARGO ABONO
            oRecordSet = (SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            StrSQL = "SELECT T3.CardCode AS CodIC, MIN(T3.CardName) AS NomIC, " +
                     "MIN(T3.CardType) AS TipIC, MIN(T3.LicTradNum) AS NifIC, " +
                     "SUM(T0.DocTotal - T0.VatSum)AS Base, " +
                     "MIN(T5.Name) AS Prov, MIN(T4.Name) AS Pais, " +
                     "T0.[Indicator] AS Indic, MIN(T3.ZipCode) AS CP, " +
                     "MIN(T4.ReportCode) AS CodInf, " +
                     "SUM(T0.VatSum - T0.EquVatSum) AS Iva, " +
                     "SUM(T0.DocTotal) AS Total " +
                     "FROM OCST T5 RIGHT OUTER JOIN " +
                     "OCRY T4 ON T5.Country = T4.Code RIGHT OUTER JOIN " +
                     "ORIN T0 INNER JOIN " +
                     "OCRD T3 ON T3.CardCode = T0.CardCode ON " +
                     "T4.Code = T3.Country AND T5.Code = T3.State1 " +
                     "WHERE (T0.DOCDATE >= '" + DesdeFecha.Substring(6, 4) + "/" + DesdeFecha.Substring(3, 2) + "/" + DesdeFecha.Substring(0, 2) + "' " +
                     "AND T0.DOCDATE    <= '" + HastaFecha.Substring(6, 4) + "/" + HastaFecha.Substring(3, 2) + "/" + HastaFecha.Substring(0, 2) + "' " +
                     "AND T3.CARDCODE   >= '" + DesdeIC + "' " +
                     "AND T3.CARDCODE   <= '" + HastaIC + "') ";
            if (ValorAcreedor && ValorDeudor)
            {
                StrSQL = StrSQL + "AND (T3.CardType = 'C' OR T3.CardType = 'S') ";
            }
            else
            {
                if (ValorDeudor)
                {
                    StrSQL = StrSQL + "AND T3.CardType = 'C' ";
                }
                if (ValorAcreedor)
                {
                    StrSQL = StrSQL + "AND T3.CardType = 'S' ";
                }
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkExtr").Specific;
            if (oCheckBox.Checked == false) //sólo españoles
            {
                StrSQL = StrSQL + "AND T3.Country = 'ES' ";
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkICRet").Specific;
            if (oCheckBox.Checked == false && StrRetencion == "Y") //sólo incluimos a los que no tienen retencion, si el sistema asi lo tiene
            {
                StrSQL = StrSQL + "AND (T3.WTLiable='N' OR T3.WTLiable='') ";
            }
            StrSQL = StrSQL + "GROUP BY T3.CardCode, T0.[Indicator]";

            SqlDataAdapter daSQLAbono = new SqlDataAdapter(StrSQL, csVariablesGlobales.conAddon);
            System.Data.DataTable dtSQLAbono = new System.Data.DataTable();
            DataRow drSQLAbono;
            daSQLAbono.Fill(dtSQLAbono);
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            if (dtSQLAbono.Rows.Count > 0)
            {
                for (int i = 0; i < dtSQLAbono.Rows.Count; i++)
                {
                    oStaticText.Caption = "Cargando Abono " + Convert.ToUInt32(i + 1) + " de " + dtSQLAbono.Rows.Count;
                    System.Windows.Forms.Application.DoEvents();
                    drSQLAbono = dtSQLAbono.Rows[i];
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = csUtilidades.DameValor("[@SIA_MOD347]", "Max(Code)", "");
                    if (StrCodMax == "")
                    {
                        StrCodMax = "0";
                    }
                    StrCodMax = Convert.ToString(Convert.ToInt32(StrCodMax) + 1);
                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    pReTabla.Code = StrCodMax;
                    pReTabla.Name = StrCodMax;
                    pReTabla.UserFields.Fields.Item("U_CodIC").Value = drSQLAbono["CodIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NomIC").Value = drSQLAbono["NomIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_TipIC").Value = drSQLAbono["TipIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NifIC").Value = drSQLAbono["NifIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(drSQLAbono["Base"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Prov").Value = drSQLAbono["Prov"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Pais").Value = drSQLAbono["Pais"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Indic").Value = drSQLAbono["Indic"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CP").Value = drSQLAbono["CP"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CodInf").Value = drSQLAbono["CodInf"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Iva").Value = Convert.ToDouble(drSQLAbono["Iva"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Total").Value = Convert.ToDouble((-1) * Convert.ToDouble(drSQLAbono["Total"]));
                    pReTabla.UserFields.Fields.Item("U_TipDoc").Value = "ABONO";

                    if (pReTabla.Add() != 0)
                    {
                        csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                        return;
                    }
                }
            }
            else
            {
                oStaticText.Caption = "No hay facturas de abono que tratar.";
            }
            #endregion
            #region CARGO ANTICIPO
            oRecordSet = (SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            StrSQL = "SELECT T3.[CardCode] as CodIC , MIN(T3.[CardName]) as NomIC, " +
                     "MIN(T3.[CardType]) as TipIC, MIN(T3.[LicTradNum]) as NifIC, " +
                     "SUM(T1.[LineTotal] / 100 * (100 - T0.[DiscPrcnt])) as Base, " +
                     "MIN(T5.[Name]) as Prov, MIN(T4.[Name]) as Pais, " +
                     "T0.[Indicator] as Indic, MIN(T3.[ZipCode]) as CP, " +
                     "MIN(T4.[ReportCode]) as CodInf, SUM(T1.[VatSum]) as Iva, " +
                     "SUM(T1.[LineTotal] / 100 * T0.[DpmPrcnt]) as Anticipo, " +
                     "T2.[AcqstnRvrs] as Adquisicion, SUM(T6.LineTotal) AS BasePortes, " +
                     "SUM(T6.VatSum) AS IvaPortes " +
                     "FROM ODPI AS T0 INNER JOIN " +
                     "DPI1 AS T1 ON T1.DocEntry = T0.DocEntry INNER JOIN " +
                     "OVTG AS T2 ON T2.Code = T1.VatGroup INNER JOIN " +
                     "OCRD AS T3 ON T3.CardCode = T0.CardCode LEFT OUTER JOIN " +
                     "OCST AS T5 ON T3.Country = T5.Country AND T3.State1 = T5.Code LEFT OUTER JOIN " +
                     "DPI3 AS T6 ON T0.DocEntry = T6.DocEntry LEFT OUTER JOIN " +
                     "OCRY AS T4 ON T3.Country = T4.Code " +
                     "WHERE  ((T2.Code = 'EI') OR " +
                     "(T2.Code = 'EIT') OR" +
                     "(T2.Code = 'EX') OR" +
                     "(T2.Code = 'R0') OR" +
                     "(T2.Code = 'R1') OR" +
                     "(T2.Code = 'R2') OR" +
                     "(T2.Code = 'R3') OR" +
                     "(T2.Code = 'RA0') OR" +
                     "(T2.Code = 'RA1') OR" +
                     "(T2.Code = 'RA2') OR" +
                     "(T2.Code = 'RA3') OR" +
                     "(T2.Code = 'RE1') OR" +
                     "(T2.Code = 'RE2') OR" +
                     "(T2.Code = 'RE3') OR" +
                     "(T2.Code = 'RIN0') OR" +
                     "(T2.Code = 'RIN1') OR" +
                     "(T2.Code = 'RIN2') OR" +
                     "(T2.Code = 'RIN3') OR" +
                     "(T2.Code = 'A1') OR" +
                     "(T2.Code = 'A2') OR" +
                     "(T2.Code = 'A3') OR" +
                     "(T2.Code = 'AI') OR" +
                     "(T2.Code = 'I0') OR" +
                     "(T2.Code = 'I1') OR" +
                     "(T2.Code = 'I2') OR" +
                     "(T2.Code = 'I3') OR" +
                     "(T2.Code = 'IBI0') OR" +
                     "(T2.Code = 'IBI1') OR" +
                     "(T2.Code = 'IBI2') OR" +
                     "(T2.Code = 'IBI3') OR" +
                     "(T2.Code = 'ND0') OR" +
                     "(T2.Code = 'ND1') OR" +
                     "(T2.Code = 'ND2') OR" +
                     "(T2.Code = 'ND3') OR" +
                     "(T2.Code = 'S0') OR" +
                     "(T2.Code = 'S1') OR" +
                     "(T2.Code = 'S2') OR" +
                     "(T2.Code = 'S3') OR" +
                     "(T2.Code = 'S4') OR" +
                     "(T2.Code = 'S5') OR" +
                     "(T2.Code = 'S6') OR" +
                     "(T2.Code = 'SA0') OR" +
                     "(T2.Code = 'SA1') OR" +
                     "(T2.Code = 'SA2') OR" +
                     "(T2.Code = 'SA3') OR" +
                     "(T2.Code = 'SI0') OR" +
                     "(T2.Code = 'SI1') OR" +
                     "(T2.Code = 'SI2') OR" +
                     "(T2.Code = 'SI3') OR" +
                     "(T2.Code = 'SIN0') OR" +
                     "(T2.Code = 'SIN1') OR" +
                     "(T2.Code = 'SIN2') OR" +
                     "(T2.Code = 'SIN3') OR" +
                     "(T2.Code = 'SV'))" +
                     "AND (T0.DOCDATE >= '" + DesdeFecha.Substring(6, 4) + "/" + DesdeFecha.Substring(3, 2) + "/" + DesdeFecha.Substring(0, 2) + "' " +
                     "AND T0.DOCDATE    <= '" + HastaFecha.Substring(6, 4) + "/" + HastaFecha.Substring(3, 2) + "/" + HastaFecha.Substring(0, 2) + "' " +
                     "AND T3.CARDCODE   >= '" + DesdeIC + "' " +
                     "AND T3.CARDCODE   <= '" + HastaIC + "') ";
            if (ValorAcreedor && ValorDeudor)
            {
                StrSQL = StrSQL + "AND (T3.CardType = 'C' OR T3.CardType = 'S') ";
            }
            else
            {
                if (ValorDeudor)
                {
                    StrSQL = StrSQL + "AND T3.CardType = 'C' ";
                }
                if (ValorAcreedor)
                {
                    StrSQL = StrSQL + "AND T3.CardType = 'S' ";
                }
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkExtr").Specific;
            if (oCheckBox.Checked == false) //sólo españoles
            {
                StrSQL = StrSQL + "AND T3.Country = 'ES' ";
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkICRet").Specific;
            if (oCheckBox.Checked == false && StrRetencion == "Y") //sólo incluimos a los que no tienen retencion, si el sistema asi lo tiene
            {
                StrSQL = StrSQL + "AND (T3.WTLiable='N' OR T3.WTLiable='') ";
            }
            StrSQL = StrSQL + "GROUP BY T3.CardCode, T0.[Indicator], T2.AcqstnRvrs";

            SqlDataAdapter daSQLCargoAbono = new SqlDataAdapter(StrSQL, csVariablesGlobales.conAddon);
            System.Data.DataTable dtSQLCargoAbono = new System.Data.DataTable();
            DataRow drSQLCargoAbono;
            daSQLCargoAbono.Fill(dtSQLCargoAbono);
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            if (dtSQLCargoAbono.Rows.Count > 0)
            {
                double IvaPortes = 0.0;
                double BasePortes = 0.0;
                double Saldo = 0.0;
                for (int i = 0; i < dtSQLCargoAbono.Rows.Count; i++)
                {
                    oStaticText.Caption = "Cargando Cargo Anticipo " + Convert.ToUInt32(i + 1) + " de " + dtSQLCargoAbono.Rows.Count;
                    System.Windows.Forms.Application.DoEvents();
                    drSQLCargoAbono = dtSQLCargoAbono.Rows[i];
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = csUtilidades.DameValor("[@SIA_MOD347]", "Max(Code)", "");
                    if (StrCodMax == "")
                    {
                        StrCodMax = "0";
                    }
                    if (drSQLCargoAbono["IvaPortes"].ToString() == null || drSQLCargoAbono["IvaPortes"].ToString() == "")
                    {
                        IvaPortes = 0;
                    }
                    else
                    {
                        IvaPortes = Convert.ToDouble(drSQLCargoAbono["IvaPortes"].ToString());
                    }
                    if (drSQLCargoAbono["BasePortes"].ToString() == null || drSQLCargoAbono["BasePortes"].ToString() == "")
                    {
                        BasePortes = 0;
                    }
                    else
                    {
                        BasePortes = Convert.ToDouble(drSQLCargoAbono["BasePortes"].ToString());
                    }
                    Saldo = Convert.ToDouble(drSQLCargoAbono["Base"].ToString()) +
                            Convert.ToDouble(drSQLCargoAbono["Iva"].ToString()) +
                            IvaPortes + BasePortes;
                    StrCodMax = Convert.ToString(Convert.ToInt32(StrCodMax) + 1);
                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    pReTabla.Code = StrCodMax;
                    pReTabla.Name = StrCodMax;
                    pReTabla.UserFields.Fields.Item("U_CodIC").Value = drSQLCargoAbono["CodIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NomIC").Value = drSQLCargoAbono["NomIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_TipIC").Value = drSQLCargoAbono["TipIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NifIC").Value = drSQLCargoAbono["NifIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(drSQLCargoAbono["Base"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Prov").Value = drSQLCargoAbono["Prov"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Pais").Value = drSQLCargoAbono["Pais"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Indic").Value = drSQLCargoAbono["Indic"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CP").Value = drSQLCargoAbono["CP"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CodInf").Value = drSQLCargoAbono["CodInf"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Iva").Value = Convert.ToDouble(drSQLCargoAbono["Iva"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Anticip").Value = Convert.ToDouble(drSQLCargoAbono["Anticipo"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Adquisi").Value = drSQLCargoAbono["Adquisicion"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Total").Value = Saldo;
                    pReTabla.UserFields.Fields.Item("U_TipDoc").Value = "ANTICIPO";

                    if (pReTabla.Add() != 0)
                    {
                        csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                        return;
                    }
                }
            }
            else
            {
                oStaticText.Caption = "No hay cargos de anticipo que tratar.";
            }
            #endregion
            #region CARGO DIARIO
            oRecordSet = (SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            StrSQL = "SELECT T4.[CardCode] as CodIC, MIN(T4.[CardName]) as NomIC, " +
                     "T4.CardType as TipIC, MIN(T4.[LicTradNum]) as NifIC, " +
                     "SUM(T1.Debit) AS Debe, SUM(T1.Credit) AS Haber, " +
                     "SUM(T1.[BaseSum]) as Base, MIN(T6.[Name]) as Prov, " +
                     "MIN(T5.[Name]) as Pais, T0.[Indicator] as Indic, " +
                     "MIN(T4.[ZipCode]) as CP," +
                     "MIN(T5.[ReportCode]) as CodInf " +
                     "FROM OJDT AS T0 INNER JOIN " +
                     "JDT1 AS T1 ON T0.TransId = T1.TransId INNER JOIN " +
                     "OCRD AS T4 ON T1.ShortName = T4.CardCode LEFT OUTER JOIN " +
                     "OCST AS T6 ON T4.Country = T6.Country AND T6.Code = T4.State1 LEFT OUTER JOIN " +
                     "OCRY AS T5 ON T5.Code = T4.Country " +
                     "WHERE  T0.Report347 = 'Y'  AND  T1.TransType='30' " +
                     "AND (T1.RefDate >= '" + DesdeFecha.Substring(6, 4) + "/" + DesdeFecha.Substring(3, 2) + "/" + DesdeFecha.Substring(0, 2) + "' " +
                     "AND T1.RefDate  <= '" + HastaFecha.Substring(6, 4) + "/" + HastaFecha.Substring(3, 2) + "/" + HastaFecha.Substring(0, 2) + "' " +
                     "AND T4.CARDCODE >= '" + DesdeIC + "' " +
                     "AND T4.CARDCODE <= '" + HastaIC + "') ";
            if (ValorAcreedor && ValorDeudor)
            {
                StrSQL = StrSQL + "AND (T4.CardType = 'C' OR T4.CardType = 'S') ";
            }
            else
            {
                if (ValorDeudor)
                {
                    StrSQL = StrSQL + "AND T4.CardType = 'C' ";
                }
                if (ValorAcreedor)
                {
                    StrSQL = StrSQL + "AND T4.CardType = 'S' ";
                }
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkICRet").Specific;
            if (oCheckBox.Checked == false && StrRetencion == "Y") //sólo incluimos a los que no tienen retencion, si el sistema asi lo tiene
            {
                StrSQL = StrSQL + "AND (T4.WTLiable='N' OR T4.WTLiable='') ";
            }
            StrSQL = StrSQL + "GROUP BY T4.CardCode, T4.CardType, T0.Indicator";

            SqlDataAdapter daSQLCargoDiario = new SqlDataAdapter(StrSQL, csVariablesGlobales.conAddon);
            System.Data.DataTable dtSQLCargoDiario = new System.Data.DataTable();
            DataRow drSQLCargoDiario;
            daSQLCargoDiario.Fill(dtSQLCargoDiario);
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            if (dtSQLCargoDiario.Rows.Count > 0)
            {
                double Saldo = 0.0;
                for (int i = 0; i < dtSQLCargoDiario.Rows.Count; i++)
                {
                    oStaticText.Caption = "Cargando Cargo Diario " + Convert.ToUInt32(i + 1) + " de " + dtSQLCargoDiario.Rows.Count;
                    System.Windows.Forms.Application.DoEvents();
                    drSQLCargoDiario = dtSQLCargoDiario.Rows[i];
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = csUtilidades.DameValor("[@SIA_MOD347]", "Max(Code)", "");
                    if (StrCodMax == "")
                    {
                        StrCodMax = "0";
                    }
                    StrCodMax = Convert.ToString(Convert.ToInt32(StrCodMax) + 1);
                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    pReTabla.Code = StrCodMax;
                    pReTabla.Name = StrCodMax;
                    pReTabla.UserFields.Fields.Item("U_CodIC").Value = drSQLCargoDiario["CodIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NomIC").Value = drSQLCargoDiario["NomIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_TipIC").Value = drSQLCargoDiario["TipIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NifIC").Value = drSQLCargoDiario["NifIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(drSQLCargoDiario["Base"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Prov").Value = drSQLCargoDiario["Prov"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Pais").Value = drSQLCargoDiario["Pais"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Indic").Value = drSQLCargoDiario["Indic"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CP").Value = drSQLCargoDiario["CP"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CodInf").Value = drSQLCargoDiario["CodInf"].ToString();
                    if (drSQLCargoDiario["TipIC"].ToString() == "C")
                    {
                        pReTabla.UserFields.Fields.Item("U_Iva").Value = 0;
                        if (drSQLCargoDiario["Base"].ToString() == "")
                        {
                            Saldo = 0.0;
                        }
                        else
                        {
                            Saldo = Convert.ToDouble(drSQLCargoDiario["Debe"].ToString());
                        }
                    }
                    else
                    {
                        pReTabla.UserFields.Fields.Item("U_Iva").Value = 0;
                        if (drSQLCargoDiario["Base"].ToString() == "")
                        {
                            Saldo = 0.0;
                        }
                        else
                        {
                            Saldo = Convert.ToDouble(drSQLCargoDiario["Haber"].ToString());
                        }
                    }
                    pReTabla.UserFields.Fields.Item("U_Anticip").Value = "0";
                    pReTabla.UserFields.Fields.Item("U_Total").Value = Saldo;
                    pReTabla.UserFields.Fields.Item("U_TipDoc").Value = "DIARIO";

                    if (pReTabla.Add() != 0)
                    {
                        csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                        return;
                    }
                }
            }
            else
            {
                oStaticText.Caption = "No hay cargos de diario que tratar.";
            }
            #endregion
            //Proveedor
            #region CARGO FACTURA
            oRecordSet = (SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            StrSQL = "SELECT T3.CardCode AS CodIC, MIN(T3.CardName) AS NomIC, " +
                     "MIN(T3.CardType) AS TipIC, MIN(T3.LicTradNum) AS NifIC, " +
                     "MIN(T5.Name) AS Prov, " +
                     "MIN(T4.Name) AS Pais, T0.[Indicator] AS Indic, MIN(T3.ZipCode) AS CP, " +
                     "MIN(T4.ReportCode) AS CodInf, " +
                     "SUM(T0.DocTotal - T0.VatSum) AS base, " +
                     "SUM(T0.VatSum - T0.EquVatSum) AS Iva, " +
                     "SUM(T0.DocTotal) AS Total " +
                     "FROM OPCH AS T0 INNER JOIN " +
                     "OCRD AS T3 ON T3.CardCode = T0.CardCode LEFT OUTER JOIN " +
                     "OCST AS T5 ON T3.Country = T5.Country AND T3.State1 = T5.Code LEFT OUTER JOIN " +
                     "OCRY AS T4 ON T3.Country = T4.Code " +
                     "WHERE (T0.DOCDATE >= '" + DesdeFecha.Substring(6, 4) + "/" + DesdeFecha.Substring(3, 2) + "/" + DesdeFecha.Substring(0, 2) + "' " +
                     "AND T0.DOCDATE    <= '" + HastaFecha.Substring(6, 4) + "/" + HastaFecha.Substring(3, 2) + "/" + HastaFecha.Substring(0, 2) + "' " +
                     "AND T3.CARDCODE   >= '" + DesdeIC + "' " +
                     "AND T3.CARDCODE   <= '" + HastaIC + "') ";
            if (ValorAcreedor && ValorDeudor)
            {
                StrSQL = StrSQL + "AND (T3.CardType = 'C' OR T3.CardType = 'S') ";
            }
            else
            {
                if (ValorDeudor)
                {
                    StrSQL = StrSQL + "AND T3.CardType = 'C' ";
                }
                if (ValorAcreedor)
                {
                    StrSQL = StrSQL + "AND T3.CardType = 'S' ";
                }
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkExtr").Specific;
            if (oCheckBox.Checked == false) //sólo españoles
            {
                StrSQL = StrSQL + "AND T3.Country = 'ES' ";
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkICRet").Specific;
            if (oCheckBox.Checked == false && StrRetencion == "Y") //sólo incluimos a los que no tienen retencion, si el sistema asi lo tiene
            {
                StrSQL = StrSQL + "AND (T3.WTLiable='N' OR T3.WTLiable='') ";
            }
            StrSQL = StrSQL + "GROUP BY T3.CardCode, T0.[Indicator]";

            SqlDataAdapter daSQLFacturaProv = new SqlDataAdapter(StrSQL, csVariablesGlobales.conAddon);
            System.Data.DataTable dtSQLFacturaProv = new System.Data.DataTable();
            DataRow drFacturaProv;
            daSQLFacturaProv.Fill(dtSQLFacturaProv);
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            if (dtSQLFacturaProv.Rows.Count > 0)
            {
                for (int i = 0; i < dtSQLFacturaProv.Rows.Count; i++)
                {
                    oStaticText.Caption = "Cargando Factura " + Convert.ToUInt32(i + 1) + " de " + dtSQLFacturaProv.Rows.Count;
                    System.Windows.Forms.Application.DoEvents();
                    drFacturaProv = dtSQLFacturaProv.Rows[i];
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = csUtilidades.DameValor("[@SIA_MOD347]", "Max(Code)", "");
                    if (StrCodMax == "")
                    {
                        StrCodMax = "0";
                    }
                    //dblSaldo = Convert.ToDouble(drFactura["Total"].ToString());
                    StrCodMax = Convert.ToString(Convert.ToInt32(StrCodMax) + 1);
                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    pReTabla.Code = StrCodMax;
                    pReTabla.Name = StrCodMax;
                    pReTabla.UserFields.Fields.Item("U_CodIC").Value = drFacturaProv["CodIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NomIC").Value = drFacturaProv["NomIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_TipIC").Value = drFacturaProv["TipIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NifIC").Value = drFacturaProv["NifIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(drFacturaProv["Base"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Prov").Value = drFacturaProv["Prov"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Pais").Value = drFacturaProv["Pais"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Indic").Value = drFacturaProv["Indic"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CP").Value = drFacturaProv["CP"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CodInf").Value = drFacturaProv["CodInf"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Iva").Value = Convert.ToDouble(drFacturaProv["Iva"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Total").Value = Convert.ToDouble(drFacturaProv["Total"].ToString());
                    pReTabla.UserFields.Fields.Item("U_TipDoc").Value = "FACTURA";

                    if (pReTabla.Add() != 0)
                    {
                        csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                        return;
                    }
                }
            }
            else
            {
                oStaticText.Caption = "No hay facturas de cargo que tratar.";
            }
            #endregion
            #region CARGO ABONO
            oRecordSet = (SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            StrSQL = "SELECT T3.CardCode AS CodIC, MIN(T3.CardName) AS NomIC, " +
                     "MIN(T3.CardType) AS TipIC, MIN(T3.LicTradNum) AS NifIC, " +
                     "MIN(T5.Name) AS Prov, MIN(T4.Name) AS Pais, " +
                     "T0.[Indicator] AS Indic, MIN(T3.ZipCode) AS CP, " +
                     "MIN(T4.ReportCode) AS CodInf, SUM(T0.DocTotal - T0.VatSum) AS base, " +
                     "SUM(T0.VatSum - T0.EquVatSum) AS Iva, SUM(T0.DocTotal) AS Total " +
                     "FROM ORPC AS T0 INNER JOIN " +
                     "OCRD AS T3 ON T3.CardCode = T0.CardCode LEFT OUTER JOIN " +
                     "OCST AS T5 ON T3.Country = T5.Country AND T3.State1 = T5.Code LEFT OUTER JOIN " +
                     "OCRY AS T4 ON T3.Country = T4.Code " +
                     "WHERE (T0.DOCDATE >= '" + DesdeFecha.Substring(6, 4) + "/" + DesdeFecha.Substring(3, 2) + "/" + DesdeFecha.Substring(0, 2) + "' " +
                     "AND T0.DOCDATE    <= '" + HastaFecha.Substring(6, 4) + "/" + HastaFecha.Substring(3, 2) + "/" + HastaFecha.Substring(0, 2) + "' " +
                     "AND T3.CARDCODE   >= '" + DesdeIC + "' " +
                     "AND T3.CARDCODE   <= '" + HastaIC + "') ";
            if (ValorAcreedor && ValorDeudor)
            {
                StrSQL = StrSQL + "AND (T3.CardType = 'C' OR T3.CardType = 'S') ";
            }
            else
            {
                if (ValorDeudor)
                {
                    StrSQL = StrSQL + "AND T3.CardType = 'C' ";
                }
                if (ValorAcreedor)
                {
                    StrSQL = StrSQL + "AND T3.CardType = 'S' ";
                }
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkExtr").Specific;
            if (oCheckBox.Checked == false) //sólo españoles
            {
                StrSQL = StrSQL + "AND T3.Country = 'ES' ";
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkICRet").Specific;
            if (oCheckBox.Checked == false && StrRetencion == "Y") //sólo incluimos a los que no tienen retencion, si el sistema asi lo tiene
            {
                StrSQL = StrSQL + "AND (T3.WTLiable='N' OR T3.WTLiable='') ";
            }
            StrSQL = StrSQL + " GROUP BY T3.CardCode, T0.[Indicator]";
            SqlDataAdapter daSQLCargoAbonoProv = new SqlDataAdapter(StrSQL, csVariablesGlobales.conAddon);
            System.Data.DataTable dtSQLCargoAbonoProv = new System.Data.DataTable();
            DataRow drCargoAbonoProv;
            daSQLCargoAbonoProv.Fill(dtSQLCargoAbonoProv);
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            if (dtSQLCargoAbonoProv.Rows.Count > 0)
            {
                for (int i = 0; i < dtSQLCargoAbonoProv.Rows.Count; i++)
                {
                    oStaticText.Caption = "Cargando Factura " + Convert.ToUInt32(i + 1) + " de " + dtSQLCargoAbonoProv.Rows.Count;
                    System.Windows.Forms.Application.DoEvents();
                    drCargoAbonoProv = dtSQLFacturaProv.Rows[i];
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = csUtilidades.DameValor("[@SIA_MOD347]", "Max(Code)", "");
                    if (StrCodMax == "")
                    {
                        StrCodMax = "0";
                    }
                    //dblSaldo = Convert.ToDouble(drFactura["Total"].ToString());
                    StrCodMax = Convert.ToString(Convert.ToInt32(StrCodMax) + 1);
                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    pReTabla.Code = StrCodMax;
                    pReTabla.Name = StrCodMax;
                    pReTabla.UserFields.Fields.Item("U_CodIC").Value = drCargoAbonoProv["CodIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NomIC").Value = drCargoAbonoProv["NomIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_TipIC").Value = drCargoAbonoProv["TipIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NifIC").Value = drCargoAbonoProv["NifIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(drCargoAbonoProv["Base"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Prov").Value = drCargoAbonoProv["Prov"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Pais").Value = drCargoAbonoProv["Pais"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Indic").Value = drCargoAbonoProv["Indic"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CP").Value = drCargoAbonoProv["CP"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CodInf").Value = drCargoAbonoProv["CodInf"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Iva").Value = Convert.ToDouble(drCargoAbonoProv["Iva"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Total").Value = ((-1) * Convert.ToDouble(drCargoAbonoProv["Total"].ToString()));
                    pReTabla.UserFields.Fields.Item("U_TipDoc").Value = "ABONO";

                    if (pReTabla.Add() != 0)
                    {
                        csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                        return;
                    }
                }
            }
            else
            {
                oStaticText.Caption = "No hay abonos de cargo que tratar.";
            }
            #endregion
            #region CARGO ANTICIPO
            oRecordSet = (SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            StrSQL = "SELECT T3.[CardCode] as CodIC, MIN(T3.[CardName]) as NomIC, " +
                     "MIN(T3.[CardType]) as TipIC, MIN(T3.[LicTradNum]) as NifIC, " +
                     "SUM(T1.[LineTotal] / 100 * (100 - T0.[DiscPrcnt])) as Base, " +
                     "MIN(T5.[Name]) as Prov, MIN(T4.[Name]) as Pais, " +
                     "T0.[Indicator] as Indic, MIN(T3.[ZipCode]) as CP, " +
                     "MIN(T4.[ReportCode]) as CodInf, SUM(T1.[VatSum]) as Iva, " +
                     "SUM(T1.[LineTotal] / 100 * T0.[DpmPrcnt]) as Anticipo, " +
                     "T2.[AcqstnRvrs] as Adquisicion " +
                     "FROM ODPO AS T0 INNER JOIN " +
                     "DPO1 AS T1 ON T1.DocEntry = T0.DocEntry INNER JOIN " +
                     "OVTG AS T2 ON T2.Code = T1.VatGroup INNER JOIN " +
                     "OCRD AS T3 ON T3.CardCode = T0.CardCode LEFT OUTER JOIN " +
                     "OCST AS T5 ON T3.Country = T5.Country AND T3.State1 = T5.Code LEFT OUTER JOIN " +
                     "OCRY AS T4 ON T3.Country = T4.Code " +
                     "WHERE  ((T2.Code = 'EI') OR " +
                     "(T2.Code = 'EIT') OR " +
                     "(T2.Code = 'EX') OR " +
                     "(T2.Code = 'R0') OR " +
                     "(T2.Code = 'R1') OR " +
                     "(T2.Code = 'R2') OR " +
                     "(T2.Code = 'R3') OR " +
                     "(T2.Code = 'RA0') OR " +
                     "(T2.Code = 'RA1') OR " +
                     "(T2.Code = 'RA2') OR " +
                     "(T2.Code = 'RA3') OR " +
                     "(T2.Code = 'RE1') OR " +
                     "(T2.Code = 'RE2') OR " +
                     "(T2.Code = 'RE3') OR " +
                     "(T2.Code = 'RIN0') OR " +
                     "(T2.Code = 'RIN1') OR " +
                     "(T2.Code = 'RIN2') OR " +
                     "(T2.Code = 'RIN3') OR " +
                     "(T2.Code = 'A1') OR " +
                     "(T2.Code = 'A2') OR " +
                     "(T2.Code = 'A3') OR " +
                     "(T2.Code = 'AI') OR " +
                     "(T2.Code = 'I0') OR " +
                     "(T2.Code = 'I1') OR " +
                     "(T2.Code = 'I2') OR " +
                     "(T2.Code = 'I3') OR " +
                     "(T2.Code = 'IBI0') OR " +
                     "(T2.Code = 'IBI1') OR " +
                     "(T2.Code = 'IBI2') OR " +
                     "(T2.Code = 'IBI3') OR " +
                     "(T2.Code = 'ND0') OR " +
                     "(T2.Code = 'ND1') OR " +
                     "(T2.Code = 'ND2') OR " +
                     "(T2.Code = 'ND3') OR " +
                     "(T2.Code = 'S0') OR " +
                     "(T2.Code = 'S1') OR " +
                     "(T2.Code = 'S2') OR " +
                     "(T2.Code = 'S3') OR " +
                     "(T2.Code = 'S4') OR " +
                     "(T2.Code = 'S5') OR " +
                     "(T2.Code = 'S6') OR " +
                     "(T2.Code = 'SA0') OR " +
                     "(T2.Code = 'SA1') OR " +
                     "(T2.Code = 'SA2') OR " +
                     "(T2.Code = 'SA3') OR " +
                     "(T2.Code = 'SI0') OR " +
                     "(T2.Code = 'SI1') OR " +
                     "(T2.Code = 'SI2') OR " +
                     "(T2.Code = 'SI3') OR " +
                     "(T2.Code = 'SIN0') OR " +
                     "(T2.Code = 'SIN1') OR " +
                     "(T2.Code = 'SIN2') OR " +
                     "(T2.Code = 'SIN3') OR " +
                     "(T2.Code = 'SV')) " +
                     "AND (T0.DOCDATE   >= '" + DesdeFecha.Substring(6, 4) + "/" + DesdeFecha.Substring(3, 2) + "/" + DesdeFecha.Substring(0, 2) + "' " +
                     "AND T0.DOCDATE    <= '" + HastaFecha.Substring(6, 4) + "/" + HastaFecha.Substring(3, 2) + "/" + HastaFecha.Substring(0, 2) + "' " +
                     "AND T3.CARDCODE   >= '" + DesdeIC + "' " +
                     "AND T3.CARDCODE   <= '" + HastaIC + "') ";
            if (ValorAcreedor && ValorDeudor)
            {
                StrSQL = StrSQL + "AND (T3.CardType = 'C' OR T3.CardType = 'S') ";
            }
            else
            {
                if (ValorDeudor)
                {
                    StrSQL = StrSQL + "AND T3.CardType = 'C' ";
                }
                if (ValorAcreedor)
                {
                    StrSQL = StrSQL + "AND T3.CardType = 'S' ";
                }
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkExtr").Specific;
            if (oCheckBox.Checked == false) //sólo españoles
            {
                StrSQL = StrSQL + "AND T3.Country = 'ES' ";
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkICRet").Specific;
            if (oCheckBox.Checked == false && StrRetencion == "Y") //sólo incluimos a los que no tienen retencion, si el sistema asi lo tiene
            {
                StrSQL = StrSQL + "AND (T3.WTLiable='N' OR T3.WTLiable='') ";
            }
            StrSQL = StrSQL + "GROUP BY T3.CardCode, T0.[Indicator], T2.AcqstnRvrs";

            SqlDataAdapter daSQLAnticipoProv = new SqlDataAdapter(StrSQL, csVariablesGlobales.conAddon);
            System.Data.DataTable dtSQLAnticipoProv = new System.Data.DataTable();
            DataRow drAnticipoProv;
            daSQLAnticipoProv.Fill(dtSQLAnticipoProv);
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            if (dtSQLAnticipoProv.Rows.Count > 0)
            {
                for (int i = 0; i < dtSQLAnticipoProv.Rows.Count; i++)
                {
                    oStaticText.Caption = "Cargando Factura " + Convert.ToUInt32(i + 1) + " de " + dtSQLAnticipoProv.Rows.Count;
                    System.Windows.Forms.Application.DoEvents();
                    drAnticipoProv = dtSQLAnticipoProv.Rows[i];
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = csUtilidades.DameValor("[@SIA_MOD347]", "Max(Code)", "");
                    if (StrCodMax == "")
                    {
                        StrCodMax = "0";
                    }
                    //dblSaldo = Convert.ToDouble(drFactura["Total"].ToString());
                    StrCodMax = Convert.ToString(Convert.ToInt32(StrCodMax) + 1);
                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    pReTabla.Code = StrCodMax;
                    pReTabla.Name = StrCodMax;
                    pReTabla.UserFields.Fields.Item("U_CodIC").Value = drAnticipoProv["CodIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NomIC").Value = drAnticipoProv["NomIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_TipIC").Value = drAnticipoProv["TipIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_NifIC").Value = drAnticipoProv["NifIC"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(drAnticipoProv["Base"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Prov").Value = drAnticipoProv["Prov"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Pais").Value = drAnticipoProv["Pais"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Indic").Value = drAnticipoProv["Indic"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CP").Value = drAnticipoProv["CP"].ToString();
                    pReTabla.UserFields.Fields.Item("U_CodInf").Value = drAnticipoProv["CodInf"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Iva").Value = drAnticipoProv["Iva"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Anticip").Value = Convert.ToDouble(drAnticipoProv["Anticipo"].ToString());
                    pReTabla.UserFields.Fields.Item("U_Adquisi").Value = drAnticipoProv["Adquisicion"].ToString();
                    pReTabla.UserFields.Fields.Item("U_Total").Value = Convert.ToDouble(drAnticipoProv["Base"].ToString()) + Convert.ToDouble(drAnticipoProv["Iva"].ToString());
                    pReTabla.UserFields.Fields.Item("U_TipDoc").Value = "ANTICIPO";

                    if (pReTabla.Add() != 0)
                    {
                        csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                        return;
                    }
                }
            }
            else
            {
                oStaticText.Caption = "No hay cargos de anticipo que tratar.";
            }
            #endregion
        }

        private void BorrarDatosInferioresAImporte()
        {
            SqlCommand cmdBorrar;
            SAPbobsCOM.UserTable pReTabla;
            SAPbobsCOM.Recordset oRecordSet;
            SqlDataAdapter daSql;
            System.Data.DataTable dtSQL = new System.Data.DataTable();
            SAPbouiCOM.EditText oEditText;
            int i;
            string StrSql;

            DataRow drSQL;

            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtMaxImp").Specific;

            StrSql = "SELECT U_CodIC, SUM(U_Total) AS Total " +
                    "FROM [@SIA_MOD347] " +
                    "GROUP BY U_CodIC " +
                    "HAVING  (SUM(U_Total) < '" + Convert.ToDouble(oEditText.String).ToString().Replace(",", ".") + "')";
            oRecordSet = (SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            daSql = new System.Data.SqlClient.SqlDataAdapter(StrSql, csVariablesGlobales.conAddon);
            daSql.Fill(dtSQL);
            if (dtSQL.Rows.Count > 0)
            {
                for (i = 0; i < dtSQL.Rows.Count; i++)
                {
                    pReTabla = csVariablesGlobales.oCompany.UserTables.Item("SIA_MOD347");
                    drSQL = dtSQL.Rows[i];
                    StrSql = "DELETE FROM [@SIA_MOD347] " +
                            "WHERE U_CodIC = '" + Convert.ToString(drSQL["U_CodIC"].ToString()) + "'";
                    cmdBorrar = new SqlCommand(StrSql, csVariablesGlobales.conAddon);
                    if (cmdBorrar.ExecuteNonQuery() < 0)
                    {
                        //MsgBox("DOR.Al limpiar la tabla TEMPORAL");
                        return;
                    }

                    cmdBorrar = new SqlCommand(StrSql, csVariablesGlobales.conAddon);
                    if (cmdBorrar.ExecuteNonQuery() < 0)
                    {
                        //MsgBox("DOR.Al limpiar la tabla TEMPORAL");
                        return;
                    }
                }
            }
        }

        private void GenerarFichero()
        {
            System.Data.DataTable dtSQL = new System.Data.DataTable();
            SqlDataAdapter daSql;
            int i;
            DataRow drSQL;
            SAPbouiCOM.EditText oEditText;
            SAPbouiCOM.StaticText oStaticText;
            string StrCodFactura = "FACTURA";
            string StrCodAbono = "ABONO";
            string StrCodDiario = "DIARIO";
            string StrCodAntic = "ANTICIPO";
            string StrCodCli;
            string StrAux;
            double DblFactura;
            double DblAbono;
            double DblDiario;
            double DblAntic;
            double DblSaldoC;
            string StrNifEmp;
            string StrNomEmp;
            string StrTelfEmp;
            string StrNifCli;
            string StrNomCli;
            string StrCpCli;
            string StrEjercicio;
            string StrFichero;
            string StrSql;
            csUtilidades Utilidades = new csUtilidades();

            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtSelFich").Specific;
            if (oEditText.String == "")
            {
                csVariablesGlobales.SboApp.MessageBox("Debe seleccionar un fichero", 1, "", "", "");
                oEditText.Active = true;
                return;
            }
            StrFichero = oEditText.String;

            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtEjer").Specific;
            if (oEditText.String == "")
            {
                csVariablesGlobales.SboApp.MessageBox("Debe introdocir un ejercicio", 1, "", "", "");
                oEditText.Active = true;
                return;
            }
            StrEjercicio = oEditText.String;
            System.IO.StreamWriter SW = new System.IO.StreamWriter(StrFichero, false);
            try
            {
                StrNifEmp = csUtilidades.DameValor("OADM", "TaxIdNum", "").ToUpper();
                StrNifEmp = StrNifEmp.Substring(2, StrNifEmp.Length - 2);// UCase(Mid(StrNifEmp, 3, Len(StrNifEmp)));
                StrNomEmp = csUtilidades.DameValor("OADM", "CompnyName", "").ToUpper();
                StrTelfEmp = csUtilidades.DameValor("OADM", "Phone1", "").ToUpper();

                //strFICHERO = TxtDestino.Text;

                //FileClose(1);
                //FileOpen(1, strFICHERO, OpenMode.Output);

                //Linea 1
                string StrLinea;
                StrLinea = "1" +
                            "347" + //Mid(StrLinea, 2, 3) = "347";
                            StrEjercicio + //Mid(StrLinea, 5, 4) = StrEjercicio;
                            StrNifEmp +//Mid(StrLinea, 9, 9) = StrNifEmp; // /*NIF EMPRESA*/
                            StrNomEmp.PadRight(40).Substring(0, 40) +// Mid(StrLinea, 18, 40) = Mid(StrNomEmp, 1, 40); // /*NOMBRE EMPRESA*/
                            "T" +// Mid(StrLinea, 58, 1) = "D";
                            StrTelfEmp.PadRight(9).Substring(0, 9) +// Mid(StrLinea, 59, 9) = Mid(StrTelfEmp, 1, 9); // /*TELEFONO EMPRESA*/
                            StrNomEmp.PadRight(40).Substring(0, 40) +// Mid(StrLinea, 68, 40) = Mid(StrNomEmp, 1, 40); // /*NOMBRE EMPRESA*/
                            "3480000000001" + //Mid(StrLinea, 108, 13) = "3480000000001"; // /*EURO*/
                            "  " +//Mid(StrLinea, 121, 2) = "  ";
                            "0000000000000" +//Mid(StrLinea, 123, 13) = "0000000000000";
                            "000000000000000000000000000000000000000000000000" +//Mid(StrLinea, 136, 48) = "000000000000000000000000000000000000000000000000";
                            "                                                                   ";
                SW.WriteLine(StrLinea); //PrintLine(1, StrLinea);

                StrSql = "SELECT DISTINCT U_CodIC, U_TipIC FROM [@SIA_MOD347] ORDER BY U_TipIC, U_CodIC";
                daSql = new System.Data.SqlClient.SqlDataAdapter(StrSql, csVariablesGlobales.conAddon);
                daSql.Fill(dtSQL);
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
                if (dtSQL.Rows.Count > 0)
                {
                    for (i = 0; i < dtSQL.Rows.Count; i++)
                    {
                        drSQL = dtSQL.Rows[i];
                        oStaticText.Caption = "Cargando datos Interlocutor " + drSQL["U_CodIC"].ToString();// +"  " + i + 1 + " de " + dtSQL.Rows.Count;
                        System.Windows.Forms.Application.DoEvents();
                        StrCodCli = drSQL["U_CodIC"].ToString();
                        DblSaldoC = 0;
                        //FACTURA
                        StrAux = csUtilidades.DameValor("[@SIA_MOD347]", "SUM(U_Total)", "U_CodIC ='" + StrCodCli +
                                             "' AND U_TipDoc = '" + StrCodFactura + "' GROUP BY U_CodIC, U_TipDoc");
                        if (StrAux == "")
                        {
                            DblFactura = 0;
                        }
                        else
                        {
                            DblFactura = Convert.ToDouble(StrAux.Replace(",0", "."));
                        }
                        //ABONO
                        StrAux = csUtilidades.DameValor("[@SIA_MOD347]", "SUM(U_Total)", "U_CodIC ='" + StrCodCli +
                                             "' AND U_TipDoc = '" + StrCodAbono + "' GROUP BY U_CodIC, U_TipDoc");
                        if (StrAux == "")
                        {
                            DblAbono = 0;
                        }
                        else
                        {
                            DblAbono = Convert.ToDouble(StrAux.Replace(",", "."));
                        }
                        //DIARIO
                        StrAux = csUtilidades.DameValor("[@SIA_MOD347]", "SUM(U_Total)", "U_CodIC ='" + StrCodCli +
                                             "' AND U_TipDoc = '" + StrCodDiario + "' GROUP BY U_CodIC, U_TipDoc");
                        if (StrAux == "")
                        {
                            DblDiario = 0;
                        }
                        else
                        {
                            DblDiario = Convert.ToDouble(StrAux.Replace(",0", "."));
                        }
                        //ANTICIPO
                        StrAux = csUtilidades.DameValor("[@SIA_MOD347]", "SUM(U_Total)", "U_CodIC ='" + StrCodCli +
                                             "' AND U_TipDoc = '" + StrCodAntic + "' GROUP BY U_CodIC, U_TipDoc");
                        if (StrAux == "")
                        {
                            DblAntic = 0;
                        }
                        else
                        {
                            DblAntic = Convert.ToDouble(StrAux.Replace(",0", "."));
                        }
                        DblSaldoC = DblFactura + DblAntic + DblDiario + DblAbono;
                        DblSaldoC = Convert.ToDouble(DblSaldoC.ToString("0.00"));
                        DblSaldoC = DblSaldoC * 100;
                        StrNifCli = csUtilidades.DameValor("[@SIA_MOD347]", "U_NifIC", "U_CodIC ='" + StrCodCli + "'");
                        if (StrNifCli.Length != 0)
                        {
                            StrNifCli = StrNifCli.Substring(2, StrNifCli.Length - 2);//UCase(Mid(StrNifCli, 3, Len(StrNifCli)));
                        }
                        else
                        {
                            StrNifCli = "         ";
                            csVariablesGlobales.SboApp.MessageBox("Falta el N.I.F. del I.C. " + StrCodCli, 1, "", "", "");
                        }
                        StrNomCli = csUtilidades.DameValor("[@SIA_MOD347]", "U_NomIC", "U_CodIC ='" + StrCodCli + "'");
                        StrNomCli = StrNomCli.ToUpper().PadRight(40).Substring(0, 40);
                        StrCpCli = csUtilidades.DameValor("[@SIA_MOD347]", "U_CP", "U_CodIC ='" + StrCodCli + "'");
                        if (StrCpCli.Length != 0)
                        {
                            StrCpCli = StrCpCli.ToUpper().Substring(0, 2);
                        }
                        else
                        {
                            StrCpCli = "  ";
                            csVariablesGlobales.SboApp.MessageBox("Falta el C.P. del I.C. " + StrCodCli, 1, "", "", "");
                        }

                        //StrLinea = Space(250);
                        StrLinea = "2" +
                                "347" +
                                StrEjercicio +
                                StrNifEmp + // /*NIF EMPRESA*/
                                StrNifCli +
                                "         " +
                                StrNomCli + // /*NOMBRE CLIENTE*/
                                "D" +
                                StrCpCli + // /*CP CLIENTE*/
                                "   ";
                        if (csUtilidades.DameValor("[@SIA_MOD347]", "U_TipIC", "U_CodIC ='" + StrCodCli + "'") == "C")
                        {
                            StrLinea = StrLinea = StrLinea + "B"; //CLIENTE B
                        }
                        else
                        {
                            StrLinea = StrLinea = StrLinea + "A"; //PROVEEDOR A
                        }
                        StrLinea = StrLinea + Convert.ToDouble(DblSaldoC).ToString("000000000000000") +
                                    "                                                                                                                                                         ";
                        SW.WriteLine(StrLinea);
                    }
                }
                SW.Close();
                //FileClose(1);
                csVariablesGlobales.SboApp.MessageBox("Fichero creado correctamente", 1, "", "", "");
                oStaticText.Caption = "Fichero creado";
                System.Windows.Forms.Application.DoEvents();
            }
            catch (Exception ex)
            {
                SW.Close();
                MessageBox.Show(ex.Message);
                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
            }
        }

        public void Imprimir347(string Menu, string Impreso)
        {
            SAPbouiCOM.EditText oEditText;
            //SBOUtilyArch.cLanzarReport LanzarReport;
            string StrParm1 = "";
            string StrParm2 = "";
            string StrParm3 = "";
            string StrParm4 = "";
            string StrParm5 = "";
            string StrFormula = "";
            string StrBD;
            string StrUsuAdo;
            DateTime Fecha;
            //LanzarReport = new SBOUtilyArch.cLanzarReport();
            StrBD = csVariablesGlobales.oCompany.ToString();
            StrUsuAdo = csVariablesGlobales.oCompany.DbUserName;
            oForm = csVariablesGlobales.SboApp.Forms.Item("FrmModelo347");
            string StrReport = csVariablesGlobales.StrRutRep + @"\REPORT\" + Impreso;

            csUtilidades Utilidades = new csUtilidades();
            switch (Impreso.Substring(0, 5))
            {
                case "Carta":
                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtMaxImp").Specific;
                    StrParm1 = "{@Importe}=" + oEditText.String;
                    StrParm2 = "{@Provincia}=" + csUtilidades.DameValor("ADM1", "County", "");
                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecCar").Specific;
                    Fecha = Convert.ToDateTime(oEditText.String);
                    StrParm3 = "{@FechaCarta}=" + Fecha.ToLongDateString();
                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtEjer").Specific;
                    StrParm4 = "{@Año}=" + oEditText.String;
                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFir").Specific;
                    StrParm5 = "{@Firmante}=" + oEditText.String;
                    StrFormula = "";
                    break;
                case "Infor":
                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtDesdeIC").Specific;
                    StrParm1 = "{@ClienteD}=" + oEditText.String;
                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtHastaIC").Specific;
                    StrParm2 = "{@ClienteH}=" + oEditText.String;
                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtDesFec").Specific;
                    StrParm3 = "{@FecD}=" + oEditText.String;
                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtHasFec").Specific;
                    StrParm4 = "{@FecH}=" + oEditText.String;
                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtMaxImp").Specific;
                    StrParm5 = "{@Importe}=" + oEditText.String;
                    StrFormula = "";
                    break;
            }
            switch (Menu)
            {
                case "519": // Vista Previa
                    //LanzarReport.Informe(ref StrReport, ref csVariablesGlobales.StrServidor, ref csVariablesGlobales.StrBaseDatos, ref csVariablesGlobales.StrUserConex, ref csVariablesGlobales.StrPassConex, StrFormula, StrParm1, StrParm2, StrParm3, StrParm4, StrParm5, csVariablesGlobales.StrSubRep1, csVariablesGlobales.StrSubRep2, csVariablesGlobales.StrSubRep3, csVariablesGlobales.StrSubRep4);
                    break;
                case "520": // Imprimir
                    //LanzarReport.Imprimir(ref StrReport, ref csVariablesGlobales.StrServidor, ref csVariablesGlobales.StrBaseDatos, ref csVariablesGlobales.StrUserConex, ref csVariablesGlobales.StrPassConex, StrFormula, StrParm1, StrParm2, StrParm3, StrParm4, StrParm5, csVariablesGlobales.StrSubRep1, csVariablesGlobales.StrSubRep2, csVariablesGlobales.StrSubRep3, csVariablesGlobales.StrSubRep4);
                    break;
            }
        }

        private bool ValidarGeneracion()
        {
            try
            {
                SAPbouiCOM.EditText oTxt = null;
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtDesFec").Specific;
                if (oTxt.String == "")
                {
                    oTxt.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("La fecha desde debe existir", 1, "Ok", "", "");
                    return false;
                }
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtHasFec").Specific;
                if (oTxt.String == "")
                {
                    oTxt.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("La fecha de hasta debe existir", 1, "Ok", "", "");
                    return false;
                }
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtSelFich").Specific;
                if (oTxt.String == "")
                {
                    oTxt.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("El fichero de destino debe exitir", 1, "Ok", "", "");
                    return false;
                }
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtDesdeIC").Specific;
                if (oTxt.String == "")
                {
                    oTxt.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("El campo Desde IC no esta informado", 1, "Ok", "", "");
                    return false;
                }
                oTxt = (SAPbouiCOM.EditText)oForm.Items.Item("txtHastaIC").Specific;
                if (oTxt.String == "")
                {
                    oTxt.Active = true;
                    csVariablesGlobales.SboApp.MessageBox("El campo Hasta IC no esta informado", 1, "Ok", "", "");
                    return false;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
