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
    class csFrmModelo349
    {
        private SAPbouiCOM.Form oForm;
        SAPbouiCOM.Item oItem = null;
        SAPbouiCOM.Button oButton = null;
        SAPbouiCOM.OptionBtn oOptionBtn = null;
        SAPbouiCOM.CheckBox oCheckBox = null;
        SAPbouiCOM.StaticText oStaticText = null;
        SAPbouiCOM.EditText oEditText = null;

        public void CargarFormulario()
        {
            CrearFormulario349();
            oForm.Visible = true;
            //csUtilidades Utilidades = new csUtilidades();
            csUtilidades.SaveAsXml(oForm, "FrmModelo349.xml", "");
            // events handled by SBO_Application_ItemEvent
            csVariablesGlobales.SboApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(FrmModelo349_ItemEvent);
            // events handled by SBO_Application_AppEvent
            csVariablesGlobales.SboApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(FrmModelo349_AppEvent);
        }

        private void CrearFormulario349()
        {
            int BaseLeft = 0;
            int BaseTop = 0;
            //SAPbouiCOM.Item oItem = null;
            //SAPbouiCOM.Button oButton = null;
            //SAPbouiCOM.Folder oFolder = null;
            //SAPbouiCOM.OptionBtn oOptionBtn = null;
            //SAPbouiCOM.CheckBox oCheckBox = null;
            //SAPbouiCOM.ComboBox oComboBox = null;
            //SAPbouiCOM.StaticText oStaticText = null;
            //SAPbouiCOM.EditText oEditText = null;

            #region Formulario
            SAPbouiCOM.FormCreationParams oCreationParams = null;
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "FrmModelo349";
            oCreationParams.FormType = "SIASL10006";

            oForm = csVariablesGlobales.SboApp.Forms.AddEx(oCreationParams);

            // set the form properties
            oForm.Title = "Modelo 349";
            oForm.Left = 300;
            oForm.ClientWidth = 500;
            oForm.Top = 100;
            oForm.ClientHeight = 380;
            //oForm.EnableMenu("519", true);
            //oForm.EnableMenu("520", true);
            #endregion

            #region Data Sources
            //  Adding a User Data Source
            oForm.DataSources.UserDataSources.Add("dsDesdeIC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsHastaIC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsDeud", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("dsAcre", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("dsOptPer", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("dsInICRet", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("dsEjer", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("dsSelFich", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsOptDes", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("dsDesdeFec", SAPbouiCOM.BoDataType.dt_DATE, 254);
            oForm.DataSources.UserDataSources.Add("dsHastaFec", SAPbouiCOM.BoDataType.dt_DATE, 254);
            oForm.DataSources.UserDataSources.Add("dsOptDesF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
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
            #region Periodo
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
            oItem.Height = 80;

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblPer", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 70;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Periodo";

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
            oItem.Height = 80;

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
            BaseTop = 175;

            //****************************
            // Adding a Rectangle Item
            // for cosmetic purposes only
            //****************************
            oItem = oForm.Items.Add("Rect4", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = BaseLeft;
            oItem.Width = 480;
            oItem.Top = BaseTop;
            oItem.Height = 80;

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
            BaseTop = 265;

            //****************************
            // Adding a Rectangle Item
            // for cosmetic purposes only
            //****************************
            oItem = oForm.Items.Add("Rect5", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = BaseLeft;
            oItem.Width = 235;
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
            #region Destino Fichero
            BaseLeft = 250;
            BaseTop = 265;

            //****************************
            // Adding a Rectangle Item
            // for cosmetic purposes only
            //****************************
            oItem = oForm.Items.Add("Rect6", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = BaseLeft;
            oItem.Width = 235;
            oItem.Top = BaseTop;
            oItem.Height = 60;

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblDesFic", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 110;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Destino Fichero";

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

            // We get automatic event handling for
            // the Ok and Cancel Buttons by setting
            // their UIDs to 1 and 2 respectively

            oItem = oForm.Items.Add("btnGen", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5;
            oItem.Width = 65;
            oItem.Top = 345;
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
            oItem.Top = 345;
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
            oItem.Top = 345;
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
            oItem.Top = 345;
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
            oItem.Top = 205;
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
            //BaseLeft = 10;
            //BaseTop = 100;

            //// Date Picker
            ////__________________________________________________________________________________________
            //////*************************
            ////// Adding a Text Edit item
            //////*************************

            //oItem = oForm.Items.Add("txtDesFec", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            //oItem.Left = BaseLeft + 120;
            //oItem.Width = 100;
            //oItem.Top = BaseTop;
            //oItem.Height = 14;

            //oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            ////// bind the text edit item to the defined used data source
            //oEditText.DataBind.SetBound(true, "", "dsDesdeFec");

            //////**********************************
            ////// Adding Static Text item
            //////**********************************

            //oItem = oForm.Items.Add("lblDesFec", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            //oItem.Left = BaseLeft;
            //oItem.Width = 100;
            //oItem.Top = BaseTop;
            //oItem.Height = 14;

            //oItem.LinkTo = "txtDesFec";
            //oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            //oStaticText.Caption = "Desde Fecha";

            //BaseLeft = 0;
            //BaseTop = 0;
            #endregion
            #region Hasta Fecha
            //BaseLeft = 10;
            //BaseTop = 120;

            //// Date Picker
            ////__________________________________________________________________________________________
            //////*************************
            ////// Adding a Text Edit item
            //////*************************

            //oItem = oForm.Items.Add("txtHasFec", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            //oItem.Left = BaseLeft + 120;
            //oItem.Width = 100;
            //oItem.Top = BaseTop;
            //oItem.Height = 14;

            //oEditText = (SAPbouiCOM.EditText)oItem.Specific;

            ////// bind the text edit item to the defined used data source
            //oEditText.DataBind.SetBound(true, "", "dsHastaFec");

            //////**********************************
            ////// Adding Static Text item
            //////**********************************

            //oItem = oForm.Items.Add("lblHasFec", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            //oItem.Left = BaseLeft;
            //oItem.Width = 100;
            //oItem.Top = BaseTop;
            //oItem.Height = 14;

            //oItem.LinkTo = "txtHasFec";
            //oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            //oStaticText.Caption = "Hasta Fecha";

            //BaseLeft = 0;
            //BaseTop = 0;
            #endregion
            #region Ejercicio
            //*************************
            // Adding a Text Edit item
            //*************************
            BaseLeft = 10;
            BaseTop = 190;

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
            BaseTop = 210;

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
            #region CheckBox Inc IC con Retención
            BaseLeft = 260;
            BaseTop = 100;

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
            BaseTop = 240;
            BaseLeft = 60;
            oItem = oForm.Items.Add("lblProces", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 250;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Pro";
            #endregion

            #region Option Button
            #region Option Pantalla
            BaseTop = 280;
            BaseLeft = 30;

            oItem = oForm.Items.Add("optPorPan", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 19;

            oOptionBtn = ((SAPbouiCOM.OptionBtn)(oItem.Specific));

            oOptionBtn.Caption = "Por Pantalla";
            oOptionBtn.DataBind.SetBound(true, "", "dsOptDes");
            //oOptionBtn.Selected = true;

            //BaseTop = 0;
            //BaseLeft = 0;
            #endregion
            #region Option Impresora
            //BaseTop = 300;
            //BaseLeft = 150;

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
            #region Option Telemático
            BaseTop = 280;
            BaseLeft = 260;

            oItem = oForm.Items.Add("optTel", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 19;

            oOptionBtn = ((SAPbouiCOM.OptionBtn)(oItem.Specific));

            oOptionBtn.Caption = "Telemático";
            oOptionBtn.DataBind.SetBound(true, "", "dsOptDesF");
            //oOptionBtn.Selected = true;

            //BaseTop = 0;
            //BaseLeft = 0;
            #endregion
            #region Option Disquete
            //BaseTop = 300;
            //BaseLeft = 150;

            oItem = oForm.Items.Add("optDis", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop + 19;
            oItem.Height = 19;

            oOptionBtn = ((SAPbouiCOM.OptionBtn)(oItem.Specific));

            oOptionBtn.Caption = "Disquete";
            oOptionBtn.GroupWith("optTel");
            oOptionBtn.DataBind.SetBound(true, "", "dsOptDesF");

            BaseTop = 0;
            BaseLeft = 0;
            #endregion
            #region Option Anual
            BaseTop = 100;
            BaseLeft = 30;

            oItem = oForm.Items.Add("optAnual", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop;
            oItem.Height = 19;

            oOptionBtn = ((SAPbouiCOM.OptionBtn)(oItem.Specific));

            oOptionBtn.Caption = "Anual";
            oOptionBtn.DataBind.SetBound(true, "", "dsOptPer");
            //oOptionBtn.Selected = true;

            //BaseTop = 0;
            //BaseLeft = 0;
            #endregion
            #region Option Trimestre 1º
            //BaseTop = 300;
            //BaseLeft = 150;

            oItem = oForm.Items.Add("optPriTri", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop + 19;
            oItem.Height = 19;

            oOptionBtn = ((SAPbouiCOM.OptionBtn)(oItem.Specific));

            oOptionBtn.Caption = "Trimestre 1º";
            oOptionBtn.GroupWith("optAnual");
            oOptionBtn.DataBind.SetBound(true, "", "dsOptPer");

            //BaseTop = 0;
            //BaseLeft = 0;
            #endregion
            #region Option Trimestre 2º
            //BaseTop = 300;
            //BaseLeft = 150;

            oItem = oForm.Items.Add("optSegTri", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            oItem.Left = BaseLeft + 100;
            oItem.Width = 100;
            oItem.Top = BaseTop + 19;
            oItem.Height = 19;

            oOptionBtn = ((SAPbouiCOM.OptionBtn)(oItem.Specific));

            oOptionBtn.Caption = "Trimestre 2º";
            oOptionBtn.GroupWith("optPriTri");
            oOptionBtn.DataBind.SetBound(true, "", "dsOptPer");

            //BaseTop = 0;
            //BaseLeft = 0;
            #endregion
            #region Option Trimestre 3º
            //BaseTop = 300;
            //BaseLeft = 150;

            oItem = oForm.Items.Add("optTerTri", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            oItem.Left = BaseLeft;
            oItem.Width = 100;
            oItem.Top = BaseTop + 38;
            oItem.Height = 19;

            oOptionBtn = ((SAPbouiCOM.OptionBtn)(oItem.Specific));

            oOptionBtn.Caption = "Trimestre 3º";
            oOptionBtn.GroupWith("optSegTri");
            oOptionBtn.DataBind.SetBound(true, "", "dsOptPer");

            //BaseTop = 0;
            //BaseLeft = 0;
            #endregion
            #region Option Trimestre 4º
            //BaseTop = 300;
            //BaseLeft = 150;

            oItem = oForm.Items.Add("optCuaTri", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            oItem.Left = BaseLeft + 100;
            oItem.Width = 100;
            oItem.Top = BaseTop + 38;
            oItem.Height = 19;

            oOptionBtn = ((SAPbouiCOM.OptionBtn)(oItem.Specific));

            oOptionBtn.Caption = "Trimestre 4º";
            oOptionBtn.GroupWith("optTerTri");
            oOptionBtn.DataBind.SetBound(true, "", "dsOptPer");

            //BaseTop = 0;
            //BaseLeft = 0;
            #endregion
            //oForm.PaneLevel = 1;
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

        public void FrmModelo349_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            csUtilidades Utilidades = new csUtilidades();
            BubbleEvent = true;
            SAPbouiCOM.StaticText oStaticText;
            SAPbouiCOM.OptionBtn oOptionBtn;

            if (pVal.FormUID == "FrmModelo349")
            {

                //oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);

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
                                    Genera349();
                                    oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
                                    oStaticText.Caption = "Filtrando registros aptos para el 349";
                                    //BorrarDatosInferioresAImporte();
                                    GeneraFichero();
                                    oStaticText.Caption = "Proceso Terminado";
                                    oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optPorPan").Specific;
                                    if (oOptionBtn.Selected)
                                    {
                                        Imprimir349(csVariablesGlobales.MenuImprimirPorPantalla, "Informe349.rpt");
                                    }
                                    else
                                    {
                                        Imprimir349(csVariablesGlobales.MenuImprimirPorImpresora, "Informe349.rpt");
                                    }
                                    BubbleEvent = false;
                                }
                            }
                            if (pVal.ItemUID == "btnInf" && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                            {
                                oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optPorPan").Specific;
                                if (oOptionBtn.Selected)
                                {
                                    Imprimir349(csVariablesGlobales.MenuImprimirPorPantalla, "Informe349.rpt");
                                }
                                else
                                {
                                    Imprimir349(csVariablesGlobales.MenuImprimirPorImpresora, "Informe349.rpt");
                                }
                                BubbleEvent = false;
                            }
                            //if (pVal.ItemUID == "btnCar" && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                            //{
                            //    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtFecCar").Specific;
                            //    if (oEditText.String != "")
                            //    {
                            //        oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optPorPan").Specific;
                            //        if (oOptionBtn.Selected)
                            //        {
                            //            Imprimir347(csVariablesGlobales.sMenuPant, "Carta347.rpt");
                            //        }
                            //        else
                            //        {
                            //            Imprimir347(csVariablesGlobales.sMenuImpr, "Carta347.rpt");
                            //        }
                            //    }
                            //    else
                            //    {
                            //        oEditText.Active = true;
                            //        csVariablesGlobales.SboApp.MessageBox("La fecha de la carta es obligatoria", 1, "Ok", "", "");
                            //    }
                            //    BubbleEvent = false;
                            //}
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

        private void FrmModelo349_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
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
            oForm = csVariablesGlobales.SboApp.Forms.Item("FrmModelo349");
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtEjer").Specific;
            oEditText.String = Convert.ToString(DateTime.Today.Year);
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optPorPan").Specific;
            oOptionBtn.Selected = true;
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optTel").Specific;
            oOptionBtn.Selected = true;
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optAnual").Specific;
            oOptionBtn.Selected = true;
        }

        private bool ValidarGeneracion()
        {
            try
            {
                SAPbouiCOM.EditText oTxt = null;
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

        private void Genera349()
        {
            #region Declaraciones
            csUtilidades Utilidades = new csUtilidades();
            SqlDataAdapter daSql;
            System.Data.DataTable dtSql;
            int i;
            DataRow drSql;
            SAPbobsCOM.UserTable oUserTable = null;
            string StrCodMax;
            double dblSaldo;
            double IvaPortes;
            double BasePortes;
            string StrRetencion; //Y si tienen en cuenta retenciones el sistema
            string StrRetCom;
            bool ValorDeudor;
            bool ValorAcreedor;
            string DesdeFecha = "";
            string HastaFecha = "";
            string DesdeIC = "";
            string HastaIC = "";
            StrRetencion = csUtilidades.DameValor("OADM", "SHandleWT", "");
            StrRetCom = csUtilidades.DameValor("OADM", "pHandleWT", "");
            string StrSql;
            #endregion
            #region Asignación Variables
            oUserTable = csVariablesGlobales.oCompany.UserTables.Item("SIA_MOD349");
            StrRetencion = csUtilidades.DameValor("OADM", "SHandleWT", "");
            StrRetCom = csUtilidades.DameValor("OADM", "pHandleWT", "");
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtEjer").Specific;
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optAnual").Specific;
            if (oOptionBtn.Selected)
            {
                DesdeFecha = "01/01/" + oEditText.String;
                HastaFecha = "31/12/" + oEditText.String;
            }
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optPriTri").Specific;
            if (oOptionBtn.Selected)
            {
                DesdeFecha = "01/01/" + oEditText.String;
                HastaFecha = "31/03/" + oEditText.String;
            }
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optSegTri").Specific;
            if (oOptionBtn.Selected)
            {
                DesdeFecha = "01/04/" + oEditText.String;
                HastaFecha = "30/06/" + oEditText.String;
            }
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optTerTri").Specific;
            if (oOptionBtn.Selected)
            {
                DesdeFecha = "01/07/" + oEditText.String;
                HastaFecha = "30/09/" + oEditText.String;
            }
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optCuaTri").Specific;
            if (oOptionBtn.Selected)
            {
                DesdeFecha = "01/10/" + oEditText.String;
                HastaFecha = "31/12/" + oEditText.String;
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkDeud").Specific;
            ValorDeudor = oCheckBox.Checked;
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkAcre").Specific;
            ValorAcreedor = oCheckBox.Checked;
            //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtDesFec").Specific;
            //DesdeFecha = oEditText.String;
            //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtHasFec").Specific;
            //HastaFecha = oEditText.String;
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtDesdeIC").Specific;
            DesdeIC = oEditText.String;
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtHastaIC").Specific;
            HastaIC = oEditText.String;
            #endregion

            //Limpiar tabla Modelo 349
            StrSql = "DELETE FROM [@SIA_MOD349]";
            SqlCommand cmdBorrar = new SqlCommand(StrSql, csVariablesGlobales.conAddon);
            cmdBorrar.ExecuteNonQuery();
            cmdBorrar = null;

            #region CARGO FACTURA
            //   Rst = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            StrSql = "SELECT T3.CardCode AS CodIC, T3.CardName AS NomIC, " +
                     "T3.CardType AS TipIC, T3.LicTradNum AS NifIC, " +
                     "SUM(T0.DocTotal - T0.VatSum) AS Base, " +
                     "T5.Name AS Prov, T4.Name AS Pais, " +
                     "T0.[Indicator] AS Indic, T3.ZipCode AS CP, " +
                     "MIN(T4.ReportCode) AS CodInf, " +
                     "SUM(T0.VatSum - T0.EquVatSum) AS Iva, " +
                     "SUM(T0.DocTotal) AS Total " +
                     "FROM OINV T0 INNER JOIN " +
                     "V_SIA_TipoIvaFactVent T1 ON T0.DocEntry = T1.DocEntry INNER JOIN " +
                     "OCRD T3 ON T3.CardCode = T0.CardCode LEFT OUTER JOIN " +
                     "OCRY T4 ON T3.Country = T4.Code LEFT OUTER JOIN " +
                     "OCST T5 ON T3.State1 = T5.Code " +
                     "WHERE (T0.DOCDATE >= '" + DesdeFecha.Substring(6, 4) + "/" + DesdeFecha.Substring(3, 2) + "/" + DesdeFecha.Substring(0, 2) + "' " +
                     "AND T0.DOCDATE    <= '" + HastaFecha.Substring(6, 4) + "/" + HastaFecha.Substring(3, 2) + "/" + HastaFecha.Substring(0, 2) + "' " +
                     "AND T3.CARDCODE   >= '" + DesdeIC + "' " +
                     "AND T3.CARDCODE   <= '" + HastaIC + "') " +
                     "AND T1.IsEC='Y' AND T1.Tipo='F' AND T3.VatStatus='E' ";//ES de la UE y sus facturas tb
            if (ValorAcreedor && ValorDeudor)
            {
                StrSql = StrSql + "AND (T3.CardType = 'C' OR T3.CardType = 'S') ";
            }
            else
            {
                if (ValorDeudor)
                {
                    StrSql = StrSql + "AND T3.CardType = 'C' ";
                }
                if (ValorAcreedor)
                {
                    StrSql = StrSql + "AND T3.CardType = 'S' ";
                }
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkICRet").Specific;
            if (oCheckBox.Checked == false && StrRetencion == "Y") //sólo incluimos a los que no tienen retencion, si el sistema asi lo tiene
            {
                StrSql = StrSql + "AND (T3.WTLiable='N' or T3.WTLiable='' or T3.WTLiable IS NULL) ";
            }

            StrSql = StrSql + "GROUP BY T3.CardCode, T0.[Indicator],T3.CardName,T3.CardType, " +
                              "T3.LicTradNum, T5.Name, T4.Name, T3.ZipCode";
            //Debug.Print(StrSql);

            daSql = new System.Data.SqlClient.SqlDataAdapter(StrSql, csVariablesGlobales.conAddon);
            dtSql = new System.Data.DataTable();
            daSql.Fill(dtSql);
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            if (dtSql.Rows.Count > 0)
            {
                for (i = 0; i < dtSql.Rows.Count; i++)
                {
                    oStaticText.Caption = "Cargando Factura " + Convert.ToUInt32(i + 1) + " de " + dtSql.Rows.Count;
                    System.Windows.Forms.Application.DoEvents();
                    drSql = dtSql.Rows[i];
                    //Label1.Text = "ALBARAN " & dr("DocNum").ToString
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = csUtilidades.DameValor("[@SIA_MOD349]", "Max(Code)", "");
                    if (StrCodMax == "")
                    {
                        StrCodMax = "0";
                    }



                    dblSaldo = Convert.ToDouble(drSql["Total"].ToString());
                    StrCodMax = Convert.ToString(Convert.ToInt64(StrCodMax) + 1);
                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    oUserTable.Code = StrCodMax;
                    oUserTable.Name = StrCodMax;
                    oUserTable.UserFields.Fields.Item("U_CodIC").Value = Convert.ToString(drSql["CodIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_NomIC").Value = Convert.ToString(drSql["NomIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_TipIC").Value = Convert.ToString(drSql["TipIC"].ToString());
                    //Debug.Print(Convert.ToString(drSql["NifIC"].ToString()));

                    oUserTable.UserFields.Fields.Item("U_NifIC").Value = "";
                    if (drSql["NifIC"].ToString() != "")
                    {
                        oUserTable.UserFields.Fields.Item("U_NifIC").Value = Convert.ToString(drSql["NifIC"].ToString());//.Substring(1, 20);
                    }
                    //oUserTable.UserFields.Fields.Item("U_NifIC").Value = Convert.ToString(drSql["NifIC"].ToString()).Substring(1, 20);

                    oUserTable.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(drSql["Base"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Prov").Value = Convert.ToString(drSql["Prov"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Pais").Value = Convert.ToString(drSql["Pais"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Indic").Value = Convert.ToString(drSql["Indic"].ToString());
                    oUserTable.UserFields.Fields.Item("U_CP").Value = Convert.ToString(drSql["CP"].ToString());
                    oUserTable.UserFields.Fields.Item("U_CodInf").Value = Convert.ToString(drSql["CodInf"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Total").Value = Convert.ToDouble(dblSaldo);
                    oUserTable.UserFields.Fields.Item("U_TipDoc").Value = Convert.ToString("FACTURA").ToString();

                    //oCompany.GetLastError(longError, strMsError);
                    if (oUserTable.Add() != 0)
                    {
                        MessageBox.Show(csVariablesGlobales.oCompany.GetLastErrorDescription());
                        return;
                    }
                }
            }
            #endregion
            #region CARGO ABONO
            //Rst = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            StrSql = "SELECT T3.CardCode AS CodIC, MIN(T3.CardName) AS NomIC, " +
                     "MIN(T3.CardType) AS TipIC, MIN(T3.LicTradNum) AS NifIC, " +
                     "SUM(T0.DocTotal - T0.VatSum)AS Base, " +
                     "MIN(T5.Name) AS Prov, MIN(T4.Name) AS Pais, " +
                     "T0.[Indicator] AS Indic, MIN(T3.ZipCode) AS CP, " +
                     " MIN(T4.ReportCode) AS CodInf, " +
                     "SUM(T0.VatSum - T0.EquVatSum) AS Iva, " +
                     "SUM(T0.DocTotal) AS Total " +
                     "FROM ORIN T0 INNER JOIN " +
                     "V_SIA_TipoIvaFactVent T1 ON T0.DocEntry = T1.DocEntry INNER JOIN " +
                     "OCRD T3 ON T3.CardCode = T0.CardCode LEFT OUTER JOIN " +
                     "OCST T5 ON T3.Country = T5.Country AND T3.State1 = T5.Code LEFT OUTER JOIN " +
                     "OCRY T4 ON T3.Country = T4.Code " +
                     "WHERE (T0.DOCDATE >= '" + DesdeFecha.Substring(6, 4) + "/" + DesdeFecha.Substring(3, 2) + "/" + DesdeFecha.Substring(0, 2) + "' " +
                     "AND T0.DOCDATE    <= '" + HastaFecha.Substring(6, 4) + "/" + HastaFecha.Substring(3, 2) + "/" + HastaFecha.Substring(0, 2) + "' " +
                     "AND T3.CARDCODE   >= '" + DesdeIC + "' " +
                     "AND T3.CARDCODE   <= '" + HastaIC + "') " +
                     "AND T1.IsEC='Y' AND T1.Tipo='A' AND T3.VatStatus='E' ";//ES de la UE y sus facturas tb
            if (ValorAcreedor && ValorDeudor)
            {
                StrSql = StrSql + "AND (T3.CardType = 'C' OR T3.CardType = 'S') ";
            }
            else
            {
                if (ValorDeudor)
                {
                    StrSql = StrSql + "AND T3.CardType = 'C' ";
                }
                if (ValorAcreedor)
                {
                    StrSql = StrSql + "AND T3.CardType = 'S' ";
                }
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkICRet").Specific;
            if (oCheckBox.Checked == false && StrRetencion == "Y") //sólo incluimos a los que no tienen retencion, si el sistema asi lo tiene
            {
                StrSql = StrSql + "AND (T3.WTLiable='N' or T3.WTLiable='' or T3.WTLiable IS NULL) ";
            }

            StrSql = StrSql + "GROUP BY T3.CardCode, T0.[Indicator]";

            daSql = new System.Data.SqlClient.SqlDataAdapter(StrSql, csVariablesGlobales.conAddon);
            dtSql = new System.Data.DataTable();
            daSql.Fill(dtSql);
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            if (dtSql.Rows.Count > 0)
            {
                for (i = 0; i < dtSql.Rows.Count; i++)
                {
                    oStaticText.Caption = "Cargando Abono " + Convert.ToUInt32(i + 1) + " de " + dtSql.Rows.Count;
                    System.Windows.Forms.Application.DoEvents();
                    drSql = dtSql.Rows[i];
                    //Label1.Text = "ALBARAN " & dr("DocNum").ToString
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = csUtilidades.DameValor("[@SIA_MOD349]", "Max(Code)", "");
                    if (StrCodMax == "")
                    {
                        StrCodMax = "0";
                    }

                    dblSaldo = Convert.ToDouble(drSql["Total"].ToString());
                    StrCodMax = Convert.ToString(Convert.ToInt64(StrCodMax) + 1);
                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    oUserTable.Code = StrCodMax;
                    oUserTable.Name = StrCodMax;
                    oUserTable.UserFields.Fields.Item("U_CodIC").Value = Convert.ToString(drSql["CodIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_NomIC").Value = Convert.ToString(drSql["NomIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_TipIC").Value = Convert.ToString(drSql["TipIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_NifIC").Value = Convert.ToString(drSql["NifIC"].ToString());
                    if ((drSql["Base"].ToString()) == "")
                    {
                        oUserTable.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(0);
                    }
                    else
                    {
                        oUserTable.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(drSql["Base"].ToString());
                    }

                    oUserTable.UserFields.Fields.Item("U_Prov").Value = Convert.ToString(drSql["Prov"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Pais").Value = Convert.ToString(drSql["Pais"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Indic").Value = Convert.ToString(drSql["Indic"].ToString());
                    oUserTable.UserFields.Fields.Item("U_CP").Value = Convert.ToString(drSql["CP"].ToString());
                    oUserTable.UserFields.Fields.Item("U_CodInf").Value = Convert.ToString(drSql["CodInf"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Total").Value = Convert.ToDouble(-dblSaldo);
                    oUserTable.UserFields.Fields.Item("U_TipDoc").Value = Convert.ToString("ABONO").ToString();

                    if (oUserTable.Add() != 0)
                    {
                        MessageBox.Show(csVariablesGlobales.oCompany.GetLastErrorDescription());
                        return;
                    }
                }
            }
            #endregion
            #region CARGO ANTICIPO
            //Rst = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            StrSql = " SELECT T3.[CardCode] as CodIC , MIN(T3.[CardName]) as NomIC, " +
                    "MIN(T3.[CardType]) as TipIC, MIN(T3.[LicTradNum]) as NifIC, " +
                    "SUM(T1.[LineTotal] / 100 * (100 - T0.[DiscPrcnt])) as Base, " +
                    "MIN(T5.[Name]) as Prov, MIN(T4.[Name]) as Pais, T0.[Indicator] as Indic, " +
                    "MIN(T3.[ZipCode]) as CP, MIN(T4.[ReportCode]) as CodInf, " +
                    "SUM(T1.[VatSum]) as Iva, SUM(T1.[LineTotal] / 100 * T0.[DpmPrcnt]) as Anticip, " +
                    "T2.[AcqstnRvrs] as Adquisi, SUM(T6.LineTotal) AS BasePortes, SUM(T6.VatSum) AS IvaPortes " +
                    "FROM ODPI AS T0 INNER JOIN " +
                    "DPI1 AS T1 ON T1.DocEntry = T0.DocEntry INNER JOIN " +
                    "OVTG AS T2 ON T2.Code = T1.VatGroup INNER JOIN " +
                    "OCRD AS T3 ON T3.CardCode = T0.CardCode LEFT OUTER JOIN " +
                    "OCST AS T5 ON T3.Country = T5.Country AND T3.State1 = T5.Code LEFT OUTER JOIN " +
                    "DPI3 AS T6 ON T0.DocEntry = T6.DocEntry LEFT OUTER JOIN " +
                    "OCRY AS T4 ON T3.Country = T4.Code " +
                    "WHERE (T0.DOCDATE >= '" + DesdeFecha.Substring(6, 4) + "/" + DesdeFecha.Substring(3, 2) + "/" + DesdeFecha.Substring(0, 2) + "' " +
                    "AND T0.DOCDATE    <= '" + HastaFecha.Substring(6, 4) + "/" + HastaFecha.Substring(3, 2) + "/" + HastaFecha.Substring(0, 2) + "' " +
                    "AND T3.CARDCODE   >= '" + DesdeIC + "' " +
                    "AND T3.CARDCODE   <= '" + HastaIC + "') " +
                    "AND T2.IsEC='Y' ";//ES de la UE y sus facturas tb
            if (ValorAcreedor && ValorDeudor)
            {
                StrSql = StrSql + "AND (T3.CardType = 'C' OR T3.CardType = 'S') ";
            }
            else
            {
                if (ValorDeudor)
                {
                    StrSql = StrSql + "AND T3.CardType = 'C' ";
                }
                if (ValorAcreedor)
                {
                    StrSql = StrSql + "AND T3.CardType = 'S' ";
                }
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkICRet").Specific;
            if (oCheckBox.Checked == false && StrRetencion == "Y") //sólo incluimos a los que no tienen retencion, si el sistema asi lo tiene
            {
                StrSql = StrSql + "AND (T3.WTLiable='N' or T3.WTLiable='' or T3.WTLiable IS NULL) ";
            }
            StrSql = StrSql + "GROUP BY T3.CardCode, T0.[Indicator], T2.AcqstnRvrs";

            daSql = new System.Data.SqlClient.SqlDataAdapter(StrSql, csVariablesGlobales.conAddon);
            dtSql = new System.Data.DataTable();
            daSql.Fill(dtSql);
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            if (dtSql.Rows.Count > 0)
            {
                for (i = 0; i < dtSql.Rows.Count; i++)
                {
                    oStaticText.Caption = "Cargando Anticipo " + Convert.ToUInt32(i + 1) + " de " + dtSql.Rows.Count;
                    System.Windows.Forms.Application.DoEvents();
                    drSql = dtSql.Rows[i];
                    //Label1.Text = "ALBARAN " & dr("DocNum").ToString
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = csUtilidades.DameValor("[@SIA_MOD349]", "Max(Code)", "");
                    if (StrCodMax == "")
                    {
                        StrCodMax = "0";
                    }

                    //dblSaldo = Convert.ToDouble(drSql["Total"].ToString());
                    //StrCodMax = Convert.ToString(Convert.ToInt64(StrCodMax) + 1);
                    //StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    //oUserTable.Code = StrCodMax;
                    //oUserTable.Name = StrCodMax;
                    if (drSql["IvaPortes"].ToString() == null || drSql["IvaPortes"].ToString() == "")
                    {
                        IvaPortes = 0;
                    }
                    else
                    {
                        IvaPortes = Convert.ToDouble(drSql["IvaPortes"].ToString());
                    }
                    if (drSql["BasePortes"].ToString() == null || drSql["BasePortes"].ToString() == "")
                    {
                        BasePortes = 0;
                    }
                    else
                    {
                        BasePortes = Convert.ToDouble(drSql["BasePortes"].ToString());
                    }
                    dblSaldo = Convert.ToDouble(drSql["Base"].ToString()) +
                            Convert.ToDouble(drSql["Iva"].ToString()) +
                            IvaPortes + BasePortes;

                    StrCodMax = Convert.ToString(Convert.ToInt64(StrCodMax) + 1);
                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    oUserTable.Code = StrCodMax;
                    oUserTable.Name = StrCodMax;
                    oUserTable.UserFields.Fields.Item("U_CodIC").Value = Convert.ToString(drSql["CodIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_NomIC").Value = Convert.ToString(drSql["NomIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_TipIC").Value = Convert.ToString(drSql["TipIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_NifIC").Value = Convert.ToString(drSql["NifIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(drSql["Base"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Prov").Value = Convert.ToString(drSql["Prov"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Pais").Value = Convert.ToString(drSql["Pais"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Indic").Value = Convert.ToString(drSql["Indic"].ToString());
                    oUserTable.UserFields.Fields.Item("U_CP").Value = Convert.ToString(drSql["CP"].ToString());
                    oUserTable.UserFields.Fields.Item("U_CodInf").Value = Convert.ToString(drSql["CodInf"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Iva").Value = Convert.ToDouble(drSql["Iva"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Anticip").Value = Convert.ToDouble(drSql["Anticip"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Adquisi").Value = Convert.ToString(drSql["Adquisi"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Total").Value = Convert.ToDouble(dblSaldo);
                    oUserTable.UserFields.Fields.Item("U_TipDoc").Value = Convert.ToString("ANTICIPO").ToString();

                    if (oUserTable.Add() != 0)
                    {
                        MessageBox.Show(csVariablesGlobales.oCompany.GetLastErrorDescription());
                        return;
                    }
                }
            }
            #endregion
            #region CARGO DIARIO
            StrSql = "SELECT T4.[CardCode] as CodIC, MIN(T4.[CardName]) as NomIC, " +
            "T4.CardType as TipIC, MIN(T4.[LicTradNum]) as NifIC, " +
            "SUM(T1.Debit) AS Debe, SUM(T1.Credit) AS Haber, " +
            "SUM(T1.[BaseSum]) as Base, MIN(T6.[Name]) as Prov, " +
            "MIN(T5.[Name]) as Pais, T0.[Indicator] as Indic, " +
            "MIN(T4.[ZipCode]) as CP, " +
            "MIN(T5.[ReportCode]) as CodInf " +
            "FROM OJDT AS T0 INNER JOIN " +
            "JDT1 AS T1 ON T0.TransId = T1.TransId INNER JOIN " +
            "OCRD AS T4 ON T1.ShortName = T4.CardCode LEFT OUTER JOIN " +
            "OCST AS T6 ON T4.Country = T6.Country AND T6.Code = T4.State1 LEFT OUTER JOIN " +
            "OCRY AS T5 ON T5.Code = T4.Country " +
            "WHERE  T0.ReportEU = 'Y'  AND  T1.TransType='30' " +
            "AND (T1.RefDate >= '" + DesdeFecha.Substring(6, 4) + "/" + DesdeFecha.Substring(3, 2) + "/" + DesdeFecha.Substring(0, 2) + "' " +
            "AND T1.REFDATE <= '" + HastaFecha.Substring(6, 4) + "/" + HastaFecha.Substring(3, 2) + "/" + HastaFecha.Substring(0, 2) + "' " +
            "AND T4.CARDCODE >='" + DesdeIC + "' " +
            "AND T4.CARDCODE <='" + HastaIC + "') " +
            "AND T4.VatStatus='E' ";
            if (ValorAcreedor && ValorDeudor)
            {
                StrSql = StrSql + "AND (T4.CardType = 'C' OR T4.CardType = 'S') ";
            }
            else
            {
                if (ValorDeudor)
                {
                    StrSql = StrSql + "AND T4.CardType = 'C' ";
                }
                if (ValorAcreedor)
                {
                    StrSql = StrSql + "AND T4.CardType = 'S' ";
                }
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkICRet").Specific;
            if (oCheckBox.Checked == false && StrRetencion == "Y") //sólo incluimos a los que no tienen retencion, si el sistema asi lo tiene
            {
                StrSql = StrSql + "AND (T4.WTLiable='N' or T4.WTLiable='' or T4.WTLiable IS NULL) ";
            }
            StrSql = StrSql + "GROUP BY T4.CardCode, T4.CardType, T0.Indicator";

            daSql = new System.Data.SqlClient.SqlDataAdapter(StrSql, csVariablesGlobales.conAddon);
            dtSql = new System.Data.DataTable();
            daSql.Fill(dtSql);
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            if (dtSql.Rows.Count > 0)
            {
                for (i = 0; i < dtSql.Rows.Count; i++)
                {
                    oStaticText.Caption = "Cargando Anticipo " + Convert.ToUInt32(i + 1) + " de " + dtSql.Rows.Count;
                    System.Windows.Forms.Application.DoEvents();
                    drSql = dtSql.Rows[i];
                    //Label1.Text = "ALBARAN " & dr("DocNum").ToString
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = csUtilidades.DameValor("[@SIA_MOD349]", "Max(Code)", "");
                    if (StrCodMax == "")
                    {
                        StrCodMax = "0";
                    }
                    StrCodMax = Convert.ToString(Convert.ToInt64(StrCodMax) + 1);
                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    oUserTable.Code = StrCodMax;
                    oUserTable.Name = StrCodMax;
                    oUserTable.UserFields.Fields.Item("U_CocIC").Value = Convert.ToString(drSql["CodIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_NomIC").Value = Convert.ToString(drSql["NomIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_TipIC").Value = Convert.ToString(drSql["TipIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_NifIC").Value = Convert.ToString(drSql["NifIC"].ToString());
                    if ((drSql["Base"].ToString()) == "")
                    {
                        oUserTable.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(0);
                    }
                    else
                    {
                        oUserTable.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(drSql["Base"].ToString());
                    }

                    oUserTable.UserFields.Fields.Item("U_Prov").Value = Convert.ToString(drSql["Prov"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Pais").Value = Convert.ToString(drSql["Pais"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Indic").Value = Convert.ToString(drSql["Indic"].ToString());
                    oUserTable.UserFields.Fields.Item("U_CP").Value = Convert.ToString(drSql["CP"].ToString());
                    oUserTable.UserFields.Fields.Item("U_CodInf").Value = Convert.ToString(drSql["CodInf"].ToString());

                    if (Convert.ToString(drSql["TipIC"].ToString()) == "C")
                    {
                        //SUMO LO DEL DEBE
                        oUserTable.UserFields.Fields.Item("U_Iva").Value = 0; //CDbl(dr([ª0001ª]).ToString)
                        if (drSql["Base"].ToString() == "")
                        {
                            dblSaldo = 0;
                        }
                        else
                        {
                            if (Convert.ToDouble(drSql["Debe"].ToString()) != 0)
                            {
                                dblSaldo = Convert.ToDouble(drSql["Debe"].ToString());
                            }
                            else
                            {
                                dblSaldo = (-1) * Convert.ToDouble(drSql["Haber"].ToString());
                            }
                        }
                    }
                    else
                    {
                        //SUMO LO DEL HABER
                        oUserTable.UserFields.Fields.Item("U_Iva").Value = 0; //CDbl(dr([ª0001ª]).ToString)
                        if (drSql["Base"].ToString() == "")
                        {
                            dblSaldo = 0;
                        }
                        else
                        {
                            if (Convert.ToDouble(drSql["Haber"].ToString()) != 0)
                            {
                                dblSaldo = Convert.ToDouble(drSql["Haber"].ToString());
                            }
                            else
                            {
                                dblSaldo = (-1) * Convert.ToDouble(drSql["Debe"].ToString());
                            }

                        }

                    }
                    oUserTable.UserFields.Fields.Item("U_Anticip").Value = Convert.ToDouble(0);
                    oUserTable.UserFields.Fields.Item("U_Total").Value = Convert.ToDouble(dblSaldo);
                    oUserTable.UserFields.Fields.Item("U_TipDoc").Value = Convert.ToString("DIARIO").ToString();

                    if (oUserTable.Add() != 0)
                    {
                        MessageBox.Show(csVariablesGlobales.oCompany.GetLastErrorDescription());
                        return;
                    }
                }
            }
            #endregion
            //PROVEEDOR
            #region CARGO FACTURA
            StrSql = " SELECT     T3.CardCode AS Codigo, MIN(T3.CardName) AS NombreIC, " +
                    "MIN(T3.CardType) AS TipoIC, MIN(T3.LicTradNum) AS Nif, " +
                    "MIN(T5.Name) AS Provincia, " +
                    "MIN(T4.Name) AS Pais, T0.[Indicator] AS Indicador,  " +
                    "MIN(T3.ZipCode) AS CodigoPostal, MIN(T4.ReportCode) AS CodigoInforme, " +
                    "SUM(T0.DocTotal - T0.VatSum) AS base, " +
                    "SUM(T0.VatSum - T0.EquVatSum) AS Iva, " +
                    "SUM(T0.DocTotal) AS Total " +
                    "FROM         OPCH AS T0 INNER JOIN " +
                    "V_SIA_TipoIvaFactComp T1 ON T1.DocEntry=T0.DocEntry INNER JOIN " +
                    "OCRD AS T3 ON T3.CardCode = T0.CardCode LEFT OUTER JOIN " +
                    "OCST AS T5 ON T3.Country = T5.Country AND T3.State1 = T5.Code LEFT OUTER JOIN " +
                    "OCRY AS T4 ON T3.Country = T4.Code " +
                    "WHERE (T0.DOCDATE >= '" + DesdeFecha.Substring(6, 4) + "/" + DesdeFecha.Substring(3, 2) + "/" + DesdeFecha.Substring(0, 2) + "' " +
                    "AND T0.DOCDATE    <= '" + HastaFecha.Substring(6, 4) + "/" + HastaFecha.Substring(3, 2) + "/" + HastaFecha.Substring(0, 2) + "' " +
                    "AND T3.CARDCODE   >= '" + DesdeIC + "' " +
                    "AND T3.CARDCODE   <= '" + HastaIC + "') " +
                    "AND T1.IsEC='Y' AND T1.Tipo='F'  AND T3.VatStatus='E' ";//ES de la UE y sus facturas tb
            if (ValorAcreedor && ValorDeudor)
            {
                StrSql = StrSql + "AND (T3.CardType = 'C' OR T3.CardType = 'S') ";
            }
            else
            {
                if (ValorDeudor)
                {
                    StrSql = StrSql + "AND T3.CardType = 'C' ";
                }
                if (ValorAcreedor)
                {
                    StrSql = StrSql + "AND T3.CardType = 'S' ";
                }
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkICRet").Specific;
            if (oCheckBox.Checked == false && StrRetencion == "Y") //sólo incluimos a los que no tienen retencion, si el sistema asi lo tiene
            {
                StrSql = StrSql + "AND (T3.WTLiable='N' or T3.WTLiable='' or T3.WTLiable IS NULL) ";
            }
            StrSql = StrSql + "GROUP BY T3.CardCode, T0.Indicator";

            daSql = new System.Data.SqlClient.SqlDataAdapter(StrSql, csVariablesGlobales.conAddon);
            dtSql = new System.Data.DataTable();
            daSql.Fill(dtSql);
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            if (dtSql.Rows.Count > 0)
            {
                for (i = 0; i < dtSql.Rows.Count; i++)
                {
                    oStaticText.Caption = "Cargando Anticipo " + Convert.ToUInt32(i + 1) + " de " + dtSql.Rows.Count;
                    System.Windows.Forms.Application.DoEvents();
                    drSql = dtSql.Rows[i];
                    //Label1.Text = "ALBARAN " & dr("DocNum").ToString
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = csUtilidades.DameValor("[@SIA_MOD349]", "Max(Code)", "");
                    if (StrCodMax == "")
                    {
                        StrCodMax = "0";
                    }
                    dblSaldo = Convert.ToDouble(drSql["Base"].ToString()) + Convert.ToDouble(drSql["Iva"].ToString());
                    StrCodMax = Convert.ToString(Convert.ToInt64(StrCodMax) + 1);
                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    oUserTable.Code = StrCodMax;
                    oUserTable.Name = StrCodMax;

                    oUserTable.UserFields.Fields.Item("U_CodIC").Value = Convert.ToString(drSql["CodIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_NomIC").Value = Convert.ToString(drSql["NomIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_TipIC").Value = Convert.ToString(drSql["TipIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_NifIC").Value = Convert.ToString(drSql["NifIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(drSql["Base"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Prov").Value = Convert.ToString(drSql["Prov"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Pais").Value = Convert.ToString(drSql["Pais"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Indic").Value = Convert.ToString(drSql["Indic"].ToString());
                    oUserTable.UserFields.Fields.Item("U_CP").Value = Convert.ToString(drSql["CP"].ToString());
                    oUserTable.UserFields.Fields.Item("U_CodInf").Value = Convert.ToString(drSql["CodInf"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Iva").Value = Convert.ToDouble(drSql["Iva"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Total").Value = Convert.ToDouble(dblSaldo);
                    oUserTable.UserFields.Fields.Item("U_TipDoc").Value = Convert.ToString("FACTURA").ToString();

                    if (oUserTable.Add() != 0)
                    {
                        MessageBox.Show(csVariablesGlobales.oCompany.GetLastErrorDescription());
                        return;
                    }
                }
            }
            #endregion
            #region CARGO ABONO
            StrSql = " SELECT T3.CardCode AS Codigo, MIN(T3.CardName) AS NombreIC, " +
                    "MIN(T3.CardType) AS TipoIC, MIN(T3.LicTradNum) AS Nif, MIN(T5.Name) AS Provincia, " +
                    "MIN(T4.Name) AS Pais, T0.[Indicator] AS Indicador,  " +
                    "MIN(T3.ZipCode) AS CodigoPostal, MIN(T4.ReportCode) AS CodigoInforme, " +
                    "SUM(T0.DocTotal - T0.VatSum) AS base, SUM(T0.VatSum - T0.EquVatSum) AS Iva,  " +
                    "SUM(T0.DocTotal) AS Total " +
                    "FROM  ORPC AS T0 INNER JOIN " +
                    "V_SIA_TipoIvaFactComp T1 ON T1.DocEntry=T0.DocEntry INNER JOIN " +
                    "OCRD AS T3 ON T3.CardCode = T0.CardCode LEFT OUTER JOIN " +
                    "OCST AS T5 ON T3.Country = T5.Country AND T3.State1 = T5.Code LEFT OUTER JOIN " +
                    "OCRY AS T4 ON T3.Country = T4.Code " +
                    "WHERE (T0.DOCDATE >= '" + DesdeFecha.Substring(6, 4) + "/" + DesdeFecha.Substring(3, 2) + "/" + DesdeFecha.Substring(0, 2) + "' " +
                    "AND T0.DOCDATE <= '" + HastaFecha.Substring(6, 4) + "/" + HastaFecha.Substring(3, 2) + "/" + HastaFecha.Substring(0, 2) + "' " +
                    "AND T3.CARDCODE >='" + DesdeIC + "' " +
                    "AND T3.CARDCODE <='" + HastaIC + "') " +
                    "AND T1.IsEC='Y' AND T1.Tipo='A'  AND T3.VatStatus='E' ";
            if (ValorAcreedor && ValorDeudor)
            {
                StrSql = StrSql + "AND (T3.CardType = 'C' OR T3.CardType = 'S') ";
            }
            else
            {
                if (ValorDeudor)
                {
                    StrSql = StrSql + "AND T3.CardType = 'C' ";
                }
                if (ValorAcreedor)
                {
                    StrSql = StrSql + "AND T3.CardType = 'S' ";
                }
            }
            oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkICRet").Specific;
            if (oCheckBox.Checked == false && StrRetencion == "Y") //sólo incluimos a los que no tienen retencion, si el sistema asi lo tiene
            {
                StrSql = StrSql + "AND (T3.WTLiable='N' or T3.WTLiable='' or T3.WTLiable IS NULL) ";
            }
            StrSql = StrSql + "GROUP BY T3.CardCode, T0.[Indicator]";
            daSql = new System.Data.SqlClient.SqlDataAdapter(StrSql, csVariablesGlobales.conAddon);
            dtSql = new System.Data.DataTable();
            daSql.Fill(dtSql);
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            if (dtSql.Rows.Count > 0)
            {
                for (i = 0; i < dtSql.Rows.Count; i++)
                {
                    oStaticText.Caption = "Cargando Abono " + Convert.ToUInt32(i + 1) + " de " + dtSql.Rows.Count;
                    System.Windows.Forms.Application.DoEvents();
                    drSql = dtSql.Rows[i];
                    //Label1.Text = "ALBARAN " & dr("DocNum").ToString
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = csUtilidades.DameValor("[@SIA_MOD349]", "Max(Code)", "");
                    if (StrCodMax == "")
                    {
                        StrCodMax = "0";
                    }
                    StrCodMax = Convert.ToString(Convert.ToInt64(StrCodMax) + 1);
                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                    oUserTable.Code = StrCodMax;
                    oUserTable.Name = StrCodMax;
                    dblSaldo = Convert.ToDouble(drSql["Total"].ToString());

                    oUserTable.UserFields.Fields.Item("U_CodIC").Value = Convert.ToString(drSql["CodIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_NomIC").Value = Convert.ToString(drSql["NomIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_TipIC").Value = Convert.ToString(drSql["TipIC"].ToString());
                    oUserTable.UserFields.Fields.Item("U_NifIC").Value = Convert.ToString(drSql["NifIC"].ToString());
                    if ((drSql["Base"].ToString()) != "")
                    {
                        oUserTable.UserFields.Fields.Item("U_Base").Value = Convert.ToDouble(drSql["Base"].ToString());
                    }
                    else
                    {
                        oUserTable.UserFields.Fields.Item("U_Base").Value = 0;
                    }
                    oUserTable.UserFields.Fields.Item("U_Prov").Value = Convert.ToString(drSql["Prov"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Pais").Value = Convert.ToString(drSql["Pais"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Indic").Value = Convert.ToString(drSql["Indic"].ToString());
                    oUserTable.UserFields.Fields.Item("U_CP").Value = Convert.ToString(drSql["CP"].ToString());
                    oUserTable.UserFields.Fields.Item("U_CodInf").Value = Convert.ToString(drSql["CodInf"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Iva").Value = Convert.ToDouble(drSql["Iva"].ToString());
                    oUserTable.UserFields.Fields.Item("U_Total").Value = Convert.ToDouble(-dblSaldo);
                    oUserTable.UserFields.Fields.Item("U_TipDoc").Value = Convert.ToString("ABONO").ToString();

                    if (oUserTable.Add() != 0)
                    {
                        MessageBox.Show(csVariablesGlobales.oCompany.GetLastErrorDescription());
                        return;
                    }
                }
            }
            #endregion

        }

        public void Imprimir349(string Menu, string Impreso)
        {
            SAPbouiCOM.EditText oEditText;
            //SBOUtilyArch.cLanzarReport a;
            string StrParm1 = "";
            string StrParm2 = "";
            string StrParm3 = "";
            string StrParm4 = "";
            string StrParm5 = "";
            string StrFormula = "";
            string StrBD;
            string StrUsuAdo;
            DateTime Fecha;
            //a = new SBOUtilyArch.cLanzarReport();
            StrBD = csVariablesGlobales.oCompany.ToString();
            StrUsuAdo = csVariablesGlobales.oCompany.DbUserName;
            oForm = csVariablesGlobales.SboApp.Forms.Item("FrmModelo349");
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
                    //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtDesFec").Specific;
                    StrParm3 = "{@FecD}=" + oEditText.String;
                    //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtHasFec").Specific;
                    StrParm4 = "{@FecH}=" + oEditText.String;
                    StrFormula = "";
                    break;
            }
            switch (Menu)
            {
                case "519": // Vista Previa
                    //a.Informe(ref StrReport, ref csVariablesGlobales.StrServidor, ref csVariablesGlobales.StrBaseDatos, ref csVariablesGlobales.StrUserConex, ref csVariablesGlobales.StrPassConex, StrFormula, StrParm1, StrParm2, StrParm3, StrParm4, StrParm5, csVariablesGlobales.StrSubRep1, csVariablesGlobales.StrSubRep2, csVariablesGlobales.StrSubRep3, csVariablesGlobales.StrSubRep4);
                    break;
                case "520": // Imprimir
                    //a.Imprimir(ref StrReport, ref csVariablesGlobales.StrServidor, ref csVariablesGlobales.StrBaseDatos, ref csVariablesGlobales.StrUserConex, ref csVariablesGlobales.StrPassConex, StrFormula, StrParm1, StrParm2, StrParm3, StrParm4, StrParm5, csVariablesGlobales.StrSubRep1, csVariablesGlobales.StrSubRep2, csVariablesGlobales.StrSubRep3, csVariablesGlobales.StrSubRep4);
                    break;
            }
        }

        public void GeneraFichero()
        {
            SqlDataAdapter daSql;
            System.Data.DataTable dtSql;
            int i;
            DataRow drSql;
            SAPbouiCOM.EditText oEditText;
            SAPbouiCOM.Form oForm;
            csUtilidades Utilidades = new csUtilidades();

            #region Declaraciones
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
            string StrLinea;
            string StrCodFactura = "FACTURA";
            string StrCodAbono = "ABONO";
            string StrCodDiario = "DIARIO";
            string StrCodAntic = "ANTICIPO";
            string StrAnnoRect = "";
            string StrPeriodoRect = "";
            string StrPais;
            string StrNomPais;
            string StrTipo;
            string StrNomRel;
            double dblBaseRec;
            long lngNumOpe;
            string ClaveOp;
            string StrSql;
            string StrFichero;
            string StrEjercicio;
            string DesdeFecha;
            string HastaFecha;
            string StrDestinoT = "";
            string StrPeriodo = "";
            #endregion

            oForm = csVariablesGlobales.SboApp.Forms.Item("FrmModelo349");
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtEjer").Specific;
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optAnual").Specific;
            if (oOptionBtn.Selected)
            {
                DesdeFecha = "01/01/" + oEditText.String;
                HastaFecha = "31/12/" + oEditText.String;
                StrPeriodo = "0A";
            }
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optPriTri").Specific;
            if (oOptionBtn.Selected)
            {
                DesdeFecha = "01/01/" + oEditText.String;
                HastaFecha = "31/03/" + oEditText.String;
                StrPeriodo = "1T";
            }
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optSegTri").Specific;
            if (oOptionBtn.Selected)
            {
                DesdeFecha = "01/04/" + oEditText.String;
                HastaFecha = "30/06/" + oEditText.String;
                StrPeriodo = "2T";
            }
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optTerTri").Specific;
            if (oOptionBtn.Selected)
            {
                DesdeFecha = "01/07/" + oEditText.String;
                HastaFecha = "30/09/" + oEditText.String;
                StrPeriodo = "3T";
            }
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optCuaTri").Specific;
            if (oOptionBtn.Selected)
            {
                DesdeFecha = "01/10/" + oEditText.String;
                HastaFecha = "31/12/" + oEditText.String;
                StrPeriodo = "4T";
            }
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optDis").Specific;
            if (oOptionBtn.Selected)
            {
                StrDestinoT = "D";
            }
            oOptionBtn = (SAPbouiCOM.OptionBtn)oForm.Items.Item("optTel").Specific;
            if (oOptionBtn.Selected)
            {
                StrDestinoT = "T";
            }
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
                StrNifEmp = csUtilidades.DameValor("OADM", "TaxIdNum", "").ToString().ToUpper();
                StrNifEmp = StrNifEmp.Substring(2, StrNifEmp.Length - 2);
                StrNomEmp = csUtilidades.DameValor("OADM", "CompnyName", "").ToString().ToUpper();
                StrTelfEmp = csUtilidades.DameValor("OADM", "Phone1", "").ToString().ToUpper();
                StrNomRel = csUtilidades.DameValor("OADM", "manager", "").ToString().ToUpper();
                //StrFICHERO = strDestino;
                //FileClose(1);
                StrSql = "SELECT U_NifIC, U_TipIC, U_NomIC, SUM(U_Total) AS Total " +
                        "FROM  [@SIA_MOD349] GROUP BY U_TipIC, U_NifIC, U_NomIC " +
                        "HAVING (SUM(U_Total) <> 0) ORDER BY U_TipIC, U_NifIC, U_NomIC";
                daSql = new System.Data.SqlClient.SqlDataAdapter(StrSql, csVariablesGlobales.conAddon);
                dtSql = new System.Data.DataTable();
                daSql.Fill(dtSql);
                //CargarRst(DTSQL, StrSql);

                if (dtSql.Rows.Count > 0)
                {
                    //aqui ya tengo el nº de registros
                    lngNumOpe = dtSql.Rows.Count; //nº de operadores
                    double dblTotal;
                    double dblRect; //total a rectificar
                    long lngRec; //nº rectificadores
                    dblTotal = Convert.ToDouble(dtSql.Compute("SUM (Total)", ""));
                    dblRect = Convert.ToDouble(csUtilidades.ComPunN(dtSql.Compute("SUM (Total)", "Total<0").ToString()));
                    lngRec = Convert.ToInt64(csUtilidades.ComPunN(dtSql.Compute("count (U_NifIC)", "Total<0 ").ToString()));
                    //FileOpen(1, StrFICHERO, OpenMode.Output);
                    dblTotal = dblTotal * 100;
                    dblRect = dblRect * 100 * -1;
                    //Linea 1
                    //StrLinea = Space(250);
                    StrLinea = "1" +// Mid(StrLinea, 1, 1) = "1";
                            "349" +
                            StrEjercicio +
                            StrNifEmp + // /*NIF EMPRESA*/
                            StrNomEmp.PadRight(40, ' ') + // /*NOMBRE EMPRESA*/
                            StrDestinoT + ///*TIPO DE FICHERO*/
                            StrTelfEmp.PadRight(9, '0') +// /*TELEFONO EMPRESA*/
                            StrNomRel.PadRight(40, ' ') + // /*NOMBRE PERSONA RELACIONARSE*/
                            "3430000000000" + // /*Nº JUSTIFICANTE DE LA DECLARACION*/
                            "  " +
                            "0000000000000" + ///*Nº JUSTIFICANTE DECLARACION ANTERIOR */
                            StrPeriodo + ///* PERIODO */'
                            lngNumOpe.ToString().PadRight(9, '0') +// /*Nº OPERADORES*/
                            dblTotal.ToString().PadRight(15, '0') +// /*Importe Operaciones*/
                            lngRec.ToString().PadRight(9, '0') + // /*Nº OPERACIONES RECTIFICADAS*/
                            dblRect.ToString().PadRight(15, '0') + // /*Importe Operaciones*/
                            " ".PadRight(64, ' ');
                    //PrintLine(1, StrLinea);

                    oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
                    for (i = 0; i < dtSql.Rows.Count; i++)
                    {
                        oStaticText.Caption = "Cargando Datos " + Convert.ToUInt32(i + 1) + " de " + dtSql.Rows.Count;
                        dblBaseRec = 0;
                        drSql = dtSql.Rows[i];
                        //LblProc.Text = "Cargando datos Interlocutor " + dr("U_NifIC").ToString() + "  " + i + 1 + " de " + DTSQL.Rows.Count;
                        System.Windows.Forms.Application.DoEvents();
                        StrCodCli = drSql["U_NifIC"].ToString();
                        StrTipo = drSql["U_TipIC"].ToString();
                        string Nombre = drSql["U_NomIC"].ToString();
                        if (StrTipo == "C")
                        {
                            ClaveOp = "E";
                        }
                        else
                        {
                            ClaveOp = "A";
                        }
                        if (StrCodCli != "")
                        {
                            DblSaldoC = 0;
                            //FACTURA
                            StrAux = csUtilidades.DameValor("[@SIA_MOD349]", "SUM(U_Saldo)", "U_NIF ='" + StrCodCli + "' AND U_TipDoc = '" + StrCodFactura + "' AND U_TipIC='" + StrTipo + "' GROUP BY U_NifIC, U_TipDoc");
                            if (StrAux == "")
                            {
                                DblFactura = 0;
                            }
                            else
                            {
                                DblFactura = Convert.ToDouble(csUtilidades.PunCom(StrAux));
                            }
                            //ABONO
                            StrAux = csUtilidades.DameValor("[@SIA_MOD349]", "SUM(U_Saldo)", "U_NIF ='" + StrCodCli + "' AND U_TipDoc = '" + StrCodAbono + "'  AND U_TipIC='" + StrTipo + "' GROUP BY U_NifIC, U_TipDoc");
                            if (StrAux == "")
                            {
                                DblAbono = 0;
                            }
                            else
                            {
                                DblAbono = Convert.ToDouble(csUtilidades.PunCom(StrAux));
                            }
                            //DIARIO
                            StrAux = csUtilidades.DameValor("[@SIA_MOD349]", "SUM(U_Saldo)", "U_NIF ='" + StrCodCli + "' AND U_TipDoc = '" + StrCodDiario + "'  AND U_TipIC='" + StrTipo + "' GROUP BY U_NifIC, U_TipDoc");
                            if (StrAux == "")
                            {
                                DblDiario = 0;
                            }
                            else
                            {
                                DblDiario = Convert.ToDouble(csUtilidades.PunCom(StrAux));
                            }
                            //ANTICIPO
                            StrAux = csUtilidades.DameValor("[@SIA_MOD349]", "SUM(U_Saldo)", "U_NIF ='" + StrCodCli + "' AND U_TipDoc = '" + StrCodAntic + "'  AND U_TipIC='" + StrTipo + "' GROUP BY U_NifIC, U_TipDoc");
                            if (StrAux == "")
                            {
                                DblAntic = 0;
                            }
                            else
                            {
                                DblAntic = Convert.ToDouble(csUtilidades.PunCom(StrAux));
                            }
                            DblSaldoC = DblFactura + DblAntic + DblDiario + DblAbono;
                            //con esto rellenamos los datos de la tabla H349
                            if (DblSaldoC < 0)
                            {
                                ////buscamos para periodos anteriores
                                //StrSqlAux = "SELECT U_NifIC, U_Total, U_Anno, U_Periodo FROM  [@DOR_H349] " +
                                //            "WHERE (U_Saldo >= '" + Utilidades.ComPun(Math.Abs(DblSaldoC)) + "') " +
                                //            "AND (U_Nif = N'" + StrCodCli + "')  AND U_TipoIC='" + StrTipo + "' " +
                                //            "ORDER BY U_Anno DESC, U_Periodo DESC";
                                ////CargarRst(rstAux, StrSql);
                                //daSqlAux = new System.Data.SqlClient.SqlDataAdapter(StrSql, csVariablesGlobales.conAddon);
                                //dtSqlAux = new System.Data.DataTable();
                                //daSqlAux.Fill(dtSqlAux);
                                //if (dtSqlAux.Rows.Count > 0)
                                //{
                                //    StrAnnoRect = dtSqlAux.Rows[0]["U_Anno"].ToString();
                                //    StrBaseRect = dtSqlAux.Rows[0]["U_Saldo"].ToString();
                                //    dblBaseRec = Convert.ToDouble(Utilidades.ComPunN(StrBaseRect));
                                //    StrPeriodoRect = dtSqlAux.Rows[0]["U_Periodo"].ToString();
                                //}
                                //else
                                //{
                                //    //sacamos un formulario pidiendo los datos
                                //    //FrmDatosRectificacion FDR;
                                //    //FDR = new FrmDatosRectificacion(StrEjercicio, StrPeriodo, DblSaldoC, StrCodCli, StrTipo);
                                //    //FDR.ShowDialog();
                                //    //StrAnnoRect = FDR.StrAnno;
                                //    //dblBaseRec = FDR.DblBase;
                                //    //StrBaseRect = dblBaseRec;
                                //    //StrPeriodoRect = FDR.StrPeriodo;
                                //}
                            }
                            else
                            {
                                //StrAnnoRect = "";
                                //StrBaseRect = "";
                                //StrPeriodoRect = "";
                            }
                            //saldo<0
                            //ahora insertamos en el historico

                            StrNifCli = StrCodCli;
                            StrNifCli = StrNifCli.Substring(2, StrNifCli.Length - 2);// 3, Len(StrNifCli)));
                            StrNomCli = csUtilidades.DameValor("[@SIA_MOD349]", "U_NombreIC", "U_NIF ='" + StrCodCli + "' AND U_TIPOIC='" + StrTipo + "'");
                            StrNomCli = StrNomCli.PadRight(40, ' ').ToUpper();
                            StrCpCli = csUtilidades.DameValor("OCRD", "ZipCode", "LicTradNum ='" + StrCodCli + "' AND CardType='" + StrTipo + "'");
                            StrCpCli = StrCpCli.PadRight(2, '0').ToUpper();
                            StrPais = csUtilidades.DameValor("OCRD", "Country", "LicTradNum ='" + StrCodCli + "' AND CardType='" + StrTipo + "'");
                            StrNomPais = csUtilidades.DameValor("OCRY", "Name", "Code='" + StrPais + "'");
                            //borramos para ese perido y ano lo que hay de ese interlocutor
                            //StrSql = "DELETE FROM [@DOR_H349] WHERE U_TipIC='" + drSql["U_TipIC"].ToString() + "'";
                            //StrSql = StrSql + " AND U_NifIC='" + drSql["U_NifIC"].ToString() + "'";
                            //StrSql = StrSql + " AND U_Anno='" + StrEjercicio + "'";
                            //StrSql = StrSql + " AND U_Periodo='" + StrPeriodo + "'";
                            //cmdSql = new SqlCommand(StrSql, csVariablesGlobales.conAddon);
                            //cmdSql.ExecuteNonQuery();

                            //StrCodMax = Utilidades.DameValor("[@DOR_TEMP349]", "Max(Code)", "");
                            //if (StrCodMax == "")
                            //{
                            //    StrCodMax = "0";
                            //}



                            //StrCodMax = Convert.ToString(Convert.ToInt64(StrCodMax) + 1);
                            //StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                            //Code = StrCodMax;

                            //StrSql = "INSERT INTO [@DOR_H349] " +
                            //        "([Code],[Name],[U_Anno],[U_Periodo],[U_NombreIC],[U_TipoIC],[U_Nif], " +
                            //        "[U_Pais] ,[U_Base] ,[U_AnoRect] ,[U_PerRect] , " +
                            //        "[U_BaseRect] ,[U_NomPais],[U_Saldo]) " +
                            //        "VALUES('" + Code + "','" + Code + "'," +
                            //        "'" + StrEjercicio + "'," +
                            //        "'" + StrPeriodo + "'," +
                            //        "'" + StrNomCli + "'," +
                            //        "'" + drSql["U_TipoIC"].ToString() + "'," +
                            //        "'" + drSql["U_NIF"].ToString() + "'," +
                            //        "'" + StrPais + "'," +
                            //        "'" + Utilidades.ComPun(DblSaldoC) + "'," +
                            //        "'" + StrAnnoRect + "'," +
                            //        "'" + StrPeriodoRect + "'," +
                            //        "'" + Utilidades.ComPunN(StrBaseRect) + "'," +
                            //        "'" + StrNomPais + "'," +
                            //        "'" + Utilidades.ComPun(DblSaldoC) + "')";
                            //cmdSql = new SqlCommand(StrSql, csVariablesGlobales.conAddon);
                            //cmdSql.ExecuteNonQuery();
                            //if (!bolSql)
                            //{
                            //    return;
                            //}
                            DblSaldoC = Convert.ToDouble(DblSaldoC); // System.String.Format(Convert.ToDouble(DblSaldoC), "0.00");
                            DblSaldoC = DblSaldoC * 100;
                            //Linea de Tipo 2
                            //StrLinea = Space(250);
                            StrLinea = "2" + // Mid(StrLinea, 1, 1) = "2";
                                        "349" +
                                        StrEjercicio.ToUpper() +
                                        StrNifEmp + // /*NIF EMPRESA*/
                                        " ".PadRight(57, ' ') + //Mid(StrLinea, 18, 57) = Space(57); ///* ESPACIOS EN BLANCO */
                                        StrPais.PadRight(2, ' ') +//Mid(StrLinea, 76, 2) = StrPais; ///*CODIGO PAIS DEL DECLARANTE */
                                        StrNifCli.PadRight(15, ' ') +//Mid(StrLinea, 78, 15) = StrNifCli; // /*NIF CLIENTE*/
                                        StrNomCli.PadRight(40, ' ') +//Mid(StrLinea, 93, 40) = StrNomCli; // /*NOMBRE CLIENTE*/
                                        ClaveOp;//Mid(StrLinea, 133, 1) = ClaveOp; // /* CLAVE OPERACION */
                            if (DblSaldoC >= 0)
                            {
                                StrLinea = StrLinea + DblSaldoC.ToString().PadRight(13, '0') +// System.String.Format(Convert.ToDouble(DblSaldoC), "0000000000000"); // /*BASE IMPONIBLE*/
                                            " ".PadRight(103, ' '); // Mid(StrLinea, 147, 103) = Space(103);
                            }
                            else
                            {
                                //es rectificativo
                                StrLinea = StrLinea + " ".PadRight(13, ' ') + //Mid(StrLinea, 134, 13) = Space(13);
                                            StrAnnoRect.PadRight(4, ' ') + //Mid(StrLinea, 147, 4) = StrAnnoRect;
                                            StrPeriodoRect.PadRight(2, ' ');//Mid(StrLinea, 151, 2) = StrPeriodoRect;
                                DblSaldoC = Math.Abs(DblSaldoC);
                                StrLinea = StrLinea + Convert.ToString(DblSaldoC.ToString().PadRight(13, '0'));
                                //Mid(StrLinea, 153, 13) = System.String.Format(Convert.ToDouble(DblSaldoC), "0000000000000"); // /*BASE IMPONIBLE*/
                                dblBaseRec = Convert.ToDouble(dblBaseRec.ToString("########0.00"));// System.String.Format(Convert.ToDouble(dblBaseRec), "0.00");
                                dblBaseRec = dblBaseRec * 100;
                                StrLinea = StrLinea + dblBaseRec.ToString().PadRight(13, '0'); //Mid(StrLinea, 166, 13) = System.String.Format(Convert.ToDouble(dblBaseRec), "0000000000000"); // /*BASE IMPONIBLE*/

                            }

                            //If DameValor("[@DOR_TEMP349]", "U_TipoIC", "U_NIF ='" & StrCodCli & "'") = "C" Then
                            //    Mid(StrLinea, 82, 1) = "B" 'CLIENTE B
                            //Else
                            //    Mid(StrLinea, 82, 1) = "A" 'PROVEEDOR A
                            //End If
                            //Mid(StrLinea, 83, 15) = Format(CDbl(DblSaldoC), "000000000000000")
                            //PrintLine(1, StrLinea);
                            SW.WriteLine(StrLinea);

                        }//nif<>[ª0000ª]
                        else
                        {
                            csVariablesGlobales.SboApp.MessageBox("El I.C. " + Nombre +
                                " no tiene informado el C.I.F./N.I.F. debe informarlo en su ficha y volver a generar el fichero."
                                , 1, "", "", "");
                        }
                    }
                    //FileClose(1);
                    SW.Close();
                    //MsgBox("DOR. Fichero creado con exito");
                }
                else
                {
                    //MsgBox("DOR. No hay datos para generar el fichero");
                }
                //FNGenerarHistorico349 = true;


                //LblProc.Text = "";
                //Application.DoEvents();
            }
            catch (Exception ex)
            {
                SW.Close();
                MessageBox.Show(ex.Message);
                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
            }
        }
    }
}
