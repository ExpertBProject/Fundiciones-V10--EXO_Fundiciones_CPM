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
    class csFrmImpReports
    {
        private SAPbouiCOM.Form oForm;
        SAPbouiCOM.Item oItem;
        SAPbouiCOM.Button oButton;
        SAPbouiCOM.StaticText oStaticText;
        SAPbouiCOM.Matrix oMatrix;
        SAPbouiCOM.Columns oColumns;
        SAPbouiCOM.Column oColumn;
        SAPbouiCOM.FormCreationParams oCreationParams;
        SAPbouiCOM.EditText oEditText;
        SAPbouiCOM.DBDataSource oDBDataSource;
        SAPbouiCOM.Conditions oConditions;
        SAPbouiCOM.Condition oCondition;

        public void CargarFormulario()
        {
            CrearFormularioImpReports();
            oForm.Visible = true;
            //csUtilidades csUtilidades = new csUtilidades();
            csUtilidades.SaveAsXml(oForm, "FrmImpReports.xml", "");
            // events handled by SBO_Application_ItemEvent
            //csVariablesGlobales.SboApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(FrmImpresionListados_ItemEvent);
            // events handled by SBO_Application_AppEvent
            //csVariablesGlobales.SboApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(FrmImpresionListados_AppEvent);
        }

        private void CrearFormularioImpReports()
        {
            int BaseLeft = 0;
            int BaseTop = 0;

            #region Formulario
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "FrmImpReports";
            oCreationParams.FormType = "SIASL20000";

            oForm = csVariablesGlobales.SboApp.Forms.AddEx(oCreationParams);

            // set the form properties
            oForm.Title = "Seleccionar Report";
            oForm.Left = 300;
            oForm.ClientWidth = 490;
            oForm.Top = 100;
            oForm.ClientHeight = 260;
            #endregion

            #region Data Sources
            //  Adding a User Data Source
            oForm.DataSources.UserDataSources.Add("dsForm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsBorr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsDoc", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsMenu", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsNomImp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            oForm.DataSources.UserDataSources.Add("dsDesImp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            #endregion

            #region Campos
            #region Formulario
            //*************************
            // Adding a Text Edit item
            //*************************
            BaseLeft = 15;
            BaseTop = 10;

            oItem = oForm.Items.Add("txtForm", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 70;
            oItem.Width = 70;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.Enabled = false;
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsForm");

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblForm", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 70;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtForm";

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Formulario";
            #endregion
            #region Borrador
            //*************************
            // Adding a Text Edit item
            //*************************
            BaseLeft = 165;
            BaseTop = 10;

            oItem = oForm.Items.Add("txtBorr", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 70;
            oItem.Width = 70;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.Enabled = false;
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsBorr");

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblBorr", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 70;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtBorr";

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Borrador";
            #endregion
            #region Documento
            //*************************
            // Adding a Text Edit item
            //*************************
            BaseLeft = 15;
            BaseTop = 30;

            oItem = oForm.Items.Add("txtDoc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 70;
            oItem.Width = 70;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.Enabled = false;
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsDoc");

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblDoc", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 70;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtDoc";

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Documento";
            #endregion
            #region Menu
            //*************************
            // Adding a Text Edit item
            //*************************
            BaseLeft = 165;
            BaseTop = 30;

            oItem = oForm.Items.Add("txtMenu", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = BaseLeft + 70;
            oItem.Width = 70;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oItem.Enabled = false;
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "dsMenu");

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("lblMenu", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 70;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oItem.LinkTo = "txtMenu";

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Menu";
            #endregion
            #endregion
            
            #region Matrix
            BaseTop = 50;
            BaseLeft = 15;

            //***************************
            // Adding a Matrix item
            //***************************

            oItem = oForm.Items.Add("matRep", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oItem.Left = BaseLeft;
            oItem.Width = 470;
            oItem.Top = BaseTop;
            oItem.Height = 180;

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
            #region Columna Nombre Impreso
            // Add a column for Nombre Impreso
            oColumn = oColumns.Add("colNomImp", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Nombre Impreso";
            oColumn.Width = 185;
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "", "dsNomImp");
            #endregion
            #region Columna Descripción Impreso
            // Add a column for Descripción Impreso
            oColumn = oColumns.Add("colDesImp", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Descripción Impresos";
            oColumn.Width = 235;
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "", "dsDesImp");
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
            oItem.Left = 15;
            oItem.Width = 65;
            oItem.Top = 230;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";
            #endregion
            #endregion
        }
    }
}
