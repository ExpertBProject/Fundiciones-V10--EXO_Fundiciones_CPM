using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Diagnostics;
using System.Threading;

namespace cLIENTE
{
    class csFrmImpresionListados
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
            CrearFormularioImpresionListados();
            oForm.Visible = true;
            //csUtilidades csUtilidades = new csUtilidades();
            csUtilidades.SaveAsXml(oForm, "FrmImpresionListados.xml", "");
            // events handled by SBO_Application_ItemEvent
            csVariablesGlobales.SboApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(FrmImpresionListados_ItemEvent);
            // events handled by SBO_Application_AppEvent
            csVariablesGlobales.SboApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(FrmImpresionListados_AppEvent);
        }

        private void CrearFormularioImpresionListados()
        {
            int BaseLeft = 0;
            int BaseTop = 0;
            
            #region Formulario
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "FrmImpresionListados";
            oCreationParams.FormType = "SIASL00002";

            oForm = csVariablesGlobales.SboApp.Forms.AddEx(oCreationParams);

            // set the form properties
            oForm.Title = "Impresión de Listados";
            oForm.Left = 300;
            oForm.ClientWidth = 500;
            oForm.Top = 100;
            oForm.ClientHeight = 400;
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
            oForm.DataSources.UserDataSources.Add("dsSelFich", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            #endregion

            #region Marcos
            #region Impresión de Listados
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
            oItem.Height = 300;

            //***************************
            // Adding a Static Text item
            //***************************
            oItem = oForm.Items.Add("lblList", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 50;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Listados";

            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #endregion

            #region Matrix
            BaseTop = 30;
            BaseLeft = 15;
            
            //***************************
            // Adding a Matrix item
            //***************************

            oItem = oForm.Items.Add("matLis", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oItem.Left = BaseLeft;
            oItem.Width = 465;
            oItem.Top = BaseTop;
            oItem.Height = 280;

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
            oColumn.Width = 200;
            oColumn.Editable = false;
            #endregion
            #region Columna Descripción Impreso
            // Add a column for Descripción Impreso
            oColumn = oColumns.Add("colDesImp", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Descripción Impresos";
            oColumn.Width = 250;
            oColumn.Editable = false;
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
            oItem.Top = 365;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";
            #endregion
            #endregion
        }

        public void FrmImpresionListados_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, 
                                                   out bool BubbleEvent)
        { 
            BubbleEvent = true;
            csImpresiones Impresiones = new csImpresiones();
            string Impreso;

            if (FormUID == "FrmImpresionListados")
            {
                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:

                        if (pVal.ItemUID == "matLis" && !pVal.BeforeAction && 
                            pVal.ActionSuccess && pVal.Row > 0)
                        {
                            oMatrix = (SAPbouiCOM.Matrix)(oForm.Items.Item(pVal.ItemUID).Specific);
                            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("colNomImp").Cells.Item(pVal.Row).Specific);
                            Impreso = csVariablesGlobales.StrRutRep + oEditText.String;
                            Impresiones.Informe(Impreso, "");
                            BubbleEvent = false;
                            break;
                        }
                        break;
                }
            }
        }

        private void FrmImpresionListados_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //System.Windows.Forms.Application.Exit();
                    break;            
            }
        }

        public void CargaMatrix(ref 
        {
            oForm = csVariablesGlobales.SboApp.Forms.Item("FrmImpresionListados");
            oMatrix = (SAPbouiCOM.Matrix)(oForm.Items.Item("matLis").Specific);
            oColumns = oMatrix.Columns;
            // Add DB data sources for the DB bound columns in the matrix
            oDBDataSource = oForm.DataSources.DBDataSources.Add("@" + csVariablesGlobales.Prefijo + "_REPORT");

            // getting the matrix column by the UID

            oColumn = oColumns.Item("colNomImp");
            // oColumn.DataBind.SetBound(True, "", "DSCardCode")
            oColumn.DataBind.SetBound(true, "@" + csVariablesGlobales.Prefijo + "_REPORT", "U_Report");

            oColumn = oColumns.Item("colDesImp");
            oColumn.DataBind.SetBound(true, "@" + csVariablesGlobales.Prefijo + "_REPORT", "U_Descrip");



            // Ready Matrix to populate data
            oMatrix.Clear();
            //oMatrix.AutoResizeColumns();

            // Querying the DB Data source
            oConditions = new SAPbouiCOM.Conditions();
            oCondition = oConditions.Add();
            oCondition.Alias = "U_TipDoc";
            oCondition.Operation = BoConditionOperation.co_EQUAL;
            oCondition.CondVal = "0";
            oDBDataSource.Query(oConditions);

            // setting the user data source data
            oMatrix.LoadFromDataSource(); 
        }
    }
}
