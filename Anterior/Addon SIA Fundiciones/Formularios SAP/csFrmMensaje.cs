using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Addon_SIA
{
    class csFrmMensaje
    {
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Item oItem = null;
        private SAPbouiCOM.Button oButton = null;
        private SAPbouiCOM.StaticText oStaticText = null;
        private SAPbouiCOM.EditText oEditText = null;
        private SAPbouiCOM.LinkedButton oLinkedButton = null;
        private SAPbouiCOM.ComboBox oComboBox = null;
        private SAPbouiCOM.Matrix oMatrix = null;


        public void CargarFormulario()
        {
            CrearFormularioMensaje();
            oForm.Visible = true;
            csUtilidades.SaveAsXml(oForm, "FrmMensaje.xml", "");
            csVariablesGlobales.SboApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(FrmMensaje_ItemEvent);
        }

        private void CrearFormularioMensaje()
        {
            int BaseLeft = 0;
            int BaseTop = 0;

            #region Formulario
            SAPbouiCOM.FormCreationParams oCreationParams = null;
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "FrmMensaje";
            oCreationParams.FormType = "SIASL00005";

            oForm = csVariablesGlobales.SboApp.Forms.AddEx(oCreationParams);

            // set the form properties
            oForm.Title = "";
            oForm.Left = 300;
            oForm.ClientWidth = 250;
            oForm.Top = 100;
            oForm.ClientHeight = 120;
            //oForm.EnableMenu("1293", true);
            #endregion

            #region Botones
            //*****************************************
            // Adding Items to the form
            // and setting their properties
            //*****************************************
            #region btnOk
            //************************
            // Adding a Cancel button
            //***********************

            oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5;
            oItem.Width = 65;
            oItem.Top = 85;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Ok";
            #endregion
            #region btnCancelar
            //************************
            // Adding a Cancel button
            //***********************

            oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 85;
            oItem.Width = 65;
            oItem.Top = 85;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";
            #endregion
            #endregion

            #region Campos
            #region Mensaje
            BaseLeft = 10;
            BaseTop = 20;

            oItem = oForm.Items.Add("lblMens", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 300;
            oItem.Top = BaseTop;
            oItem.Height = 14;
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Esta operación puede llevar algún tiempo.";
            oItem = oForm.Items.Add("lblMens2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 300;
            oItem.Top = BaseTop + 15;
            oItem.Height = 14;
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Espere a que se cierre la ventana.";
            oItem = oForm.Items.Add("lblMens3", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 300;
            oItem.Top = BaseTop + 30;
            oItem.Height = 14;
            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "¿Está seguro de realizar la acción?";
            BaseLeft = 0;
            BaseTop = 0;
            #endregion
            #endregion
        }

        public void FrmMensaje_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            //csUtilidades csUtilidades = new csUtilidades();
            BubbleEvent = true;

            if (pVal.FormUID == "FrmMensaje")
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == "1" && !pVal.BeforeAction && pVal.ActionSuccess)
                        {
                            switch (csVariablesGlobales.FormularioDesde)
                            {
                                case "3002":
                                    oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                                    oForm.Close();
                                    oForm = csVariablesGlobales.SboApp.Forms.Item(csVariablesGlobales.FormularioDesdeUID);
                                    SAPbobsCOM.Documents oDrafts;
                                    oDrafts = (SAPbobsCOM.Documents)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts));
                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                                    for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                                    {
                                        oDrafts.GetByKey(Convert.ToInt32(((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").
                                                                         Cells.Item(i).Specific).String));
                                        oDrafts.Remove();
                                    }
                                    csVariablesGlobales.FormularioDesde = "";
                                    csVariablesGlobales.FormularioDesdeUID = "";
                                    oForm.Close();
                                    break;
                            }
                        }
                        break;
                }
            }
        }
    }
}
