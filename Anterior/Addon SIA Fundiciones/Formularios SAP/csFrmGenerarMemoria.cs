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
    class csFrmGenerarMemoria
    {
        private SAPbouiCOM.Form oForm;

        public void CargarFormulario()
        {
            CrearFormularioGenerarMemoria();
            oForm.Visible = true;
            //csUtilidades Utilidades = new csUtilidades();
            csUtilidades.SaveAsXml(oForm, "FrmGenerarMemoria.xml", "");
            // events handled by SBO_Application_ItemEvent
            csVariablesGlobales.SboApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(FrmGenerarMemoria_ItemEvent);
            // events handled by SBO_Application_AppEvent
            csVariablesGlobales.SboApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(FrmGenerarMemoria_AppEvent);
        }

        private void CrearFormularioGenerarMemoria()
        {
            int BaseLeft = 0;
            int BaseTop = 0;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Button oButton = null;
            SAPbouiCOM.StaticText oStaticText = null;

            #region Formulario
            SAPbouiCOM.FormCreationParams oCreationParams = null;
            oCreationParams = ((SAPbouiCOM.FormCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "FrmGenerarMemoria";
            oCreationParams.FormType = "SIASL10003";

            oForm = csVariablesGlobales.SboApp.Forms.AddEx(oCreationParams);
            // set the form properties
            oForm.Title = "Generar Memoria";
            oForm.Left = 300;
            oForm.ClientWidth = 500;
            oForm.Top = 100;
            oForm.ClientHeight = 120;
            #endregion

            #region Marcos
            #region Generar Memoria
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
            oItem = oForm.Items.Add("lblGenMem", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft + 10;
            oItem.Width = 100;
            oItem.Top = BaseTop - 10;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Generar Memoria";

            BaseLeft = 0;
            BaseTop = 0;
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

            oItem = oForm.Items.Add("btnGenMem", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 80;
            oItem.Width = 105;
            oItem.Top = 85;
            oItem.Height = 22;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Generar Memoria";
            #endregion
            #endregion

            #region Labels
            BaseTop = 30;
            BaseLeft = 60;
            oItem = oForm.Items.Add("lblProces", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = BaseLeft;
            oItem.Width = 250;
            oItem.Top = BaseTop;
            oItem.Height = 14;

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "";
            #endregion
        }

        public void FrmGenerarMemoria_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (FormUID == "FrmGenerarMemoria")
            {
                //oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                        if (pVal.ItemUID == "btnGenMem" && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                        {
                            try
                            {
                                GeneraMemoria();
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

        private void FrmGenerarMemoria_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //System.Windows.Forms.Application.Exit();
                    break;
            }
        }

        private void GeneraMemoria()
        {
            Recordset oRecordSet;
            SAPbouiCOM.StaticText oStaticText;
            SqlDataAdapter daSQL = new SqlDataAdapter();
            System.Data.DataTable dtSQL = new System.Data.DataTable();
            System.Data.DataTable dtSQLEst = new System.Data.DataTable();
            DataRow drSQL;
            string strSql = "";
            string NombreQuery = "";
            string Categoria = "";
            string Identificador = "";
            string StrCodMax;
            SAPbobsCOM.UserTable pReTabla = null;

            //csUtilidades Utilidades = new csUtilidades();
            pReTabla = csVariablesGlobales.oCompany.UserTables.Item("SIA_DESMEM");

            //Limpiar tabla Destino Memoria
            strSql = "DELETE FROM [@SIA_DESMEM]";
            SqlCommand cmdBorrar = new SqlCommand(strSql, csVariablesGlobales.conAddon);
            cmdBorrar.ExecuteNonQuery();

            oRecordSet = (SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            csCrearConsultas CrearConsultas = new csCrearConsultas();
            oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("lblProces").Specific;
            oStaticText.Caption = "Generando Memoria";

            strSql = csUtilidades.DameValor("[@SIA_ESTCON]", "Count(Code)", "U_tipo = 'Memoria'");
            if (strSql != "")
            {
                strSql = "SELECT * FROM [@SIA_ESTCON] WHERE U_tipo = 'Memoria'";
                daSQL.SelectCommand = new SqlCommand(strSql, csVariablesGlobales.conAddon);
                daSQL.Fill(dtSQLEst);
                if (dtSQLEst.Rows.Count > 0)
                {
                    for (int i = 0; i < dtSQLEst.Rows.Count; i++)
                    {
                        drSQL = dtSQLEst.Rows[i];
                        NombreQuery = drSQL["U_nom_sql"].ToString();
                        Categoria = drSQL["U_cat_sql"].ToString();
                        Identificador = drSQL["U_ident"].ToString();
                        strSql = csUtilidades.DameValor("OUQR", "QString", "QName = '" + NombreQuery + "' AND QCategory = " + CrearConsultas.ExisteCategoria(Categoria));

                        if (strSql != "")
                        {
                            daSQL.SelectCommand = new SqlCommand(strSql, csVariablesGlobales.conAddon);
                            daSQL.Fill(dtSQL);
                            if (dtSQL.Rows.Count > 0)
                            {
                                for (int j = 0; j < dtSQL.Rows.Count; j++)
                                {
                                    if (oStaticText.Caption.Length >= 70)
                                    {
                                        oStaticText.Caption = "Generando Memoria";
                                    }
                                    else
                                    {
                                        oStaticText.Caption = oStaticText.Caption + ".";
                                    }
                                    System.Windows.Forms.Application.DoEvents();
                                    drSQL = dtSQL.Rows[j];
                                    System.Windows.Forms.Application.DoEvents();
                                    StrCodMax = csUtilidades.DameValor("[@SIA_DESMEM]", "Max(Code)", "");
                                    if (StrCodMax == "")
                                    {
                                        StrCodMax = "0";
                                    }
                                    StrCodMax = Convert.ToString(Convert.ToInt32(StrCodMax) + 1);
                                    StrCodMax = Convert.ToInt64(StrCodMax).ToString("00000000");
                                    pReTabla.Code = StrCodMax;
                                    pReTabla.Name = StrCodMax;
                                    pReTabla.UserFields.Fields.Item("U_ident").Value = Identificador;
                                    //double a = Convert.ToDouble(drSQL["Valor"].ToString());
                                    pReTabla.UserFields.Fields.Item("U_valor").Value = drSQL["Valor"].ToString();
                                }
                                if (pReTabla.Add() != 0)
                                {
                                    csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                                    return;
                                }
                            }
                            else
                            {
                                oStaticText.Caption = "No hay facturas de cargo que tratar.";
                            }
                        }
                    }
                }
                oStaticText.Caption = "Proceso Terminado";
            }
            else
            {
                oStaticText.Caption = "No ha estructura para generar la memoria";
            }
        }
    }
}
