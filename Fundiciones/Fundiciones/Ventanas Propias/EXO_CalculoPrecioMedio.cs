
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Cliente
{
    public class EXO_CalculoPrecioMedio
    {
        //static bool lLanzadoYo = false;

        public EXO_CalculoPrecioMedio()
        { }

        public EXO_CalculoPrecioMedio(bool lCreacion)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.FormCreationParams oParametrosCreacion = (SAPbouiCOM.FormCreationParams)(Matriz.oGlobal.conexionSAP.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams));

            try
            {
                Type Tipo = this.GetType();
                string strXML = Utilidades.LeoQueryFich("Formularios.EXO_CalculoPrecioMedio.srf", Tipo);
                oParametrosCreacion.XmlData = strXML;
                oParametrosCreacion.UniqueID = "";
                oParametrosCreacion.BorderStyle = BoFormBorderStyle.fbs_Fixed;

                oForm = Matriz.oGlobal.conexionSAP.SBOApp.Forms.AddEx(oParametrosCreacion);
                oForm.Visible = true;
            }
            catch (Exception ex)
            {
                Matriz.oGlobal.conexionSAP.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
            }            
        }

        public bool ItemEvent(EXO_Generales.EXO_infoItemEvent infoEvento)
        {
            SAPbouiCOM.Form oForm = Matriz.oGlobal.conexionSAP.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);

            switch (infoEvento.EventType)
            {

                // Ojo - Funciona sólo si lo lanza SAP, NO si lo cargo yo con la instanciacion
                // case BoEventTypes.et_FORM_VISIBLE:
                //    {
                //        if (!lLanzadoYo && !infoEvento.BeforeAction && infoEvento.ActionSuccess && infoEvento.InnerEvent)
                //        {
                //            if (oForm.Visible)
                //            {
                //                Inicializo(ref oForm);
                //            }
                //            return true;
                //        }
                //    }
                //    break;

                case BoEventTypes.et_MATRIX_LOAD:
                    #region Pinto los datos de la matriz
                    if (!infoEvento.BeforeAction && infoEvento.ActionSuccess && infoEvento.ItemUID == "matAmi")
                    {
                        //PintoMatrizAmigos(ref oForm);
                    }

                    if (!infoEvento.BeforeAction && infoEvento.ActionSuccess && infoEvento.ItemUID == "matEqui")
                    {
                        //PintoMatrizEquipos(ref oForm);
                    }

                    if (!infoEvento.BeforeAction && infoEvento.ActionSuccess && infoEvento.ItemUID == "matExclu")
                    {
                        //PintoMatrizExcluidos(ref oForm); 
                    }

                    if (!infoEvento.BeforeAction && infoEvento.ActionSuccess && infoEvento.ItemUID == "GrdTodos")
                    {
                        //PintoMatrizTodas(ref oForm);
                    }

                    #endregion

                    break;
                case BoEventTypes.et_ITEM_PRESSED:

                    #region Actualizar Grid

                    if (infoEvento.ItemUID == "bt_GLP" && !infoEvento.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.Grid oGrid = oForm.Items.Item("grdCli").Specific;
                            SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)Matriz.oGlobal.conexionSAP.compañia.GetBusinessObject(BoObjectTypes.BoRecordset);
                            string sValorItem = "";
                            string sValorPrice = "";
                            string gListaPrecios = "";

                            string sql = "SELECT U_EXO_INFV FROM [@EXO_OGEN1] WHERE U_EXO_NOMV = 'EXO_LISTAPRECIOS'";
                            oRec.DoQuery(sql);

                            if (oRec.RecordCount > 0)
                            {
                                gListaPrecios = oRec.Fields.Item("U_EXO_INFV").Value;
                            }
                            else
                            {
                                Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("La variable global EXO_LISAPRECIOS esta sin definir", 1, "Ok", "", "");
                            }

                            //string gListaPrecios = Matriz.oGlobal.conexionSAP.refCompañia.OGEN.valorVariable("EXO_LISTAPRECIOS");

                            for (int i = 0; i < oGrid.Rows.Count; i++)
                            {
                                try
                                {
                                    sValorItem = Convert.ToString(oGrid.DataTable.GetValue("Artículo", i));
                                    sValorPrice = Convert.ToString(oGrid.DataTable.GetValue("Precio Calculado Expert", i));
                                    if (sValorPrice != "0")
                                    { 
                                        string sqlUPD = "UPDATE ITM1 SET Price = " + sValorPrice.Replace(",", ".") + " WHERE PriceList = '" + gListaPrecios + "' AND ItemCode = '" + sValorItem + "'";
                                        oRec.DoQuery(sqlUPD);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("Error al actualizar el precio del artículo, " + sValorItem + " en la lista de precios, " + gListaPrecios + ". ERROR:" + ex.Message, 1, "Ok", "", "");
                                }
                            }
                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("La Lista de precios se actualizo correctamente.", 1, "Ok", "", "");
                        }
                        catch (Exception ex)
                        {
                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
                        }
                    }

                    if (infoEvento.ItemUID == "bt_VM" && !infoEvento.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.Grid oGrid = oForm.Items.Item("grdCli").Specific;
                            SAPbouiCOM.ComboBox oComboBox;

                            if (oGrid.Rows.SelectedRows.Count != 0)
                            {

                                //oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("CmbVal").Specific;
                                //string CmbVal = Convert.ToString(oComboBox.Selected.Value);
                                string CmbVal = "Ambas";

                                SAPbouiCOM.Grid oGrid2 = oForm.Items.Item("grdCli2").Specific;
                                oGrid2.DataTable = null;

                                int rowIndex = oGrid.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder);
                                string sValorItem = Convert.ToString(oGrid.DataTable.GetValue("Artículo", rowIndex));

                                string cDesdeFecha = oForm.DataSources.UserDataSources.Item("dsFD").ValueEx;
                                if (cDesdeFecha == "" || cDesdeFecha == null) cDesdeFecha = "20100101";

                                string cHastaFecha = oForm.DataSources.UserDataSources.Item("dsFH").ValueEx;
                                if (cHastaFecha == "" || cHastaFecha == null) cHastaFecha = "20900101";

                                string sql = "SELECT T1.ItemCode AS 'Artículo',T0.ItemName AS 'Nombre',T1.BASE_REF as 'Nº Documento',T1.DocLineNum  AS 'Linea',T1.DocDate as 'Fecha',T1.JrnlMemo AS 'Descripción',T1.InQty as 'Cantidad recibida',T1.OutQty  as 'Cantidad emitida',T1.Price as 'Precio'";
                                sql += " FROM OITM T0";
                                sql += " INNER JOIN OINM T1 ON T0.Itemcode = T1. Itemcode";
                                sql += " INNER JOIN OWHS T2 ON T2.WhsCode = T1.Warehouse";
                                sql += " WHERE T0.ItemCode = '" + sValorItem + "'";
                                sql += " AND T1.DocDate BETWEEN '" + cDesdeFecha + "' AND '" + cHastaFecha + "'";
                                if (CmbVal == "Ambas")
                                {
                                    sql += " AND T1.Price > 0 AND (T1.InQty > 0 or T1.OutQty > 0)";
                                }
                                if (CmbVal == "Entrada")
                                {
                                    sql += " AND T1.Price > 0 AND (T1.InQty > 0)";
                                }
                                if (CmbVal == "Salida")
                                {
                                    sql += " AND T1.Price > 0 AND (T1.OutQty > 0)";
                                }
                                sql += " AND(ISNULL(T2.U_EXO_ACPM, 'NO') = 'SI')";
                                sql += " ORDER BY T0.Itemcode, T1.DocDate";

                                SAPbobsCOM.Recordset oRecAmi2 = Matriz.oGlobal.SQL.sqlComoRsB1(sql);
                                SAPbouiCOM.DataTable oTabla2 = oForm.DataSources.DataTables.Item("TablaMov");

                                if (oRecAmi2.RecordCount > 0)
                                {
                                    oTabla2.ExecuteQuery(sql);
                                    oGrid2.DataTable = oTabla2;
                                }
                                else
                                {
                                    Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("No hay resgistros", 1, "Ok", "", "");
                                }
                            }
                            else
                            {
                                Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("No hay ningún registro seleccionado ", 1, "Ok", "", "");
                            }
                        }
                        catch (Exception ex)
                        {
                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
                        }
                    }


                    if (infoEvento.ItemUID == "bt_act" && infoEvento.BeforeAction)
                    {
                        try
                        {
                            string cGrupoArticulo = oForm.DataSources.UserDataSources.Item("dsGru2").Value;
                            SAPbouiCOM.ComboBox oComboBox;

                            //oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("CmbVal").Specific;
                            //string CmbVal = Convert.ToString(oComboBox.Selected.Value);

                            string CmbVal = "Ambas";
                            string cDesdeFecha = oForm.DataSources.UserDataSources.Item("dsFD").ValueEx;
                            if (cDesdeFecha == "" || cDesdeFecha == null) cDesdeFecha = "20100101";

                            string cHastaFecha = oForm.DataSources.UserDataSources.Item("dsFH").ValueEx;
                            if (cHastaFecha == "" || cHastaFecha == null) cHastaFecha = "20900101";

                            SAPbouiCOM.Grid oGrid = oForm.Items.Item("grdCli").Specific;
                            SAPbouiCOM.Grid oGrid2 = oForm.Items.Item("grdCli2").Specific;
                            oGrid2.DataTable = null;
                            
                            string sql = "SELECT T0.Itemcode AS 'Artículo', T0.ItemName AS 'Descripción',";
                            sql += " cast(ISNULL(dbo.EXO_CantidadFinal(T0.itemcode,'" + CmbVal + "', convert(DATETIME, '##DESDEFECHA', 112), convert(DATETIME, '##HASTAFECHA', 112)),0) as numeric(19, 2)) as 'Cantidad tratada',";
                            sql += " ISNULL(dbo.EXO_PrecioCalculadoExpert(T0.itemcode,'" + CmbVal + "', convert(DATETIME, '##DESDEFECHA', 112), convert(DATETIME, '##HASTAFECHA', 112)),0) as 'Precio Calculado Expert',";
                            sql += " T0.LstEvlPric as 'Precio SAP',";
                            sql += " ISNULL(dbo.EXO_PrecioCalculadoExpert(T0.itemcode,'" + CmbVal + "', convert(DATETIME, '##DESDEFECHA', 112), convert(DATETIME, '##HASTAFECHA', 112))*dbo.EXO_CantidadFinal(T0.itemcode,'" + CmbVal + "', convert(DATETIME, '##DESDEFECHA', 112), convert(DATETIME, '##HASTAFECHA', 112)),0) as 'Importe Calculado Expert',";
                            sql += " ISNULL(T0.LstEvlPric*dbo.EXO_CantidadFinal(T0.itemcode,'" + CmbVal + "', convert(DATETIME, '##DESDEFECHA', 112), convert(DATETIME, '##HASTAFECHA', 112)),0) as 'Importe SAP',";
                            sql += " (ISNULL(dbo.EXO_PrecioCalculadoExpert(T0.itemcode,'" + CmbVal + "',convert(DATETIME, '##DESDEFECHA', 112), convert(DATETIME, '##HASTAFECHA', 112))*dbo.EXO_CantidadFinal(T0.itemcode,'" + CmbVal + "', convert(DATETIME, '##DESDEFECHA', 112), convert(DATETIME, '##HASTAFECHA', 112)),0) - (ISNULL(T0.LstEvlPric*dbo.EXO_CantidadFinal(T0.itemcode,'" + CmbVal + "', convert(DATETIME, '##DESDEFECHA', 112), convert(DATETIME, '##HASTAFECHA', 112)),0))) as 'diferencia'";
                            sql += " FROM OITM T0";
                            sql += " WHERE T0.ItemCode BETWEEN '" + oForm.DataSources.UserDataSources.Item("dsArt").Value + "' AND '" + oForm.DataSources.UserDataSources.Item("dsArt2").Value + "'";
                                                        
                            if (cGrupoArticulo != null && cGrupoArticulo != "")
                            {
                                sql += " AND T0.ItmsGrpCod ='" + cGrupoArticulo + "'";
                            }

                            sql = sql.Replace("##DESDEFECHA", cDesdeFecha).Replace("##HASTAFECHA", cHastaFecha);

                            SAPbobsCOM.Recordset oRecAmi = Matriz.oGlobal.SQL.sqlComoRsB1(sql);
                            SAPbouiCOM.DataTable oTabla = oForm.DataSources.DataTables.Item("TablaPre");
                            
                            if (oRecAmi.RecordCount > 0)
                            {
                                Matriz.oGlobal.conexionSAP.SBOApp.SetStatusBarMessage("Calculando precios ... ", BoMessageTime.bmt_Short, false);
                                oTabla.ExecuteQuery(sql);
                                oGrid.DataTable = oTabla;
                            }
                            else
                            {
                                Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("No hay resgistros", 1, "Ok", "", "");
                            }
                        }
                        catch (Exception ex)
                        {
                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
                        }
                    }
                    #endregion

                    break;

                case BoEventTypes.et_CHOOSE_FROM_LIST:
                    if (!infoEvento.BeforeAction) // Antes de BeforeAction
                    {
                        try
                        {
                            switch (infoEvento.ItemUID)
                            {
                                #region Articulos
                                case "txtArt":
                                    oForm.DataSources.UserDataSources.Item("dsArt").Value = Convert.ToString(infoEvento.SelectedObjects.GetValue("ItemCode", 0));
                                    break;

                                case "txtArt2":
                                    oForm.DataSources.UserDataSources.Item("dsArt2").Value = Convert.ToString(infoEvento.SelectedObjects.GetValue("ItemCode", 0));
                                    break;

                                case "TXTGA":
                                    oForm.DataSources.UserDataSources.Item("dsGru").Value = Convert.ToString(infoEvento.SelectedObjects.GetValue("ItmsGrpNam", 0));
                                    oForm.DataSources.UserDataSources.Item("dsGru2").Value = Convert.ToString(infoEvento.SelectedObjects.GetValue("ItmsGrpCod", 0));
                                    break;
                                #endregion
                            }
                        }
                        catch (Exception ex)
                        {
                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
                        }
                }
                break;
            }
            return true;
        }

        public void MenuEvent(EXO_Generales.EXO_MenuEvent infoEvento)
        {
            //SAPbouiCOM.Form oForm = Matriz.oGlobal.conexionSAP.SBOApp.Forms.ActiveForm;

            //if (infoEvento.MenuUID == "1282")
            //{
            //    ((SAPbouiCOM.ComboBox)oForm.Items.Item("cmbLanza").Specific).Select(0, BoSearchKey.psk_Index);
            //    ((SAPbouiCOM.ComboBox)oForm.Items.Item("cmbGDR").Specific).Select(0, BoSearchKey.psk_Index);
            //    ((SAPbouiCOM.ComboBox)oForm.Items.Item("cmbTipInst").Specific).Select(0, BoSearchKey.psk_Index);
            //    ((SAPbouiCOM.ComboBox)oForm.Items.Item("cmbTipLin").Specific).Select(0, BoSearchKey.psk_Index);

            //    oForm.ActiveItem = "txtCode";
            //}

            //if (infoEvento.MenuUID == "1281")
            //{

            //}
            //EXO_CleanCOM.CLiberaCOM.Form(ref oForm);
        }

        public bool DataEvent(EXO_Generales.EXO_BusinessObjectInfo args)
        {
            SAPbouiCOM.Form oForm = Matriz.oGlobal.conexionSAP.SBOApp.Forms.Item(args.FormUID);

            switch (args.EventType)
            {
                case BoEventTypes.et_FORM_DATA_LOAD:
                    {
                        //if (!args.BeforeAction)
                        //{
                        //    PintoMatrizTodas(ref oForm);
                        //}

                    }
                    break;

                case BoEventTypes.et_FORM_DATA_ADD:
                case BoEventTypes.et_FORM_DATA_UPDATE:
                    {
                        
                    }
                    break;


            }

            EXO_CleanCOM.CLiberaCOM.Form(ref oForm);
            return true;
        }

        private void PintoMatriz(ref SAPbouiCOM.Form oForm)
        {
            //SAPbouiCOM.Matrix oMatLin = (SAPbouiCOM.Matrix)oForm.Items.Item("matExclu").Specific;

            //oForm.DataSources.UserDataSources.Item("dsRefEx").Value = "";
            //oForm.DataSources.UserDataSources.Item("dsDesc").Value = "";


            //for (int i = 1; i <= oMatLin.RowCount; i++)
            //{
            //    string cEquipo = ((SAPbouiCOM.EditText)oMatLin.GetCellSpecific("Col_0", i)).Value;
            //    string sql = "SELECT T0.Code AS 'Equipo',  T0.U_EXO_ItemCode AS 'Referencia', T1.ItemName AS 'DescArt' ";
            //    sql += " FROM [@EXO_CODFABRI] T0 INNER JOIN OITM T1 ON T0.U_EXO_ItemCode = T1.ItemCode ";
            //    sql += " WHERE T0.Code = '" + cEquipo + "'";
            //    SAPbobsCOM.Recordset oRecAmi = Matriz.oGlobal.SQL.sqlComoRsB1(sql);

            //    ((SAPbouiCOM.EditText)oMatLin.Columns.Item("Col_1").Cells.Item(i).Specific).Value = oRecAmi.Fields.Item("Referencia").Value;
            //    ((SAPbouiCOM.EditText)oMatLin.Columns.Item("Col_2").Cells.Item(i).Specific).Value = oRecAmi.Fields.Item("DescArt").Value;
            //}

        }


        private void ModificoCFL(ref SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.ChooseFromList oCFL;
            SAPbouiCOM.Conditions oCons;
            SAPbouiCOM.Condition oCon;

            #region CFL  ART
            oCFL = oForm.ChooseFromLists.Item("CFLArt");

            oCFL.SetConditions(null);
            oCons = oCFL.GetConditions();
            oCon = oCons.Add();
            oCon.Alias = "InvntItem";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "Y";
            oCFL.SetConditions(oCons);
            #endregion

            #region CFLART2
            SAPbouiCOM.ChooseFromList oCFL2;
            SAPbouiCOM.Conditions oCons2;
            SAPbouiCOM.Condition oCon2;
            

           
            oCFL2 = oForm.ChooseFromLists.Item("CFLArt2");
            oCFL2.SetConditions(null);
            oCons2 = oCFL2.GetConditions();
            oCon2 = oCons2.Add();
            oCon2.Alias = "InvntItem";
            oCon2.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon2.CondVal = "Y";
            oCFL2.SetConditions(oCons2);

            #endregion


        }
    }
}





