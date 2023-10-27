using System;
using System.Collections.Generic;
using System.Xml;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;


namespace Cliente
{
    public class EXO_CalculoPrecioMedio
    {
        public enum Estado { ParaConsulta, ParaActualizar }


        public EXO_CalculoPrecioMedio()
        { }

        public EXO_CalculoPrecioMedio(bool lCreacion)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.FormCreationParams oParametrosCreacion = (SAPbouiCOM.FormCreationParams)(Matriz.gen.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams));

            try
            {
                string strXML = Utilidades.LeoFichEmbebido("Formularios.EXO_CalculoPrecioMedio.srf");
                oParametrosCreacion.XmlData = strXML;
                oParametrosCreacion.UniqueID = "";
                oForm = Matriz.gen.SBOApp.Forms.AddEx(oParametrosCreacion);

                #region Lleno combo de grupos
                SAPbouiCOM.ValidValues oValores = ((SAPbouiCOM.ComboBox)oForm.Items.Item("cmbGrupo").Specific).ValidValues;
                Matriz.gen.funcionesUI.cargaCombo(oValores, sqlGrupos());
                ((SAPbouiCOM.ComboBox)oForm.Items.Item("cmbGrupo").Specific).ExpandType = BoExpandType.et_DescriptionOnly;
                #endregion


                ((SAPbouiCOM.OptionBtn)oForm.Items.Item("opFecS").Specific).GroupWith("opFecC");
                ((SAPbouiCOM.OptionBtn)oForm.Items.Item("opFecC").Specific).ValOff = "S";
                ((SAPbouiCOM.OptionBtn)oForm.Items.Item("opFecC").Specific).ValOn = "C";
                ((SAPbouiCOM.OptionBtn)oForm.Items.Item("opFecS").Specific).ValOff = "C";
                ((SAPbouiCOM.OptionBtn)oForm.Items.Item("opFecS").Specific).ValOn = "S";


                EstadoObjeto(ref oForm, Estado.ParaConsulta);
                Limpiar(ref oForm);

                oForm.Visible = true;

                //oForm.ActiveItem = "TXTFD"
                if (Matriz.gen.refDi.OGEN.valorVariable("EXO_LISTAPRECIOS") == "")
                {
                    Matriz.gen.SBOApp.MessageBox("La variable EXO_LISTAPRECIOS esta sin definir", 1, "Ok", "", "");
                    oForm.Items.Item("btnCons").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False);
                    oForm.Items.Item("bt_GLP").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False);
                }

                Utilidades.DeshabilitoMenus(ref oForm);


                oForm.ActiveItem = "TXTFD";
            }
            catch (Exception ex)
            {
                Matriz.gen.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public bool ItemEvent(ItemEvent infoEvento)
        {
            switch (infoEvento.EventType)
            {
                case BoEventTypes.et_DOUBLE_CLICK:
                    if (infoEvento.ItemUID == "GrdAlma" && !infoEvento.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = Matriz.gen.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);
                        RellenoAlmacenes(ref oForm, true);
                    }
                    break;

                case BoEventTypes.et_FORM_RESIZE:
                    if (!infoEvento.BeforeAction)
                    {
                        #region Cambio el tamaño
                        SAPbouiCOM.Form oForm = Matriz.gen.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);
                        if (oForm.Visible)
                        {
                            oForm.Items.Item("grdCli").Top = oForm.Items.Item("cmbGrupo").Top + 100;
                            oForm.Items.Item("grdCli").Height = oForm.Items.Item("bt_GLP").Top - oForm.Items.Item("grdCli").Top - 15;

                            oForm.Items.Item("GrdAlma").Top = oForm.Items.Item("TXTFH").Top;
                            oForm.Items.Item("GrdAlma").Height = oForm.Items.Item("grdCli").Top - oForm.Items.Item("TXTFH").Top - 10;

                            //
                            if (((SAPbouiCOM.Grid)oForm.Items.Item("GrdAlma").Specific).Columns.Count > 0)
                            {
                                ((SAPbouiCOM.Grid)oForm.Items.Item("GrdAlma").Specific).Columns.Item("x").Width = 15;
                            }

                            if (((SAPbouiCOM.Grid)oForm.Items.Item("grdCli").Specific).Columns.Count > 0)
                            {
                                ((SAPbouiCOM.Grid)oForm.Items.Item("grdCli").Specific).Columns.Item("x").Width = 15;
                            }

                        }
                        #endregion
                    }
                    break;

                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                    if (infoEvento.ItemUID == "grdCli" && infoEvento.ColUID == "x" && infoEvento.BeforeAction)
                    {
                        #region Flecha del detalle de precio
                        SAPbouiCOM.Form oForm = Matriz.gen.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);
                        SAPbouiCOM.Grid oGrid = oForm.Items.Item("grdCli").Specific;

                        string cArt = oGrid.DataTable.GetValue("Articulo", oGrid.GetDataTableRowIndex(infoEvento.Row));
                        string cNomArt = oGrid.DataTable.GetValue("Descripcion", oGrid.GetDataTableRowIndex(infoEvento.Row));

                        string cDesdeFecha = oForm.DataSources.UserDataSources.Item("dsFD").ValueEx;
                        if (cDesdeFecha == "") cDesdeFecha = "20000101";

                        string cHastaFecha = oForm.DataSources.UserDataSources.Item("dsFH").ValueEx;
                        if (cHastaFecha == "") cHastaFecha = "20490101";

                        //Almacenes
                        string cListaAlmacenes = DoyListaAlmacenes(ref oForm);
                        string cTipoCalculo = (((SAPbouiCOM.OptionBtn)oForm.Items.Item("opFecS").Specific).Selected ? "S" : "C");

                        //
                        string cTitulo = "Detalle Precio Medio " + cArt + " " + cNomArt;
                        EXO_DetaPrecMed fDetaPrec = new EXO_DetaPrecMed(cArt, cDesdeFecha, cHastaFecha, cListaAlmacenes, cTipoCalculo, cTitulo);
                        fDetaPrec = null;
                        #endregion

                        return false;
                    }
                    break;

                case BoEventTypes.et_ITEM_PRESSED:
                    if (infoEvento.ItemUID == "btnSalir" && !infoEvento.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = Matriz.gen.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);

                        oForm.Close();
                    }

                    if (infoEvento.ItemUID == "bt_GLP" && !infoEvento.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = Matriz.gen.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);
                        SAPbouiCOM.Grid oGrid = oForm.Items.Item("grdCli").Specific;
                        SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)Matriz.gen.compañia.GetBusinessObject(BoObjectTypes.BoRecordset);

                        if (!oGrid.DataTable.IsEmpty)
                        {
                            #region Recorro el grid
                            List<Matriz.Mensajes> ListaMensajes = new List<Matriz.Mensajes>();
                            string cListaPrec = Matriz.gen.refDi.OGEN.valorVariable("EXO_LISTAPRECIOS");
                            int nMulti = Convert.ToInt32("1" + "".PadRight(VarGlobal.PriceDec, '0'));
                            string cMenError = "";

                            System.Globalization.NumberFormatInfo nfi = new System.Globalization.NumberFormatInfo();
                            nfi.NumberDecimalSeparator = ".";
                            nfi.NumberGroupSeparator = "";
                            nfi.NumberDecimalDigits = 6;

                            #region xml grid
                            System.Xml.XmlDocument oXmlGridTabla = new System.Xml.XmlDocument();
                            {
                                string sXmlTabla = oGrid.DataTable.SerializeAsXML(BoDataTableXmlSelect.dxs_All);
                                oXmlGridTabla.LoadXml(sXmlTabla);
                            }
                            XmlNodeList oXmlNodesGridTabla = oXmlGridTabla.SelectNodes("/DataTable/Rows/Row/Cells/Cell[./ColumnUid='Bloquear Precio Expert' and ./Value='N']/..");
                            #endregion


                            foreach (XmlNode oNodo in oXmlNodesGridTabla)
                            {
                                string cValorXML = oNodo.SelectNodes("Cell[./ColumnUid='Precio Calculado Expert']").Item(0).SelectSingleNode("Value").InnerText;
                                string cArticulo = oNodo.SelectNodes("Cell[./ColumnUid='Articulo']").Item(0).SelectSingleNode("Value").InnerText;

                                int nValor = (int)Math.Truncate(Convert.ToDouble(cValorXML, nfi) * nMulti);
                                string cPrecAct = nValor.ToString() + ".0/" + nMulti.ToString() + ".0";


                                string sqlUPD = "UPDATE ITM1 SET Price = " + cPrecAct + " WHERE PriceList = '" + cListaPrec + "' AND ItemCode = '" + cArticulo + "'";

                                Utilidades.EjecutoSQL(sqlUPD, ref cMenError);
                                if (cMenError != "")
                                {
                                    #region Por si hay error
                                    Matriz.Mensajes AuxMensa;
                                    AuxMensa.Clave = "";
                                    AuxMensa.Mensaje = "ERROR. No se puedo actualizar articulo " + cArticulo;
                                    AuxMensa.Objeto = "";
                                    AuxMensa.TipoObjeto = Matriz.ClasesObjetos.Caracter;
                                    ListaMensajes.Add(AuxMensa);

                                    AuxMensa.Clave = "";
                                    AuxMensa.Mensaje = cMenError;
                                    AuxMensa.Objeto = "";
                                    AuxMensa.TipoObjeto = Matriz.ClasesObjetos.Caracter;
                                    ListaMensajes.Add(AuxMensa);

                                    AuxMensa.Clave = "";
                                    AuxMensa.Mensaje = "";
                                    AuxMensa.Objeto = "";
                                    AuxMensa.TipoObjeto = Matriz.ClasesObjetos.Caracter;
                                    ListaMensajes.Add(AuxMensa);
                                    #endregion
                                }

                                Matriz.gen.SBOApp.SetStatusBarMessage("Articulo " + cArticulo + " actualizado", BoMessageTime.bmt_Short, false);
                            }

                            Matriz.gen.SBOApp.SetStatusBarMessage("Proceso terminado", BoMessageTime.bmt_Short, false);

                            if (ListaMensajes.Count > 0)
                            {
                                EXO_VENREG fVenReg = new EXO_VENREG(ListaMensajes, "Errores actualizacion lista precios margen");
                                fVenReg = null;
                            }

                            Limpiar(ref oForm);
                            EstadoObjeto(ref oForm, Estado.ParaConsulta);
                            ((SAPbouiCOM.Button)oForm.Items.Item("btnCons").Specific).Caption = "Consultar";
                            #endregion
                        }
                    }

                    if (infoEvento.ItemUID == "btnCons" && infoEvento.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = Matriz.gen.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);
                        if (((SAPbouiCOM.Button)oForm.Items.Item("btnCons").Specific).Caption == "Consultar")
                        {
                            #region Consultar
                            try
                            {
                                #region Recojo los filtros
                                string cGrupoArticulo = "";
                                try
                                {
                                    cGrupoArticulo = ((SAPbouiCOM.ComboBox)oForm.Items.Item("cmbGrupo").Specific).Selected.Value;
                                }
                                catch
                                { }


                                string cDesdeFecha = oForm.DataSources.UserDataSources.Item("dsFD").ValueEx;
                                if (cDesdeFecha == "") cDesdeFecha = "20000101";

                                string cHastaFecha = oForm.DataSources.UserDataSources.Item("dsFH").ValueEx;
                                if (cHastaFecha == "") cHastaFecha = "20490101";

                                string cDesdeArt = oForm.DataSources.UserDataSources.Item("dsArt").Value;
                                string cHastaArt = oForm.DataSources.UserDataSources.Item("dsArt2").Value;
                                #endregion

                                //Almacenes
                                string cListaAlmacenes = DoyListaAlmacenes(ref oForm);
                                if (cListaAlmacenes == "")
                                {
                                    Matriz.gen.SBOApp.SetStatusBarMessage("Ha de seleccionar algun almacen", BoMessageTime.bmt_Short, true);
                                    return false;
                                }


                                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grdCli").Specific;

                                string cTipoCalculo = (((SAPbouiCOM.OptionBtn)oForm.Items.Item("opFecS").Specific).Selected ? "S" : "C");
                                string sql = Especifico.sqlQueryCalculo(cGrupoArticulo, cDesdeArt, cHastaArt, cDesdeFecha, cHastaFecha, cListaAlmacenes, false, cTipoCalculo);
                                Matriz.gen.SBOApp.SetStatusBarMessage("Calculando precios ... ", BoMessageTime.bmt_Short, false);
                                oForm.Freeze(true);
                                oForm.DataSources.DataTables.Item("TablaPre").ExecuteQuery(sql);
                                if (oForm.DataSources.DataTables.Item("TablaPre").IsEmpty)
                                {
                                    oForm.Freeze(false);
                                    Matriz.gen.SBOApp.MessageBox("No hay resgitros", 1, "Ok", "", "");
                                }
                                else
                                {
                                    ((SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Articulo")).LinkedObjectType = "4";
                                    ((SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("x")).LinkedObjectType = "17";

                                    oGrid.AutoResizeColumns();
                                    oGrid.Columns.Item("x").Width = 15;
                                    oGrid.Columns.Cast<GridColumn>().ToList().ForEach(c => { c.TitleObject.Sortable = true; });

                                    #region columna si/no
                                    oGrid.Columns.Item("Bloquear Precio Expert").Type = BoGridColumnType.gct_ComboBox;
                                    ComboBoxColumn oComboCol = (ComboBoxColumn)oGrid.Columns.Item("Bloquear Precio Expert");
                                    oComboCol.ValidValues.Add("Y", "Si");
                                    oComboCol.ValidValues.Add("N", "No");
                                    oComboCol.DisplayType = BoComboDisplayType.cdt_Description;
                                    oComboCol.ExpandType = BoExpandType.et_DescriptionOnly;
                                    #endregion

                                    Matriz.gen.SBOApp.SetStatusBarMessage("Calculos realizados", BoMessageTime.bmt_Short, false);
                                }

                            }
                            catch (Exception ex)
                            {
                                Matriz.gen.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
                            }
                            #endregion

                            EstadoObjeto(ref oForm, Estado.ParaActualizar);

                            ((SAPbouiCOM.Button)oForm.Items.Item("btnCons").Specific).Caption = "Limpiar";

                            oForm.Freeze(false);
                        }
                        else
                        {
                            EstadoObjeto(ref oForm, Estado.ParaConsulta);
                            Limpiar(ref oForm);
                            ((SAPbouiCOM.Button)oForm.Items.Item("btnCons").Specific).Caption = "Consultar";


                            oForm.ActiveItem = "TXTFD";
                        }
                    }
                    break;

                case BoEventTypes.et_CHOOSE_FROM_LIST:
                    if (!infoEvento.BeforeAction) // Antes de BeforeAction
                    {
                        try
                        {
                            SAPbouiCOM.Form oForm = Matriz.gen.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)infoEvento;
                            switch (infoEvento.ItemUID)
                            {
                                #region Articulos
                                case "txtArt":
                                    oForm.DataSources.UserDataSources.Item("dsArt").Value = Convert.ToString(oCFLEvento.SelectedObjects.GetValue("ItemCode", 0));
                                    break;

                                case "txtArt2":
                                    oForm.DataSources.UserDataSources.Item("dsArt2").Value = Convert.ToString(oCFLEvento.SelectedObjects.GetValue("ItemCode", 0));
                                    break;
                                    #endregion
                            }
                        }
                        catch (Exception ex)
                        {
                            Matriz.gen.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
                        }
                    }
                    break;
            }
            return true;
        }

        public void MenuEvent(MenuEvent infoEvento)
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

        public bool DataEvent(BusinessObjectInfo args)
        {
            SAPbouiCOM.Form oForm = Matriz.gen.SBOApp.Forms.Item(args.FormUID);

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

            EXO_CleanCOM.CLiberaCOM.Form(oForm);
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


        private string DoyListaAlmacenes(ref SAPbouiCOM.Form oForm)
        {
            #region Almacenes
            string cListaAlmacenes = "";

            string sXmlTabla = oForm.DataSources.DataTables.Item("TablaAlma").SerializeAsXML(BoDataTableXmlSelect.dxs_All);
            System.Xml.XmlDocument oXmlGridTabla = new System.Xml.XmlDocument();
            oXmlGridTabla.LoadXml(sXmlTabla);

            string sXPath = "/DataTable/Rows/Row/Cells/Cell[./ColumnUid='x' and ./Value='Y']";
            XmlNodeList oXmlNodesGridTabla = oXmlGridTabla.SelectNodes(sXPath);
            foreach (XmlNode oNodo in oXmlNodesGridTabla)
            {
                string cAlmacen = oNodo.ParentNode.SelectNodes("Cell[./ColumnUid='Almacen']").Item(0).SelectSingleNode("Value").InnerText;
                cListaAlmacenes += (cAlmacen + ",");
            }

            if (cListaAlmacenes.Length > 0)
            {
                cListaAlmacenes = cListaAlmacenes.Substring(0, cListaAlmacenes.Length - 1);
            }
            #endregion

            return cListaAlmacenes;
        }


        private void EstadoObjeto(ref SAPbouiCOM.Form oForm, Estado oEstado)
        {
            oForm.Items.Item("TXTFD").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, oEstado == Estado.ParaActualizar ? BoModeVisualBehavior.mvb_False : BoModeVisualBehavior.mvb_True);
            oForm.Items.Item("TXTFH").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, oEstado == Estado.ParaActualizar ? BoModeVisualBehavior.mvb_False : BoModeVisualBehavior.mvb_True);
            oForm.Items.Item("txtArt").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, oEstado == Estado.ParaActualizar ? BoModeVisualBehavior.mvb_False : BoModeVisualBehavior.mvb_True);
            oForm.Items.Item("txtArt2").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, oEstado == Estado.ParaActualizar ? BoModeVisualBehavior.mvb_False : BoModeVisualBehavior.mvb_True);
            oForm.Items.Item("txtArt2").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, oEstado == Estado.ParaActualizar ? BoModeVisualBehavior.mvb_False : BoModeVisualBehavior.mvb_True);
            oForm.Items.Item("cmbGrupo").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, oEstado == Estado.ParaActualizar ? BoModeVisualBehavior.mvb_False : BoModeVisualBehavior.mvb_True);

            oForm.Items.Item("opFecS").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, oEstado == Estado.ParaActualizar ? BoModeVisualBehavior.mvb_False : BoModeVisualBehavior.mvb_True);
            oForm.Items.Item("opFecC").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, oEstado == Estado.ParaActualizar ? BoModeVisualBehavior.mvb_False : BoModeVisualBehavior.mvb_True);


            oForm.Items.Item("bt_GLP").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, oEstado == Estado.ParaActualizar ? BoModeVisualBehavior.mvb_True : BoModeVisualBehavior.mvb_False);


            oForm.Items.Item("GrdAlma").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, oEstado == Estado.ParaActualizar ? BoModeVisualBehavior.mvb_False : BoModeVisualBehavior.mvb_True);
        }


        private void Limpiar(ref SAPbouiCOM.Form oForm)
        {
            oForm.DataSources.UserDataSources.Item("dsFD").Value = "";
            oForm.DataSources.UserDataSources.Item("dsFH").Value = "";
            oForm.DataSources.UserDataSources.Item("dsArt").Value = "";
            oForm.DataSources.UserDataSources.Item("dsArt2").Value = "";
            oForm.DataSources.UserDataSources.Item("dsOpt").Value = "C";

            ((SAPbouiCOM.ComboBox)oForm.Items.Item("cmbGrupo").Specific).Select("-1", BoSearchKey.psk_ByValue);

            ((SAPbouiCOM.Grid)oForm.Items.Item("grdCli").Specific).DataTable.Clear();

            RellenoAlmacenes(ref oForm, false);
        }

        public void RellenoAlmacenes(ref SAPbouiCOM.Form oForm, bool lMarca)
        {
            //Relleno almacenes
            SAPbouiCOM.DataTable oTabAlma = oForm.DataSources.DataTables.Item("TablaAlma");
            oTabAlma.ExecuteQuery(sqlAlmacenes(lMarca));
            if (!oTabAlma.IsEmpty)
            {
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("GrdAlma").Specific;
                oGrid.Columns.Item("x").Type = BoGridColumnType.gct_CheckBox;

                oGrid.Columns.Item("x").Editable = true;
                oGrid.Columns.Item("x").Width = 15;
                oGrid.Columns.Item("Almacen").Editable = false;
                oGrid.Columns.Item("Nombre").Editable = false;

                oGrid.AutoResizeColumns();
            }
        }


        public string sqlGrupos()
        {
            string sql = "(SELECT - 1 AS 'Grupo', 'Ninguno seleccionado') ";
            sql += " UNION all ";
            sql += " (SELECT T0.ItmsGrpCod AS 'Grupo', T0.ItmsGrpNam AS 'Nombre' FROM OITB T0 ) ";

            return sql;

        }


        public string sqlAlmacenes(bool lMarcados = false)
        {
            string sql = "SELECT '##MARCA' AS 'x', T0.WhsCode AS 'Almacen', T0.WhsName AS 'Nombre' from OWHS T0 ORDER BY T0.WhsCode";
            sql = sql.Replace("##MARCA", lMarcados ? "Y" : "N");

            return sql;

        }
    }
}





