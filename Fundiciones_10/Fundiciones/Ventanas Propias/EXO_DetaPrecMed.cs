
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Cliente
{
    public class EXO_DetaPrecMed
    {        
        public EXO_DetaPrecMed()
        { }

        public EXO_DetaPrecMed(string cArt, string cDesdeFecha, string cHastaFecha, string cListaAlmacenes, string cTipoCalculo, string cTitulo)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.FormCreationParams oParametrosCreacion = (SAPbouiCOM.FormCreationParams)(Matriz.gen.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams));

            try
            {
                string strXML = Utilidades.LeoFichEmbebido("Formularios.EXO_DetaPrecMedExpert.srf");
                oParametrosCreacion.XmlData = strXML;
                oParametrosCreacion.UniqueID = "";
                oForm = Matriz.gen.SBOApp.Forms.AddEx(oParametrosCreacion);
                oForm.Title = cTitulo;
            }
            catch (Exception ex)
            {
                Matriz.gen.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
            }

            

            SAPbouiCOM.DataTable oTabla = oForm.DataSources.DataTables.Item("TablaReg");
            oTabla.ExecuteQuery(Especifico.sqlQueryCalculo("-1", cArt, cArt, cDesdeFecha, cHastaFecha, cListaAlmacenes, true, cTipoCalculo));
            if (!oTabla.IsEmpty)
            {
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("GrdCon").Specific;

                ((SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("x")).LinkedObjectType = "17";
                oGrid.Columns.Item("x").Width = 15;

                oGrid.Columns.Item("Orden").Visible = false;

                oGrid.Columns.Item("Tipo Doc").Type = BoGridColumnType.gct_ComboBox;
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).DisplayType = BoComboDisplayType.cdt_Description;
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("59", "Entradas");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("60", "Salidas");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("15", "Albaran Venta");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("20", "Albaran Compras");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("21", "Devolucion Compras");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("16", "Devolucion Venta");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("13", "Factura Venta");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("14", "Abono Venta");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("18", "Factura Compra");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("19", "Abono Compra");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("67", "Traslados");

                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).DisplayType = BoComboDisplayType.cdt_Description;
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("I20", "Inicial (A)");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("I59", "Inicial (E)");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("I60", "Inicial X");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("I15", "Inicial X");                
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("I21", "Inicial X");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("I16", "Inicial X");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("I13", "Inicial X");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("I14", "Inicial X");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("I18", "Inicial X");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("I19", "Inicial X");
                ((SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item("Tipo Doc")).ValidValues.Add("I67", "Inicial X");                
                oGrid.AutoResizeColumns();
            }

            oForm.Visible = true;

        }


        public bool ItemEvent(ItemEvent infoEvento)
        {
            switch (infoEvento.EventType)
            {
                case BoEventTypes.et_FORM_RESIZE:
                    if (!infoEvento.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = Matriz.gen.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);
                        SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("GrdCon").Specific;
                        oGrid.Columns.Item("x").Width = 15;

                    }
                    break;


                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                    if (infoEvento.ItemUID == "GrdCon" && infoEvento.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = Matriz.gen.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);

                        //infoEvento
                        SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("GrdCon").Specific;

                        string cObjeto = oGrid.DataTable.GetValue("Tipo Doc", oGrid.GetDataTableRowIndex(infoEvento.Row));
                        if (cObjeto == "I0") return false;

                        if (cObjeto.Substring(0, 1) == "I") cObjeto = cObjeto.Substring(1);
                        ((SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("x")).LinkedObjectType = cObjeto;
                        
                        return true;                     
                    }
                    break;
            }

            return true;
        }

    }
}