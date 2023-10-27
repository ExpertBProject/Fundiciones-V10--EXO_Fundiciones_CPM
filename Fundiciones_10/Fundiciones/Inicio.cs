using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Cliente
{

    public class Matriz : EXO_UIAPI.EXO_DLLBase
    {
        public enum ClasesObjetos { Numero, Caracter }

        public struct Mensajes
        {
            public string Clave;
            public ClasesObjetos TipoObjeto;
            public string Objeto;
            public string Mensaje;
        }

        public static EXO_UIAPI.EXO_UIAPI gen;
        public static EXO_DIAPI.EXO_DIAPI conexionSAP;

        public static Type TypeMatriz;
        public static bool bModalDESCOMPUESTO = false;
        public static bool LanzarImpresionCrystal = false;
        public static Object ThisMatriz;

        public Matriz(EXO_UIAPI.EXO_UIAPI gen, Boolean act, Boolean usalicencia, int idAddon)
                  : base(gen, act, usalicencia, idAddon)
        {
            Matriz.gen = this.objGlobal;
            TypeMatriz = this.GetType();

            Object ThisMatriz = this;             
            if (act)
            {
                string cMen = "";

                #region UDFs itm1
                string fich = Utilidades.LeoFichEmbebido("XML_DB.UDFs_ITM1.xml");
                if (!objGlobal.refDi.comunes.LoadBDFromXML(fich, cMen))
                {
                    aplicacionB1.MessageBox(cMen, 1, "Ok", "", "");
                }
                else
                {
                    gen.SBOApp.SetStatusBarMessage("Actualizado UDF ITM1", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                #endregion
            }

            SAPbobsCOM.Recordset oRec = null;
            #region Decimales de la aplicacion
            oRec = Matriz.gen.refDi.SQL.sqlComoRsB1("SELECT T0.SumDec, T0.PriceDec, T0.RateDec, T0.QtyDec, T0.PercentDec, T0.MeasureDec, T0.ThousSep, T0.DecSep FROM OADM T0");
            VarGlobal.SumDec = Convert.ToInt32(oRec.Fields.Item(0).Value);
            VarGlobal.PriceDec = Convert.ToInt32(oRec.Fields.Item(1).Value);
            VarGlobal.RateDec = Convert.ToInt32(oRec.Fields.Item(2).Value);
            VarGlobal.QtyDec = Convert.ToInt32(oRec.Fields.Item(3).Value);
            VarGlobal.PercentDec = Convert.ToInt32(oRec.Fields.Item(4).Value);
            VarGlobal.MeasureDec = Convert.ToInt32(oRec.Fields.Item(5).Value);
            VarGlobal.SepMill = Convert.ToString(oRec.Fields.Item(6).Value);
            VarGlobal.SepDec = Convert.ToString(oRec.Fields.Item(7).Value);
            #endregion
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            oRec = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public override SAPbouiCOM.EventFilters filtros()
        {
            SAPbouiCOM.EventFilters oFilter = new SAPbouiCOM.EventFilters();

            #region Mando filtros
            try
            {
                Type Tipo = this.GetType();
                string fXML = Utilidades.LeoFichEmbebido("xFiltrosFundiciones_PM.xml");
                
                oFilter.LoadFromXML(fXML);
            }
            catch (Exception ex)
            {
                Matriz.gen.SBOApp.MessageBox("Error en carga de filtros Fundiciones PM", 1, "Ok", "", "");
                oFilter = null;
            }
            #endregion
            return oFilter;
        }

        public override System.Xml.XmlDocument menus()
        {

            System.Xml.XmlDocument menu = new System.Xml.XmlDocument();
            string mXML = "";

            if (Matriz.gen.SBOApp.ClientType == BoClientType.ct_Desktop)
            {
                mXML = Utilidades.LeoFichEmbebido("xMenuFundiciones_PM.xml");

                //mXML = Matriz.gen.funciones.leerEmbebido(this.GetType(), "xMenuFundiciones_PM.xml");                
                menu.LoadXml(mXML);
                return menu;
            }
            else return null;
        }

        public override bool SBOApp_ItemEvent(ItemEvent infoEvento)
        {
            bool lRetorno = true;
            if (infoEvento.FormTypeEx == "EXOCPM")
            {
                EXO_CalculoPrecioMedio fEXOCPM = new EXO_CalculoPrecioMedio();
                lRetorno = fEXOCPM.ItemEvent(infoEvento);
                fEXOCPM = null;                
            } 
            if (infoEvento.FormTypeEx == "DETAPREC")
            {
                EXO_DetaPrecMed fDetaPrec = new EXO_DetaPrecMed();
                lRetorno = fDetaPrec.ItemEvent(infoEvento);
                fDetaPrec = null;
            }
            if (infoEvento.FormTypeEx == "VENREGPM")
            {
                EXO_VENREG fVenReg = new EXO_VENREG();
                lRetorno = fVenReg.ItemEvent(infoEvento);
                fVenReg = null;
            }
            if(infoEvento.FormTypeEx == "157")
            {
                EXO_157 fListaPrecios = new EXO_157();
                lRetorno = fListaPrecios.ItemEvent(infoEvento);
                fListaPrecios = null;
            }
            return lRetorno;
        }

        public override bool SBOApp_MenuEvent(MenuEvent infoMenuEvent)
        {
            bool lRetorno = true;

            switch (infoMenuEvent.MenuUID)
            {
                case "mCalPre":
                    // Calculo de Precio Medio
                    if (!infoMenuEvent.BeforeAction)
                    {
                        EXO_CalculoPrecioMedio fPrecio = new EXO_CalculoPrecioMedio(true);
                        fPrecio = null;
                    }
                    break;
            }

            return lRetorno;
        }

        public override bool SBOApp_FormDataEvent(BusinessObjectInfo infoDataEvent)
        {
            bool lRetorno = true;

            //if (infoDataEvent.FormTypeEx == "133")
            //{
            //    EXO_133 f133;
            //    f133 = new EXO_133();
            //    lRetorno = f133.DataEvent(infoDataEvent);
            //    f133 = null;
            //}

            return lRetorno;
        }
    }
}
