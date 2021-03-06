using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Cliente
{
    public class Matriz : EXO_UIAPI.EXO_DLLBase
    {
        public static  EXO_UIAPI.EXO_UIAPI oGlobal;
        public static Type TypeMatriz;
        public static bool bModalDESCOMPUESTO = false;
        //public static CrystalDecisions.CrystalReports.Engine.ReportDocument crReport = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
        public static bool LanzarImpresionCrystal = false;
        public static Object ThisMatriz;

        public Matriz(EXO_UIAPI.EXO_UIAPI gen, Boolean act, Boolean usalicencia, int idAddon)
            : base(gen, act, usalicencia, idAddon)
        {
            oGlobal = this.objGlobal;
            TypeMatriz = this.GetType();

            Object ThisMatriz = this;

            if (act)
            {
                Utilidades.NuevoReportType("Listados Expert", "Listados Expert", "LISTEXPERT", "MnuXList", true);
            }

            SAPbobsCOM.Recordset oRec = null;

            #region Decimales de la aplicacion
            oRec = oGlobal.refDi.SQL.sqlComoRsB1("SELECT T0.SumDec, T0.PriceDec, T0.RateDec, T0.QtyDec, T0.PercentDec, T0.MeasureDec, T0.ThousSep, T0.DecSep FROM OADM T0");
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

            #region Cargo menus
            string mXML = Matriz.oGlobal.funciones.leerEmbebido(TypeMatriz, "xMenuFundiciones.xml");
            objGlobal.SBOApp.LoadBatchActions(mXML);
            string res = objGlobal.SBOApp.GetLastBatchResults();
            #endregion

        }

        public override SAPbouiCOM.EventFilters filtros()
        {
            SAPbouiCOM.EventFilters oFilter = new SAPbouiCOM.EventFilters();

            #region Mando filtros
            try
            {
                Type Tipo = this.GetType();
                string fXML = Matriz.oGlobal.funciones.leerEmbebido(Tipo, "xFiltrosFundiciones.xml");
                oFilter.LoadFromXML(fXML);
            }
            catch (Exception ex)
            {
                objGlobal.SBOApp.MessageBox("Error en carga de filtros Fundiciones", 1, "Ok", "", "");
                oFilter = null;
            }
            #endregion
            return oFilter;
        }

        public  bool SBOApp_ItemEvent(ref ItemEvent infoEvento)
        {
            bool lRetorno = true;


            if (infoEvento.FormTypeEx == "EXOCPM")
            {
                EXO_CalculoPrecioMedio fEXOCPM = new EXO_CalculoPrecioMedio();
                lRetorno = fEXOCPM.ItemEvent(infoEvento);
                fEXOCPM = null;
            }


            return lRetorno;
        }

        public  bool SBOApp_MenuEvent(ref MenuEvent infoMenuEvent)
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

        public  bool DataEvent(BusinessObjectInfo InfoEvento)
        {
            bool lRetorno = true;

            return lRetorno;
        }

        public override XmlDocument menus()
        {
            throw new NotImplementedException();
        }
    }
}
