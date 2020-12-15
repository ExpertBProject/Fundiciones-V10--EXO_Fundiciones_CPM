using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Cliente
{

    public class Matriz : EXO_Generales.EXO_DLLBase
    {
        public static EXO_Generales.EXO_General oGlobal;
        public static Type TypeMatriz;
        public static bool bModalDESCOMPUESTO = false;
        //public static CrystalDecisions.CrystalReports.Engine.ReportDocument crReport = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
        public static bool LanzarImpresionCrystal = false;
        public static Object ThisMatriz;

        public Matriz(EXO_Generales.EXO_General gen, Boolean act)
            : base(ref gen, act)
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
            oRec = oGlobal.SQL.sqlComoRsB1("SELECT T0.SumDec, T0.PriceDec, T0.RateDec, T0.QtyDec, T0.PercentDec, T0.MeasureDec, T0.ThousSep, T0.DecSep FROM OADM T0");
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
            string mXML = Matriz.oGlobal.Functions.leerEmbebido(ref TypeMatriz, "xMenuFundiciones.xml");
            SboApp.LoadBatchActions(mXML);
            string res = SboApp.GetLastBatchResults();
            #endregion

        }

        public override SAPbouiCOM.EventFilters filtros()
        {
            SAPbouiCOM.EventFilters oFilter = new SAPbouiCOM.EventFilters();

            #region Mando filtros
            try
            {
                Type Tipo = this.GetType();
                string fXML = Matriz.oGlobal.Functions.leerEmbebido(ref Tipo, "xFiltrosFundiciones.xml");
                oFilter.LoadFromXML(fXML);
            }
            catch (Exception ex)
            {
                this.SboApp.MessageBox("Error en carga de filtros Fundiciones", 1, "Ok", "", "");
                oFilter = null;
            }
            #endregion
            return oFilter;
        }

        public override bool SBOApp_ItemEvent(ref EXO_Generales.EXO_infoItemEvent infoEvento)
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

        public override bool SBOApp_MenuEvent(ref EXO_Generales.EXO_MenuEvent infoMenuEvent)
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

        public override bool SBOApp_FormDataEvent(ref EXO_Generales.EXO_BusinessObjectInfo infoDataEvent)
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
