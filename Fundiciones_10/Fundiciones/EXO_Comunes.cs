using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Text.RegularExpressions;
using System.Reflection;
using Microsoft.VisualBasic.CompilerServices;



namespace Cliente
{
    public class VarGlobal
    {
        public static int SumDec;
        public static int PriceDec;
        public static int RateDec;
        public static int QtyDec;
        public static int PercentDec;
        public static int MeasureDec;
        public static string SepMill;
        public static string SepDec;
    }

    public class Conversiones
    {
      
      public static double ValueSAPToDoubleSistema(string Texto)
            {
                string Cadena = Texto;

                double Valor = 0.0;
                System.Globalization.NumberFormatInfo nfi = new System.Globalization.NumberFormatInfo();
                string SepDecSistema = nfi.NumberGroupSeparator;
                //string SepMilSistema = nfi.NumberDecimalSeparator;

                //En pantalla el separador decimal es .
                if (SepDecSistema != ".")
                {
                    Cadena = Cadena.Replace('.', ',');
                }
                double.TryParse(Cadena, out Valor);
                return Valor;
            }
      
      public static double StringSAPToDoubleSistema(string Texto)
      {
          double nRetorno = 0;

          //Quito la moneda y el sep miles
          string Cadena = Texto;
          Cadena = Cadena.Replace(VarGlobal.SepMill, "");

          System.Globalization.NumberFormatInfo nfi = new System.Globalization.NumberFormatInfo();
          string SepDecSistema = nfi.NumberGroupSeparator;
          Cadena = Cadena.Replace(VarGlobal.SepDec, SepDecSistema);
          double.TryParse(Cadena, out nRetorno);          

          return nRetorno;
      }

    }
        
    public class Utilidades
    {                   
        public static string LeoQueryFich(string cNomQueryIncrustada, Type Tipo)

        {             
             string cQuery = Matriz.gen.funciones.leerEmbebido(Tipo, cNomQueryIncrustada);
             cQuery = cQuery.Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");

             return cQuery;
        }

        public static string LeoFichEmbebido(string cFichEmbebido)
        {
            string result = "";
            try
            {
                Type tipo = Matriz.TypeMatriz;
                Assembly assembly = tipo.Assembly;
                StreamReader streamReader = new StreamReader(tipo.Assembly.GetManifestResourceStream(tipo.Namespace + "." + cFichEmbebido));
                result = streamReader.ReadToEnd();
                result = result.Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");
                streamReader.Close();
            }
            catch (Exception expr_40)
            {
                ProjectData.SetProjectError(expr_40);
                ProjectData.ClearProjectError();
            }

            return result;
        }

        public static void EjecutoSQL(string sqlUPD, ref string cMenError)
        {
            SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)Matriz.gen.compañia.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oRec.DoQuery(sqlUPD);
            }
            catch (Exception EX)
            {
                cMenError = EX.Message;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
                oRec = null;
            }
        }

        public static void DeshabilitoMenus(ref SAPbouiCOM.Form oForm)
        {
            oForm.EnableMenu("1281", false);
            oForm.EnableMenu("1282", false);
            oForm.EnableMenu("1290", false);
            oForm.EnableMenu("1288", false);
            oForm.EnableMenu("1289", false);
            oForm.EnableMenu("1291", false);
        }

    }

}

