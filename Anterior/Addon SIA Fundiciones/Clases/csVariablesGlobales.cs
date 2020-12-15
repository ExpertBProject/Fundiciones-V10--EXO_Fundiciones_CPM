using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.ComponentModel;
using SAPbouiCOM;
using SAPbobsCOM;
using System.Text.RegularExpressions;

namespace Addon_SIA
{   
    public static class csVariablesGlobales
    {
        #region Variables Base
        public static SAPbobsCOM.Company oCompany;
        public static SAPbouiCOM.Application SboApp;
        public static SAPbouiCOM.SboGuiApi oSboGuiApi = new SboGuiApi();
        public static string StrConexion;
        public static string DBPassword;
        public static SqlConnection conAddon;
        public static string StrImpCryV;
        public static string StrImpCryC;
        public static string Prefijo = "SIA";
        public static string StrRutRep;
        public static string StrRutaImagenes = "";
        public static string StrIni = System.Windows.Forms.Application.StartupPath + @"\ConfigSIA.ini";
        public static string StrPath = System.IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString();
        public static string CrearDtosEnDocumentos;
        public static SAPbouiCOM.UserDataSource oUserDataSourceNombreReport;
        public static SAPbouiCOM.UserDataSource oUserDataSourceDescripcionReport;
        public const string Estructura = "[a-z,ñ]";
        public static readonly Regex Estructura_Regex = new Regex(Estructura);
        public static bool RecalculoDePrecio = false;
        public static bool LanzarImpresionCrystal = false;
        public static CrystalDecisions.CrystalReports.Engine.ReportDocument crReport;
        public static int NumeroAsiento = 0;
        public static bool FormularioAsientoGastosAbierto = false;
        public static string StrMsError;
        public static string MenuImprimirPorPantalla = "519";
        public static string MenuImprimirPorImpresora = "520"; //512 previsualizar ,Imprimir"
        public static string MenuExportarAPdf = "7176";
        public static string MenuExportarAExcel = "7169";
        public static string MenuExportarAWord = "7170";
        public static string StrReport;
        public static SAPbouiCOM.EventFilters oFilters;
        public static SAPbouiCOM.EventFilter oFilter;
        public static int PanelActualIC = 0;
        public static string InstanciaFormularioSAP = "";
        public static string ImpresoraPorDefecto = "";
        public static string FormularioEmail = "";
        public static string FormularioEnvioEmail = "";
        public static string NombreArchivoEmail = "";
        public static string DirectorioActual = Environment.CurrentDirectory;
        #endregion

        #region Variables Addon Fundiciones
        public static string NumeroOrden;
        public static string NumeroProduccion;
        public static string FormularioDesde = "";
        public static string FormularioDesdeUID = "";
        public static bool ParaCambiodeReferencia = false;
        public static string AñoModelo = "";
        #endregion
    }
}