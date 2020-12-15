using System;
using System.Collections.Generic;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Collections;
using System.ComponentModel;
using Microsoft.Win32;
using System.Diagnostics;
using System.Threading;

namespace Addon_SIA
{
    public class csUtilidades
    {
        static SAPbobsCOM.UserTable oUserTable = null;
        static SAPbouiCOM.Form oForm;
        static SAPbouiCOM.ComboBox oComboBox;
        static SAPbouiCOM.EditText oEditText;
        static SAPbouiCOM.Matrix oMatrix;
        static SAPbouiCOM.Menus oMenus;
        static SAPbouiCOM.MenuItem oMenuItem;
        static SAPbouiCOM.MenuCreationParams oMenuCreationParams;
        static SAPbobsCOM.Recordset oRecordset;

        public static bool LeerConexion(bool ForzarDesconexion)
        {
            try
            {
                if (csVariablesGlobales.oCompany == null)
                {
                    csVariablesGlobales.oCompany = new SAPbobsCOM.Company();
                    csVariablesGlobales.oCompany.SetSboLoginContext(
                                    csVariablesGlobales.SboApp.Company.GetConnectionContext(
                                    csVariablesGlobales.oCompany.GetContextCookie()));
                }
                if (csVariablesGlobales.oCompany.Connected == true && ForzarDesconexion)
                {
                    csVariablesGlobales.oCompany.Disconnect();
                    csVariablesGlobales.oCompany = null;
                    csVariablesGlobales.oCompany = new SAPbobsCOM.Company();
                    csVariablesGlobales.oCompany.SetSboLoginContext(
                                    csVariablesGlobales.SboApp.Company.GetConnectionContext(
                                    csVariablesGlobales.oCompany.GetContextCookie()));
                }
                
                if (csVariablesGlobales.oCompany.Connected == false)
                {
                    csVariablesGlobales.oCompany.Connect();
                }
                return true;
            }
            catch
            {
                csVariablesGlobales.SboApp.SetStatusBarMessage("No conecta", BoMessageTime.bmt_Short, true);
                return false;
            }
        }

        public static void OperacionesIniciales()
        {
            csVariablesGlobales.SboApp.SetStatusBarMessage("Realizando operaciones iniciales", BoMessageTime.bmt_Short, false);
            System.Windows.Forms.Application.DoEvents();
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            ActuzalizarBaseDatos();
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
        }

        public static bool Inicio()
        {
            csFuncionalidadSAP FuncionalidadSAP = new csFuncionalidadSAP();
            try
            {
                string StrCon = Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
                csVariablesGlobales.oSboGuiApi.Connect(StrCon);
                csVariablesGlobales.SboApp = csVariablesGlobales.oSboGuiApi.GetApplication(-1);
                if (LeerConexion(true))
                {
                    csVariablesGlobales.SboApp.SetStatusBarMessage("Add-On conectado", BoMessageTime.bmt_Short, false);
                    LeerTxtConexion();
                    csVariablesGlobales.StrConexion = "Persist Security Info=False;User ID=" +
                                                  csVariablesGlobales.oCompany.DbUserName +
                                                  ";Password=" + csVariablesGlobales.DBPassword +
                                                  ";Initial Catalog=" + csVariablesGlobales.oCompany.CompanyDB +
                                                  ";Data Source=" + csVariablesGlobales.oCompany.Server;
                    csVariablesGlobales.conAddon = new SqlConnection();
                    csVariablesGlobales.conAddon.ConnectionString = csVariablesGlobales.StrConexion;
                    csVariablesGlobales.conAddon.Open();
                    csVariablesGlobales.SboApp.SetStatusBarMessage("Inicializando Add-On SIA", BoMessageTime.bmt_Medium, false);
                    OperacionesIniciales();
                    csVariablesGlobales.StrRutRep = DameValor("[@" + csVariablesGlobales.Prefijo + "_PARAM]", "U_Ruta", "U_TipoRuta = 'Reports'");
                    csVariablesGlobales.StrRutaImagenes = DameValor("[@" + csVariablesGlobales.Prefijo + "_PARAM]", "U_Ruta", "U_TipoRuta = 'Imágenes'");
                    FuncionalidadSAP.EventFilter();
                    FuncionalidadSAP.SetFilters();
                    return true;
                }
                else
                {
                    csVariablesGlobales.SboApp.SetStatusBarMessage("Add-On no conectado", BoMessageTime.bmt_Short, true);
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error," + ex.Message);
                return false;
            }
        }

        public static bool LeerTxtConexion()
        {
            //ConfigSIA.ini
            string StrLinea;
            int i;
            string strCampo;
            string StrDato;
            int[] IDato = new int[4];
            int[] IDatoL = new int[4];
            int IntChecksum = 0;
            try
            {
                csVariablesGlobales.SboApp.SetStatusBarMessage("Leyendo fichero de configuración", BoMessageTime.bmt_Medium, false);

                if (System.IO.File.Exists(csVariablesGlobales.StrIni) == false)
                {
                    csVariablesGlobales.SboApp.MessageBox("Debes indicar un fichero que exista", 1, "Ok", "", "");
                    csOpenFileDialog OpenFileDialog = new csOpenFileDialog();
                    OpenFileDialog.Filter = "(*.ini)|*.ini";
                    OpenFileDialog.InitialDirectory = csVariablesGlobales.StrPath;// csUtilidades.DameValor("[@SIASL_PARAM]", "U_Ruta", "U_TipoRuta = 'Destino Ficheros SIA'");
                    Thread threadGetFile = new Thread(new ThreadStart(OpenFileDialog.GetFileName));
                    threadGetFile.TrySetApartmentState(ApartmentState.STA);
                    try
                    {
                        threadGetFile.Start();
                        while (!threadGetFile.IsAlive) ; // Wait for thread to get started
                        Thread.Sleep(1);  // Wait a sec more
                        threadGetFile.Join();    // Wait for thread to end

                        // Use file name as you will here
                        csVariablesGlobales.StrIni = OpenFileDialog.FileName;
                        threadGetFile.Abort();
                        threadGetFile = null;
                        OpenFileDialog.InitialDirectory = "";
                        OpenFileDialog = null;
                    }
                    catch (Exception ex)
                    {
                        csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "OK", "", "");
                        threadGetFile.Abort();
                        threadGetFile = null;
                        OpenFileDialog.InitialDirectory = "";
                        OpenFileDialog = null;
                        return false;
                    }
                }

                System.IO.StreamReader FicheroIni = new System.IO.StreamReader(csVariablesGlobales.StrIni);
                while (FicheroIni.Peek() != -1)
                {
                    StrLinea = FicheroIni.ReadLine();

                    i = StrLinea.IndexOf("=", 0);
                    strCampo = "";
                    StrDato = "";
                    if (i != 0)
                    {
                        strCampo = StrLinea.Substring(0, i);
                        StrDato = StrLinea.Substring(i + 1, StrLinea.Length - i - 1);
                    }
                    else
                    {
                        MessageBox.Show("Estructura de fichero Ini incorrecta (=)", "Fichero Erróneo");
                    }
                    if (strCampo == "")
                    {
                        MessageBox.Show("Identificador Nº " + Convert.ToString(i) + " no encontrado", "Fichero Erróneo");
                    }

                    switch (strCampo)
                    {
                        case "DBPassword":
                            csVariablesGlobales.DBPassword = StrDato;
                            IDato[1] = IDato[1] + 1;
                            break;
                        case "CrearDtosEnDocumentos":
                            csVariablesGlobales.CrearDtosEnDocumentos = StrDato;
                            IDato[1] = IDato[1] + 1;
                            break;
                        case "ImpresoraEtiquetas":
                            csVariablesGlobales.ImpresoraPorDefecto = StrDato;
                            IDato[1] = IDato[1] + 1;
                            break;
                    }
                }
                FicheroIni.Close();
                for (i = 1; i <= 3; i++)
                {
                    IntChecksum = IntChecksum + IDato[i];
                }
                if (IntChecksum != 3)
                {
                    csVariablesGlobales.SboApp.SetStatusBarMessage("Fichero erróneo", BoMessageTime.bmt_Medium, false);
                    MessageBox.Show("Faltan datos en la estructura del fichero ini", "Fichero Erróneo");
                    return false;
                }
                return true;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        public static void ActuzalizarBaseDatos()
        {
            csVariablesGlobales.SboApp.SetStatusBarMessage("Comprobando Base de Datos", BoMessageTime.bmt_Medium, false);
            System.Windows.Forms.Application.DoEvents();
            #region Crear Tablas
            csCrearTablas CrearTablas = new csCrearTablas();
            CrearTablas.TablaParametros();
            CrearTablas.TablaReports();
            CrearTablas.TablaFormularios();
            CrearTablas.TablaLineasDocumentosMarketing();
            CrearTablas.TablaInterlocutorComercial();
            CrearTablas.TablaPlanos();
            CrearTablas.TablaUbicaciones();
            CrearTablas.TablaFormularios();
            CrearTablas.TablaModelo347();
            CrearTablas.TablaEstructuraConsultas();
            CrearTablas.TablaDestinoMemoria();
            CrearTablas.TablaBalances();
            CrearTablas.TablaPedidoComprasDevuelta();
            CrearTablas.TablaModelo349();
            #endregion
            #region Crear Consultas
            csCrearConsultas CrearConsultas = new csCrearConsultas();
            #endregion
            #region Crear Vistas
            csCrearVistas CrearVistas = new csCrearVistas();
            CrearVistas.V_SIA_TipoIvaFactVent();
            CrearVistas.V_SIA_TipoIvaFactComp();
            #endregion
            LlenaTablasPropias();
        }

        public static void InsertaRegistroTablaFormularios(string Tabla, string FormSIA, string FormSAP, string Descripcion)
        {
            try
            {
                string StrCodMax = "";
                if (DameValor("[@" + Tabla + "]", "U_FormSIA", "U_FormSIA='" + FormSIA + "'") == "")
                {
                    oUserTable = csVariablesGlobales.oCompany.UserTables.Item(Tabla);
                    StrCodMax = CompletaConCeros(8, UltimoCode("[@" + Tabla + "]"), 1);
                    oUserTable.Code = StrCodMax;
                    oUserTable.Name = StrCodMax;
                    oUserTable.UserFields.Fields.Item("U_FormSIA").Value = FormSIA;
                    oUserTable.UserFields.Fields.Item("U_FormSAP").Value = FormSAP;
                    oUserTable.UserFields.Fields.Item("U_descrip").Value = Descripcion;
                    if (oUserTable.Add() != 0)
                    {
                        csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                    }
                }
            }
            catch
            {
                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
            }
        }

        public static void InsertaRegistroTablaParametros(string Tabla, string Ruta, string TipoRuta, 
                                                   string Concepto, string Valor) //, string CryVen, string CryCom, string Per349
        {
            try
            {
                string StrCodMax = "";
                if (DameValor("[@" + Tabla + "]", "U_TipoRuta", "U_TipoRuta = '" + TipoRuta + "'") == "")
                {
                    oUserTable = csVariablesGlobales.oCompany.UserTables.Item(Tabla);
                    System.Windows.Forms.Application.DoEvents();
                    StrCodMax = CompletaConCeros(8, UltimoCode("[@" + Tabla + "]"), 1);
                    oUserTable.Code = StrCodMax;
                    oUserTable.Name = StrCodMax;
                    oUserTable.UserFields.Fields.Item("U_Ruta").Value = Ruta;
                    oUserTable.UserFields.Fields.Item("U_TipoRuta").Value = TipoRuta;
                    oUserTable.UserFields.Fields.Item("U_Concepto").Value = Concepto;
                    oUserTable.UserFields.Fields.Item("U_Valor").Value = Valor;
                    if (oUserTable.Add() != 0)
                    {
                        csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                    }
                }
            }
            catch
            {
                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
            }
        }

        public static void InsertaRegistroTablaReport(string Tabla, string NombreReport, string Descripcion,
                                               string TipoDocumento, string Borrador)
        {
            try
            {
                string StrCodMax = "";
                if (DameValor("[@" + Tabla + "]", "U_Report", "U_Report = '" + NombreReport +
                              "' AND U_TipDoc = '" + TipoDocumento + "' AND U_Borrador = '" + Borrador + "'") == "")
                {
                    oUserTable = csVariablesGlobales.oCompany.UserTables.Item(Tabla);
                    StrCodMax = CompletaConCeros(8, UltimoCode("[@" + Tabla + "]"), 1);
                    oUserTable.Code = StrCodMax;
                    oUserTable.Name = StrCodMax;
                    oUserTable.UserFields.Fields.Item("U_Report").Value = NombreReport;
                    oUserTable.UserFields.Fields.Item("U_Descrip").Value = Descripcion;
                    oUserTable.UserFields.Fields.Item("U_TipDoc").Value = TipoDocumento;
                    oUserTable.UserFields.Fields.Item("U_Borrador").Value = Borrador;
                    if (oUserTable.Add() != 0)
                    {
                        csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                    }
                }
            }
            catch
            {
                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
            }
        }

        private static void LlenaTablasPropias()
        {
            LeerConexion(true);
            #region Tabla Formularios
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "SIASL00001", "SIASL00001", "SIA - Importaciones SIA");//FrmImportacionesSIA
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "SIASL00002", "SIASL00002", "SIA - Impresión de Listados");//FrmImpresionListados
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "SIASL00003", "SIASL00003", "SIA - Asiento de Gastos");//FrmAsientoGastos
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "SIASL00004", "SIASL00004", "SIA - Seleccionar Report");//FrmImpReports
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "SIASL00005", "SIASL00005", "SIA - Mensaje");//FrmMensaje
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "@SIA_REPORT", "", "SIA - Tabla Reports");//@SIA_REPORT
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "@SIA_FORMS", "", "SIA - Tabla Formularios");//@SIA_FORMS
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "@SIA_PARAM", "", "SIA - Tabla Parámetros");//@SIA_PARAM
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "0", "0", "SAP - Mensajes SAP");//0
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "41", "41", "SAP - Selección de Lote en Formularios SAP");//41
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "133", "133", "SAP - Facturas de Venta");//133
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "134", "134", "SAP - Interlocutores Comerciales");//134
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "139", "139", "SAP - Pedidos de Venta");//139
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "140", "140", "SAP - Entregas");//140
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "141", "141", "SAP - Facturas de Compra");//141
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "142", "142", "SAP - Pedidos de Compra");//142
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "143", "143", "SAP - Entrada Mercancías");//143
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "149", "149", "SAP - Oferta");//149
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "179", "179", "SAP - Abono Deudores");//179
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "180", "180", "SAP - Devolucion de Ventas");//180
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "181", "181", "SAP - Abono Proveedores");//181
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "182", "182", "SAP - Devolucion de Compras");//182
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "392", "392", "SAP - Asientos");//392
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "606", "606", "SAP - Depósito");//606
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "720", "720", "SAP - Salida Mercancías");//720
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "721", "721", "SAP - Entrada Mercancías");//721
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "3002", "3002", "SAP - Documentos Preliminares");//3002
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "60052", "60052", "SAP - Operaciones de Efectos");//60052
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "60090", "60090", "SAP - Factura Clientes + Pago");//60090
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "60091", "60091", "SAP - Factura Anticipo Clientes");//60091
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "60092", "60092", "SAP - Factura Anticipo Proveedores");//60092
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "65300", "65300", "SAP - Factura Anticipo Deudores");//65300
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "65301", "65301", "SAP - Factura Anticipo Proveedores");//65301
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "65308", "65308", "SAP - Solicitud Anticipo Deudores");//65308
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "65309", "65309", "SAP - Solicitud Anticipo Proveedores");//65309
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "65211", "65211", "SAP - Orden de Fabricación");//65211
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "10003", "10003", "SAP - Selección Artículos");//10001
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "65010", "65010", "SAP - Informe 347");//65010
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "60051", "60051", "SAP - Gestión de Efectos");//60051
            InsertaRegistroTablaFormularios(csVariablesGlobales.Prefijo + "_FORMS", "65011", "65011", "SAP - Criterios de Selccion 347");//Criterios de Selccion 347
            #endregion
            #region Tabla Parámetros
            InsertaRegistroTablaParametros(csVariablesGlobales.Prefijo + "_PARAM", @"Y:\Report\", "Reports", "Ubicación Impresos", "");//Ubicación Impresos  , "Y", "Y", "T",
            InsertaRegistroTablaParametros(csVariablesGlobales.Prefijo + "_PARAM", @"Y:\Destino Ficheros Modelos\", "Destino Ficheros Modelos", "Ubicación Destino Ficheros Modelos", "");//Ubicación Destino Ficheros Modelos  , "N", "N", ""
            InsertaRegistroTablaParametros(csVariablesGlobales.Prefijo + "_PARAM", @"Y:\Destino Ficheros Varios\", "Destino Ficheros Varios", "Ubicación Destino Ficheros Varios", "");//Ubicación Destino Ficheros Varios  , "N", "N", ""
            InsertaRegistroTablaParametros(csVariablesGlobales.Prefijo + "_PARAM", @"Y:\Destino Ficheros SIA\", "Destino Ficheros SIA", "Ubicación Destino Ficheros SIA", "");//Ubicación Destino Ficheros SIA  , "N", "N", ""
            InsertaRegistroTablaParametros(csVariablesGlobales.Prefijo + "_PARAM", @"", @"Aplicar Dtos. En: Precio/Descuento", "Aplicar Dtos. En: Precio/Descuento", "");//Aplicar Dtos. En:  , "N", "N", ""
            InsertaRegistroTablaParametros(csVariablesGlobales.Prefijo + "_PARAM", @"Y:\Imagenes\", "Imágenes", "Ubicación Imágenes", "");//Imágenes  , "N", "N", ""
            InsertaRegistroTablaParametros(csVariablesGlobales.Prefijo + "_PARAM", @"C:\Destino Ficheros Modelos\", "Destino Ficheros Modelos En Local", "Ubicación Destino Ficheros Modelos En Local", "");//Ubicación Destino Ficheros Modelos  , "N", "N", ""
            InsertaRegistroTablaParametros(csVariablesGlobales.Prefijo + "_PARAM", @"C:\Destino Ficheros Varios\", "Destino Ficheros Varios En Local", "Ubicación Destino Ficheros Varios En Local", "");//Ubicación Destino Ficheros Varios  , "N", "N", ""
            InsertaRegistroTablaParametros(csVariablesGlobales.Prefijo + "_PARAM", @"C:\TemporalesSIA\", "Destino Ficheros Temporales En Local", "Ubicación Destino Ficheros Varios En Local", "");//Ubicación Destino Ficheros Temporales  , "N", "N", ""
            #endregion
        }

        public static bool BuscarTablaReport(string TypeEx, string strBorr, string StrReport, ref string StrTabla,
                                      ref string StrDirReport)
        {
            if (strBorr == "Y")
            {
                StrTabla = "ODRF";
                StrDirReport = csVariablesGlobales.StrRutRep + StrReport;
                return true;
            }
            switch (TypeEx)
            {
                case "133":
                case "60091":
                    StrTabla = "OINV";
                    break;
                case "139":
                    StrTabla = "ORDR";
                    break;
                case "140":
                    StrTabla = "ODLN";
                    break;
                case "141":
                case "60092":
                    StrTabla = "OPCH";
                    break;
                case "142":
                    StrTabla = "OPOR";
                    break;
                case "143":
                    StrTabla = "OPDN";
                    break;
                case "149":
                    StrTabla = "OQUT";
                    break;
                case "179":
                    StrTabla = "ORIN";
                    break;
                case "180":
                    StrTabla = "ORDN";
                    break;
                case "181":
                    StrTabla = "ORPC";
                    break;
                case "182":
                    StrTabla = "ORPD";
                    break;
                case "392":
                    StrTabla = "OJDT";
                    break;
                case "606":
                    StrTabla = "ODPS";
                    break;
                case "3002":
                    StrTabla = "ODRF";
                    break;
                case "60052":
                    StrTabla = "OBOT";
                    break;
                case "60090":
                    StrTabla = "OINV";
                    break;
                case "65300":
                case "65308": //factura de anticipo
                    StrTabla = "ODPI";
                    break;
                case "65309":
                case "65301":
                    StrTabla = "ODPO";
                    break;
                case "SIASL10004":
                    StrTabla = "";
                    break;
                default:
                    return false;
            }
            StrDirReport = csVariablesGlobales.StrRutRep + StrReport;
            return true;
        }

        public static bool IsNumeric(object Expression)
        {
            bool isNum;
            double retNum;

            isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }

        public static string ComPunN(string StrValor)
        {
            if (!IsNumeric(StrValor))
            {
                return "0";
            }
            StrValor.Replace(",", ".");
            if (StrValor == "")
            {
                return "0";
            }
            return StrValor.Replace(",", ".");
        }

        public static bool ExisteForm(string TypeEx, SAPbouiCOM.Application SboApp, ref  string StrTabla,
                               ref string StrReport, ref string StrDocEntry, string Menu)
        {
            csImpresiones Impresiones = new csImpresiones();
            string StrRep;
            string StrBorr;
            int Num;
            SAPbouiCOM.Form oFormMenu;

            oForm = SboApp.Forms.ActiveForm;
            StrBorr = "N";
            switch (TypeEx)
            {
                case "3002":
                    StrBorr = "Y";
                    break;
                case "60052":
                case "188":
                    StrBorr = "N";
                    SAPbouiCOM.Form oForm2;
                    oForm2 = csVariablesGlobales.SboApp.Forms.Item(csVariablesGlobales.FormularioEmail);
                    TypeEx = oForm2.TypeEx.ToString();
                    break;
                default:
                    int Longitud = 5;
                    if (TypeEx.Length < 5)
                    {
                        Longitud = Convert.ToInt32(TypeEx.Length);
                    }
                    if (TypeEx.Substring(0, Longitud) != csVariablesGlobales.Prefijo.Substring(0, Longitud))
                    {
                        oComboBox = (SAPbouiCOM.ComboBox)(oForm.Items.Item("81").Specific);
                        if (oComboBox.Selected.Value == "6")
                        {
                            StrBorr = "Y";
                        }
                        else
                        {
                            StrBorr = "N";
                        }
                    }
                    break;
            }
            StrRep = "";
            Num = Convert.ToInt32(ComPunN(DameValor("[@" + csVariablesGlobales.Prefijo + "_REPORT]", "count(U_REPORT)", ("U_TipDoc='" + TypeEx.ToString() + "' AND U_Borrador='" + StrBorr.ToString() + "'"))));
            if (Num > 1)
            {
                switch (TypeEx)
                {
                    case "133":
                    case "139":
                    case "140":
                    case "142":
                    case "149":
                    case "179":
                    case "180":
                    case "3002":
                    case "60052":
                    case "60090":
                    case "60091":
                    case "65300":
                        //LoadFromXML(@"Formularios\ImpReports.xml", SboApp);
                        LoadFromXML(System.Windows.Forms.Application.StartupPath + @"\Formularios\ImpReports.xml", 
                                    csVariablesGlobales.SboApp);
                        break;
                }
                oFormMenu = SboApp.Forms.Item("FrmImpReports");
                csVariablesGlobales.oUserDataSourceNombreReport = oFormMenu.DataSources.UserDataSources.Item("dsNomImp");
                csVariablesGlobales.oUserDataSourceDescripcionReport = oFormMenu.DataSources.UserDataSources.Item("dsDesImp");
                Impresiones.CargarReports(oFormMenu.UniqueID.ToString(), TypeEx.ToString(), StrBorr.ToString(), StrDocEntry, Menu, IdiomaDocumento(StrDocEntry, TypeEx));
                StrDocEntry = "0";
                return true;
            }
            else
            {
                StrRep = DameValor("[@" + csVariablesGlobales.Prefijo + "_REPORT]", "U_Report", "U_TipDoc='" + TypeEx + "' AND U_Borrador='" + StrBorr + "'");
            }
            if (StrRep == "")
            {
                return false;
            }
            return BuscarTablaReport(TypeEx, StrBorr, StrRep, ref StrTabla, ref StrReport);
        }

        public static string IdiomaDocumento(string StrDocEntry, string TypeEx)
        {
            string StrSql = "ES";
            try
            {
                switch (TypeEx)
                {
                    case "133":
                    case "60090":
                    case "60091":
                        StrSql = "SELECT ShortName FROM OINV, OLNG WHERE LangCode = Code AND DocEntry = " + StrDocEntry;
                        break;
                    case "139":
                        StrSql = "SELECT ShortName FROM ORDR, OLNG WHERE LangCode = Code AND DocEntry = " + StrDocEntry;
                        break;
                    case "140":
                        StrSql = "SELECT ShortName FROM ORDR, OLNG WHERE LangCode = Code AND DocEntry = " + StrDocEntry;
                        break;
                    case "141":
                    case "60092":
                        StrSql = "SELECT ShortName FROM OPCH, OLNG WHERE LangCode = Code AND DocEntry = " + StrDocEntry;
                        break;
                    case "142":
                        StrSql = "SELECT ShortName FROM OPOR, OLNG WHERE LangCode = Code AND DocEntry = " + StrDocEntry;
                        break;
                    case "143":
                        StrSql = "SELECT ShortName FROM OPDN, OLNG WHERE LangCode = Code AND DocEntry = " + StrDocEntry;
                        break;
                    case "149":
                        StrSql = "SELECT ShortName FROM OQUT, OLNG WHERE LangCode = Code AND DocEntry = " + StrDocEntry;
                        break;
                    case "179":
                        StrSql = "SELECT ShortName FROM ORIN, OLNG WHERE LangCode = Code AND DocEntry = " + StrDocEntry;
                        break;
                    case "180":
                        StrSql = "SELECT ShortName FROM ORDN, OLNG WHERE LangCode = Code AND DocEntry = " + StrDocEntry;
                        break;
                    case "3002":
                        StrSql = "SELECT ShortName FROM ORDR, OLNG WHERE LangCode = Code AND DocEntry = " + StrDocEntry;
                        break;
                }
                System.Data.DataTable RstDato = new System.Data.DataTable();
                SqlCommand SelectCMD = new SqlCommand(StrSql, csVariablesGlobales.conAddon);
                SqlDataAdapter SQlDAGrid = new SqlDataAdapter();
                SQlDAGrid.SelectCommand = SelectCMD;
                SQlDAGrid.Fill(RstDato);
                return RstDato.Rows[0][0].ToString();
            }
            catch
            {
                return StrSql;
            }
        }

        public static void LoadFromXML(string FicheroXml, SAPbouiCOM.Application SboApp)
        {
            //throw new Exception("The method or operation is not implemented.");
            System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();
            // load the content of the XML File
            oXmlDoc.Load(FicheroXml);
            // load the form to the SBO application in one batch
            string Xml = oXmlDoc.InnerXml.ToString();
            csVariablesGlobales.SboApp.LoadBatchActions(ref Xml);
        }

        public static string docEntry(SAPbouiCOM.Application SboApp)
        {
            string Tabla;
            try
            {
                oForm = SboApp.Forms.ActiveForm;
                switch (oForm.TypeEx)
                {
                    case "3002":                
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                        for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                        {
                            if (oMatrix.IsRowSelected(i))
                            {
                                oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific;
                                return oEditText.Value;
                            }
                        }
                        return "";
                    case "188":
                        SAPbouiCOM.Form oForm2;
                        oForm2 = csVariablesGlobales.SboApp.Forms.Item(csVariablesGlobales.FormularioEmail);
                        oEditText = (SAPbouiCOM.EditText)(oForm2.Items.Item("8").Specific);
                        Tabla = oEditText.DataBind.TableName;
                        return oForm2.DataSources.DBDataSources.Item(Tabla).GetValue("DocEntry", 0).ToString();
                    default:
                        oEditText = (SAPbouiCOM.EditText)(oForm.Items.Item("8").Specific);
                        Tabla = oEditText.DataBind.TableName;
                        return oForm.DataSources.DBDataSources.Item(Tabla).GetValue("DocEntry", 0).ToString();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                return "";
            }
        }

        public static string absEntry(SAPbouiCOM.Application SBOApp)
        {
            string Tabla;
            try
            {
                oForm = SBOApp.Forms.ActiveForm;
                switch (oForm.TypeEx)
                {
                    case "60052":
                        oEditText = (SAPbouiCOM.EditText)(oForm.Items.Item("5").Specific);
                        Tabla = oEditText.DataBind.TableName;
                        return oForm.DataSources.DBDataSources.Item(Tabla).GetValue("absEntry", 0).ToString();
                    case "188":
                        oMatrix = (SAPbouiCOM.Matrix)(oForm.Items.Item("22").Specific);
                        oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("2").Cells.Item(1).Specific);
                        Tabla = oEditText.DataBind.TableName;
                        return oForm.DataSources.DBDataSources.Item(Tabla).GetValue("absEntry", 0).ToString();
                    default:
                        oEditText = (SAPbouiCOM.EditText)(oForm.Items.Item("6").Specific);
                        Tabla = oEditText.DataBind.TableName;
                        return oForm.DataSources.DBDataSources.Item(Tabla).GetValue("absEntry", 0).ToString();
                }
            }
            catch
            {
                return "";
            }
        }

        public static void ActualizarUnCampo(string Tabla, string Campo, string Condicion, string Valor)
        {
            try
            {
                string CadenaSelect;
                if (Condicion != "")
                {
                    CadenaSelect = "UPDATE " + Tabla + " SET " + Campo + " = '" + Valor + "' WHERE " + Condicion;
                }
                else
                {
                    CadenaSelect = "UPDATE " + Tabla + " SET " + Campo + " = '" + Valor + "'";
                }
                oRecordset = ((SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));
                oRecordset.DoQuery(CadenaSelect);
            }
            catch
            {
                return;
            }
        }

        public static string DameValor(string StrTabla, string StrCampo, string StrCondicion)
        {
            try
            {
                string CadenaSelect;
                if (StrCondicion == "")
                {
                    CadenaSelect = "SELECT " + StrCampo + " FROM " + StrTabla;
                }
                else
                {
                    CadenaSelect = "SELECT " + StrCampo + " FROM " + StrTabla + " WHERE " + StrCondicion;
                }

                oRecordset = ((SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));
                oRecordset.DoQuery(CadenaSelect);
                oRecordset.MoveFirst();
                if (oRecordset.RecordCount <= 0)
                {
                    return "";
                }
                else
                {
                    if (oRecordset.Fields.Item(0).Value.ToString() != "")
                    {
                        string a = oRecordset.Fields.Item(0).Value.ToString();
                        return oRecordset.Fields.Item(0).Value.ToString();
                    }
                    else
                    {
                        return "";
                    }
                }
            }
            catch
            {
                return "";
            }
        }

        public static string GetVersionDll()
        {
            RegistryKey regVersion;
            string KeyValue;
            string StrValor;
            StrValor = "";
            KeyValue = @"software\\SAP\\SIA";
            try
            {
                regVersion = Registry.LocalMachine.OpenSubKey(KeyValue, false);
                if ((regVersion != null))
                {
                    StrValor = regVersion.GetValue("VersionDll", "").ToString();
                    regVersion.Close();
                }
            }
            catch
            {
                return "";
            }
            return StrValor;
        }

        public static bool SetVersionDll(string StrValor)
        {
            RegistryKey regVersion;
            string KeyValue;
            try
            {
                KeyValue = @"software\\SAP\\SIA";
                regVersion = Registry.LocalMachine.OpenSubKey(KeyValue, true);
                if (regVersion == null)
                {
                    regVersion = Registry.LocalMachine.CreateSubKey(KeyValue);
                }
                if (regVersion != null)
                {
                    regVersion.SetValue("VersionDll", StrValor);
                    regVersion.Close();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void CargarRecordSet(ref System.Data.DataTable RstDatos, string StrConsulta)
        {
            SqlCommand sqlCMD = new SqlCommand(StrConsulta, csVariablesGlobales.conAddon);
            SqlDataAdapter oDa = new SqlDataAdapter();
            oDa.SelectCommand = sqlCMD;
            RstDatos = new System.Data.DataTable("datos");
            oDa.Fill(RstDatos);
        }

        public static string CompletaConCeros(int NumeroCaracteres, string Valor, int AumentarContadorEn)
        {
            if (Valor == "")
            {
                Valor = "0";
            }
            string Numero = Convert.ToString(Convert.ToInt32(Valor) + AumentarContadorEn);
            for (int i = Numero.Length; i < NumeroCaracteres; i++)
            {
                Numero = "0" + Numero;
            }
            return Numero;
        }

        public static string UltimoCode(string StrTabla)
        {
            oRecordset = ((SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));
            oRecordset.DoQuery("SELECT Max(Code) As Code FROM " + StrTabla);
            oRecordset.MoveFirst();
            if (oRecordset.Fields.Item("Code").Value.ToString() != "")
            {
                return oRecordset.Fields.Item("Code").Value.ToString();
            }
            else
            {
                return "";
            }
        }

        public static void SBImprimirReportDoc(string FormUID, ref SAPbouiCOM.ItemEvent pVal, string AbsDocEntry)
        {
            csImpresiones Impresiones = new csImpresiones();
            string StrMenu;
            string DocEntry;
            string TypeEx;
            string Borrador;
            string StrRep;
            string StrTabla;
            string StrDirReport;
            bool BolResult;
            StrRep = "";
            StrTabla = "";
            StrDirReport = "";
            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
            oEditText = (SAPbouiCOM.EditText)(oForm.Items.Item("txtDoc").Specific);
            DocEntry = oEditText.String;
            oEditText = (SAPbouiCOM.EditText)(oForm.Items.Item("txtForm").Specific);
            TypeEx = oEditText.String;
            oEditText = (SAPbouiCOM.EditText)(oForm.Items.Item("txtBorr").Specific);
            Borrador = oEditText.String;
            oEditText = (SAPbouiCOM.EditText)(oForm.Items.Item("txtMenu").Specific);
            StrMenu = oEditText.String;
            oMatrix = (SAPbouiCOM.Matrix)(oForm.Items.Item("matRep").Specific);
            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("colNomImp").Cells.Item(pVal.Row).Specific);
            StrRep = oEditText.String;
            BolResult = BuscarTablaReport(TypeEx, Borrador, StrRep, ref StrTabla, ref StrDirReport);
            if (BolResult == false | StrTabla == "" | StrDirReport == "")
            {
                csVariablesGlobales.SboApp.MessageBox("No se han encontrado datos para imprimir", 1, "Ok", "", "");
                return;
            }
            BolResult = Impresiones.ImprimirDocumento(ref csVariablesGlobales.SboApp, StrMenu, StrTabla, DocEntry, TypeEx, StrDirReport, AbsDocEntry, 1, StrRep);
            if (BolResult == true)
            {
                oForm.Close();
            }
        }

        public static void ShutDown()
        {
            try
            {
                System.Windows.Forms.Application.Exit();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error, " + ex.Message);
            }
        }

        public static bool CargarMenuXML()
        {
            try
            {
                try
                {
                    oMenus = csVariablesGlobales.SboApp.Menus;
                    oMenus.RemoveEx("SubMnuUtil");
                }
                catch { }

                oMenuCreationParams = ((SAPbouiCOM.MenuCreationParams)(csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                oMenuItem = csVariablesGlobales.SboApp.Menus.Item("43520"); //moudles'

                string sPath;

                sPath = System.Windows.Forms.Application.StartupPath;

                oMenuItem = csVariablesGlobales.SboApp.Menus.Item("43520");

                oMenus = oMenuItem.SubMenus;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMenus);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMenuItem);

                LoadFromXML(sPath + @"\XML\XMLMenu.xml", csVariablesGlobales.SboApp);

                return true;

            }
            catch (Exception er)
            { // Menu ya existente
                csVariablesGlobales.SboApp.MessageBox("Error al crear el menu " + er.Message, 0, "Ok", "", "");
                return false;
            }
        }

        public static void CrearMenu(string MenuPadre, SAPbouiCOM.BoMenuType oBoMenuType, string UniqueId,
                              string Descripcion, string Imagen, bool Checked, bool Enabled)
        {
            oMenuItem = csVariablesGlobales.SboApp.Menus.Item(MenuPadre);
            //oMenuCreationParams.Type = oBoMenuType;
            oMenuCreationParams.UniqueID = UniqueId;
            oMenuCreationParams.String = Descripcion;
            oMenuCreationParams.Image = Imagen;
            oMenuCreationParams.Checked = Checked;
            oMenuCreationParams.Enabled = Enabled;
            oMenuCreationParams.Position = oMenuItem.SubMenus.Count + 1;
            switch (oBoMenuType)
            {
                case BoMenuType.mt_POPUP:
                    oMenuCreationParams.Type = oBoMenuType;
                    oMenus = oMenuItem.SubMenus;                    
                    break;
                case BoMenuType.mt_STRING:
                    oMenus = oMenuItem.SubMenus;
                    oMenuCreationParams.Type = oBoMenuType;
                    break;
            }
            try
            {
                oMenus.AddEx(oMenuCreationParams);
            }
            catch (Exception er)
            {
                csVariablesGlobales.SboApp.MessageBox("Error al crear el menu " + er.Message, 0, "Ok", "", "");
            }
            //switch (oBoMenuType)
            //{
            //    case BoMenuType.mt_POPUP:
            //        oMenuItem = csVariablesGlobales.SboApp.Menus.Item(MenuPadre);
            //        oMenuCreationParams.Type = oBoMenuType;
            //        oMenuCreationParams.UniqueID = UniqueId;
            //        oMenuCreationParams.String = Descripcion;
            //        oMenuCreationParams.Image = Imagen;
            //        oMenuCreationParams.Checked = Checked;
            //        oMenuCreationParams.Enabled = Enabled;
            //        oMenuCreationParams.Position = oMenuItem.SubMenus.Count + 1;

            //        oMenus = oMenuItem.SubMenus;
            //        try
            //        {
            //            oMenus.AddEx(oMenuCreationParams);
            //        }
            //        catch (Exception er)
            //        {
            //            csVariablesGlobales.SboApp.MessageBox("Error al crear el menu " + er.Message, 0, "Ok", "", "");
            //        }
            //        break;
            //    case BoMenuType.mt_STRING:
            //        oMenuItem = csVariablesGlobales.SboApp.Menus.Item(MenuPadre);
            //        oMenus = oMenuItem.SubMenus;
            //        oMenuCreationParams.Type = oBoMenuType;
            //        oMenuCreationParams.UniqueID = UniqueId;
            //        oMenuCreationParams.String = Descripcion;
            //        oMenuCreationParams.Image = Imagen;
            //        oMenuCreationParams.Checked = Checked;
            //        oMenuCreationParams.Enabled = Enabled;
            //        oMenuCreationParams.Position = oMenuItem.SubMenus.Count + 1;
            //        oMenus.AddEx(oMenuCreationParams);
            //        break;
            //}
        }

        public static void CargarMenu()
        {
            try
            {
                try
                {
                    oMenus = csVariablesGlobales.SboApp.Menus;
                    oMenus.RemoveEx("SubMnuUtil");
                }
                catch
                {
                }
                try
                {
                    oMenuCreationParams = ((SAPbouiCOM.MenuCreationParams)
                                           (csVariablesGlobales.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                    CrearMenu("43520", BoMenuType.mt_POPUP, "SubMnuUtil", "Utilidades SIA", csVariablesGlobales.StrRutaImagenes +
                          @"\SIA.jpg", true, true);
                    CrearMenu("SubMnuUtil", BoMenuType.mt_POPUP, "SubMnuManten", "Mantenimientos", "",
                          true, true);
                    CrearMenu("SubMnuManten", BoMenuType.mt_STRING, "SubMnuEntSalStock", "Entradas Salidas Stock",
                              "", true, true);
                    CrearMenu("SubMnuManten", BoMenuType.mt_STRING, "SubMnuProcesos", "Procesos",
                              "", true, true);
                    CrearMenu("SubMnuManten", BoMenuType.mt_STRING, "SubMnuImpLis", "Impresion de Listados",
                              "", true, true);
                    CrearMenu("SubMnuManten", BoMenuType.mt_STRING, "SubMnuGenMem", "Generar Memoria",
                              "", true, true);
//                    CrearMenu("SubMnuManten", BoMenuType.mt_STRING, "SubMnuGenFicCesce", "Generar Fichero CESCE",
  //                            "", true, true);
                    CrearMenu("SubMnuUtil", BoMenuType.mt_POPUP, "SubMnuModelos", "Modelos", "",
                          true, true);
                    //CrearMenu("SubMnuModelos", BoMenuType.mt_STRING, "SubModelo347", "Modelo 347",
                      //        "", true, true);
                    CrearMenu("SubMnuModelos", BoMenuType.mt_STRING, "SubModelo349", "Modelo 349",
                              "", true, true);
                    CrearMenu("SubMnuUtil", BoMenuType.mt_POPUP, "MnuUtilAdministrador", "Util. Administrador", "",
                          true, true);
                    CrearMenu("MnuUtilAdministrador", BoMenuType.mt_STRING, "SubMnuImpSIA", "Importaciones SIA",
                              "", true, true);
                }
                catch
                {
                    csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                }
            }
            catch (Exception er)
            {
                csVariablesGlobales.SboApp.MessageBox("Error al crear el menu " + er.Message, 0, "Ok", "", "");
            }
        }

        public static void SaveAsXml(SAPbouiCOM.Form oForm, string Archivo, string RutaAlternativa)
        {
            System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();
            oXmlDoc.LoadXml(oForm.GetAsXML());
            if (RutaAlternativa != "")
            {
                oXmlDoc.Save((csVariablesGlobales.StrPath + @"\Debug\" + RutaAlternativa + @"\" + Archivo));
            }
            else
            {
                oXmlDoc.Save((csVariablesGlobales.StrPath + @"\Debug\Formularios\" + Archivo));
            }
        }

        public static bool TodasMayusculas(string Cadena)
        {
            if (csVariablesGlobales.Estructura_Regex.IsMatch(Cadena))
            {
                return false;
            }
            return true;
        }

        public static string PunCom(object StrValor)
        {
            if (!IsNumeric(StrValor))
            {
                return "0";
            }
            StrValor.ToString().Replace(".", ",");
            if (StrValor.ToString() == "")
            {
                return "0";
            }
            return StrValor.ToString().Replace(".", ",");
        }

        public static string ComPun(object StrValor)
        {
            if (!IsNumeric(StrValor))
            {
                return "0";
            }
            StrValor.ToString().Replace(",", ".");
            if (StrValor.ToString() == "")
            {
                return "0";
            }
            return StrValor.ToString().Replace(",", ".");
        }

        public static double TextoADouble(string StrTexto)
        {
            string StrAux;
            if (!IsNumeric(StrTexto))
            {
                return 0;
            }
            StrAux = StrTexto.Replace(".", ",");
            double a = Convert.ToDouble(StrAux);
            return a; // Convert.ToDouble(StrAux);
        }

        public static double MonedaADouble(string StrTexto)
        {
            string StrAux;
            double a;
            if (StrTexto == "")
            {
                StrTexto = "0";
            }
            StrAux = StrTexto;
            if (IsNumeric(StrAux))
            {
                if (StrAux == "0")
                {
                    StrAux = StrAux.Replace(".", "");
                    return Convert.ToDouble(StrAux);
                }
            }
            StrAux = StrAux.Trim();
            StrAux = StrAux.Replace("EUR", "");
            string b = StrAux.IndexOf(" ", 0, StringComparison.CurrentCulture).ToString();
            if (StrAux.IndexOf(" ", 0, StringComparison.CurrentCulture) != -1)
            {
                StrAux = StrAux.Substring(0, StrAux.IndexOf(" ", 0, StringComparison.CurrentCulture));
            }
            a = Convert.ToDouble(StrAux);
            if (!IsNumeric(a))
            {
                return 0;
            }
            StrAux = StrAux.Replace(".", "");
            return Convert.ToDouble(StrAux);
        }

        public static void CargarDescuentos(SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.ItemEvent pVal, out double Dto1,
                                     out double Dto2, out double Dto3, out double Dto4, out double Dto5,
                                     out double dblPrecio)
        {
            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_PRET").Cells.Item(pVal.Row).Specific);
            dblPrecio = MonedaADouble(oEditText.String);
            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_DTO1").Cells.Item(pVal.Row).Specific);
            Dto1 = TextoADouble(oEditText.String);
            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_DTO2").Cells.Item(pVal.Row).Specific);
            Dto2 = TextoADouble(oEditText.String);
            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_DTO3").Cells.Item(pVal.Row).Specific);
            Dto3 = TextoADouble(oEditText.String);
            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_DTO4").Cells.Item(pVal.Row).Specific);
            Dto4 = TextoADouble(oEditText.String);
            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_DTO5").Cells.Item(pVal.Row).Specific);
            Dto5 = TextoADouble(oEditText.String);
        }

        public static void CalcularDescuentosVentas(double dPreComIni, double dDto1, double dDto2, double dDto3,
                                             double dDto4, double dDto5, SAPbouiCOM.Matrix oMatrix,
                                             SAPbouiCOM.ItemEvent pVal, string AsignarAPrecioODescuento)
        {
            SAPbouiCOM.EditText oEditText;
            double dblPreciofinal;
            double Descuento = 0;
            int IntDec;
            IntDec = Convert.ToInt32(DameValor("OADM", "PriceDec", ""));
            bool a = csVariablesGlobales.RecalculoDePrecio;
            IntDec = 4;
            dblPreciofinal = dPreComIni;
            if (dDto1 != 100 && dDto2 != 100 && dDto3 != 100 && dDto4 != 100 && dDto5 != 100)
            {
                if (dDto1 != 0)
                {
                    dblPreciofinal = Math.Round(dblPreciofinal - (dblPreciofinal * dDto1 / 100), IntDec);
                    Descuento = 1 - (dDto1 / 100);
                }
                if (dDto2 != 0)
                {
                    dblPreciofinal = Math.Round(dblPreciofinal - (dblPreciofinal * dDto2 / 100), IntDec);
                    Descuento *= (1 - (dDto2 / 100));
                }
                if (dDto3 != 0)
                {
                    dblPreciofinal = Math.Round(dblPreciofinal - (dblPreciofinal * dDto3 / 100), IntDec);
                    Descuento *= (1 - (dDto3 / 100));
                }
                if (dDto4 != 0)
                {
                    dblPreciofinal = Math.Round(dblPreciofinal - (dblPreciofinal * dDto4 / 100), IntDec);
                    Descuento *= (1 - (dDto4 / 100));
                }
                if (dDto5 != 0)
                {
                    dblPreciofinal = Math.Round(dblPreciofinal - (dblPreciofinal * dDto5 / 100), IntDec);
                    Descuento *= (1 - (dDto5 / 100));
                }
            }
            else
            {
                Descuento = 1;
                dblPreciofinal = 0;
            }
            oEditText = null;
            try
            {
                //Columna unidades base
                switch (AsignarAPrecioODescuento)
                {
                    case "Descuento":
                        oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific);
                        oEditText.String = dPreComIni.ToString();
                        oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific);
                        oEditText.String = "100";
                        if (Descuento != 1)
                        {
                            oEditText.String = ((1 - Descuento) * 100).ToString();
                        }
                        if (Descuento == 0)
                        {
                            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific);
                            oEditText.String = "0";
                        }
                        break;
                    case "Precio":
                        oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific);
                        oEditText.String = "0";
                        oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific);
                        oEditText.String = dblPreciofinal.ToString();
                        break;
                }
                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public static double ConvertirCantidad(string Valor)
        {
            string SeparadorDecimal = DameValor("OADM", "DecSep", "");
            string SeparadorMiles = DameValor("OADM", "ThousSep", "");

            Valor = Valor.Replace(SeparadorMiles, "");
                
            if (SeparadorDecimal != ",")
            {
                return Convert.ToDouble(Valor.Replace(SeparadorDecimal, ","));
            }
            return Convert.ToDouble(Valor);
        }
        
        public static string LeerCadenaSeparadaPorCaracter(string Cadena, char Caracter, int PosicionCadena)
        {
            char[] Caracteres = { Caracter };
            string[] CadenaSeparada = Cadena.Split(Caracteres);
            return CadenaSeparada[PosicionCadena - 1];
        }

        public static double Truncar(double Valor, int Decimales)
        {
            if (Valor > 0)
                Valor = Math.Floor(Valor * Math.Pow(10, Decimales)) / Math.Pow(10, Decimales);
            else
                Valor = Math.Ceiling(Valor * Math.Pow(10, Decimales)) / Math.Pow(10, Decimales);
            return Valor;
        }

        public static void CancelarApunte(string TipoMovimiento, string Clave)
        {
            string Transaccion = "";
            SAPbobsCOM.JournalEntries oJournalEntries;
            SAPbobsCOM.JournalEntries oJournalEntries2;

            oJournalEntries = (SAPbobsCOM.JournalEntries)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);
            oJournalEntries2 = (SAPbobsCOM.JournalEntries)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);

            if (TipoMovimiento == "Salida")
            {
                Transaccion = DameValor("OIGE", "TransId", "DocEntry='" + Clave + "'");
            }
            else if (TipoMovimiento == "Entrada")
            {
                Transaccion = DameValor("OIGN", "TransId", "DocEntry='" + Clave + "'");
            }

            if (Transaccion != "")
            {
                oJournalEntries.GetByKey(Convert.ToInt32(Transaccion));
                oJournalEntries2.DueDate = oJournalEntries.DueDate;
                oJournalEntries2.TaxDate = oJournalEntries.TaxDate;
                oJournalEntries2.ReferenceDate = oJournalEntries.ReferenceDate;
                oJournalEntries2.Memo = "";
                for (int i = 0; i < oJournalEntries.Lines.Count; i++)
                {
                    if (i > 0)
                    {
                        oJournalEntries2.Lines.Add();
                    }
                    oJournalEntries.SetCurrentLine(i);
                    oJournalEntries2.Lines.AccountCode = oJournalEntries.Lines.AccountCode;
                    oJournalEntries2.Lines.ShortName = oJournalEntries.Lines.ShortName;
                    oJournalEntries2.Lines.Debit = -1 * oJournalEntries.Lines.Debit;
                    oJournalEntries2.Lines.Credit = -1 * oJournalEntries.Lines.Credit;
                }
                if (oJournalEntries2.Add() != 0)
                {
                    csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                }
            }
        }

        public static void ActualizarDocumentoAEnviar()
        {
            if (csVariablesGlobales.NombreArchivoEmail != "" && csVariablesGlobales.FormularioEnvioEmail != "")
            {
                SAPbouiCOM.Form oForm2;
                oForm2 = csVariablesGlobales.SboApp.Forms.Item(csVariablesGlobales.FormularioEnvioEmail);
                oMatrix = (SAPbouiCOM.Matrix)oForm2.Items.Item("22").Specific;
                oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("4").Cells.Item(1).Specific;
                string Indice = oEditText.String;

                SAPbobsCOM.Attachments2 oAttachments2 = null;
                SAPbobsCOM.Attachments2_Lines oAttachments2_Lines = null;

                Indice = absEntry(csVariablesGlobales.SboApp);
                oAttachments2 = (SAPbobsCOM.Attachments2)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2));
                oAttachments2.GetByKey(System.Convert.ToInt32(Indice));
                oAttachments2_Lines = oAttachments2.Lines;
                oAttachments2_Lines.SetCurrentLine(0);
                string RutaActual = DameValor("[@" + csVariablesGlobales.Prefijo + "_PARAM]",
                                              "U_Ruta", "U_TipoRuta = 'Destino Ficheros Varios'") + csVariablesGlobales.NombreArchivoEmail;
                string RutaDestino = oAttachments2_Lines.SourcePath + @"\" + oAttachments2_Lines.FileName + "." + oAttachments2_Lines.FileExtension;
                File.Delete(RutaDestino);
                File.Move(RutaActual, RutaDestino);
                csVariablesGlobales.NombreArchivoEmail = "";
            }
        }

        public static string DameValorEditText(string Item, SAPbouiCOM.Form Formulario)
        {
            oForm = Formulario;
            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item(Item).Specific;
            return oEditText.String;
        }

        public static bool TesteoConexion()
        {
            if ((csVariablesGlobales.oCompany == null) || (csVariablesGlobales.oCompany.Connected == false))
            {
                csVariablesGlobales.oCompany = new SAPbobsCOM.Company();
                csVariablesGlobales.oCompany.SetSboLoginContext(csVariablesGlobales.SboApp.Company.GetConnectionContext(csVariablesGlobales.oCompany.GetContextCookie()));
                if (csVariablesGlobales.oCompany.Connect() != 0)
                {
                    csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                    return false;
                }
                return true;
            }
            return true;
        }
    }
}
