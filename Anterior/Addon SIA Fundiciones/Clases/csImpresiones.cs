using System;
using System.Collections.Generic;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Data;
using System.IO;
using System.Threading;

namespace Addon_SIA
{
    class csImpresiones
    {
        SAPbouiCOM.Form oForm = null;
        SAPbouiCOM.EditText oEditText = null;
        SAPbouiCOM.Matrix oMatrix = null;

        public CrystalDecisions.Shared.TableLogOnInfo ConInfo = new CrystalDecisions.Shared.TableLogOnInfo();

        public void CargarReports(string FormUID, string TypeEx, string Borrador, string DocEntry,
                                   string Menu, string StrIdioma)
        {
            //csUtilidades csUtilidades = new csUtilidades();
            int i;
            string StrBorrador;
            string StrTypeEx;
            System.Data.DataTable SqlTabla;
            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
            StrBorrador = Borrador;
            StrTypeEx = TypeEx;
            DataRow dr;
            string StrSql;
            try
            {
                oEditText = (SAPbouiCOM.EditText)(oForm.Items.Item("txtDoc").Specific);
                oEditText.String = DocEntry;
                oEditText = (SAPbouiCOM.EditText)(oForm.Items.Item("txtForm").Specific);
                oEditText.String = TypeEx;
                oEditText = (SAPbouiCOM.EditText)(oForm.Items.Item("txtBorr").Specific);
                oEditText.String = Borrador;
                oEditText = (SAPbouiCOM.EditText)(oForm.Items.Item("txtMenu").Specific);
                oEditText.String = Menu;
                StrSql = "SELECT U_Report, U_Descrip FROM [@" + csVariablesGlobales.Prefijo + 
                        "_REPORT] where u_Borrador='" + StrBorrador +
                        "' AND U_TipDoc='" + StrTypeEx + "' GROUP BY U_Report, U_Descrip HAVING Right(U_Report, 6)='" +
                        StrIdioma + ".rpt'";
                SqlTabla = new System.Data.DataTable();
                csUtilidades.CargarRecordSet(ref SqlTabla, StrSql);
                oMatrix = (SAPbouiCOM.Matrix)(oForm.Items.Item("matRep").Specific);
                oMatrix.Clear();
                if (SqlTabla.Rows.Count > 0)
                {
                    for (i = 0; i < SqlTabla.Rows.Count; i++)
                    {
                        dr = SqlTabla.Rows[i];
                        csVariablesGlobales.oUserDataSourceNombreReport.Value = Convert.ToString(dr["U_Report"].ToString());
                        csVariablesGlobales.oUserDataSourceDescripcionReport.Value = Convert.ToString(dr["U_Descrip"].ToString());
                        oMatrix.AddRow(1, -1);
                    }
                }
                StrSql = "SELECT U_Report, U_Descrip FROM [@" + csVariablesGlobales.Prefijo + 
                        "_REPORT] where u_Borrador='" + StrBorrador +
                        "' AND U_TipDoc='" + StrTypeEx + "' GROUP BY U_Report, U_Descrip HAVING Right(U_Report, 6)<>'" +
                        StrIdioma + ".rpt'";
                SqlTabla = new System.Data.DataTable();
                csUtilidades.CargarRecordSet(ref SqlTabla, StrSql);
                if (SqlTabla.Rows.Count > 0)
                {
                    for (i = 0; i < SqlTabla.Rows.Count; i++)
                    {
                        dr = SqlTabla.Rows[i];
                        csVariablesGlobales.oUserDataSourceNombreReport.Value = Convert.ToString(dr["U_Report"].ToString());
                        csVariablesGlobales.oUserDataSourceDescripcionReport.Value = Convert.ToString(dr["U_Descrip"].ToString());
                        oMatrix.AddRow(1, -1);
                    }
                }
            }
            catch (Exception ex)
            {
                csVariablesGlobales.SboApp.MessageBox("Error al comprobar los datos, " + ex.Message, 1, "Ok", "", "");
            }
        }

        public bool ImprimirDocumento(ref SAPbouiCOM.Application SBOApp, string Menu, string StrTabla,
                                  string DocEntry, string TypeEx, string StrReport, string AbsDocEntry,
                                  short NumeroCopias, string NombreParaExportar)
        {
            string StrFormula = "";
            string StrServer = csVariablesGlobales.oCompany.Server.ToString();
            string StrUser = csVariablesGlobales.oCompany.DbUserName.ToString();
            try
            {
                if (DocEntry != "")
                {
                    StrFormula = " {" + StrTabla + "." + AbsDocEntry + "}=" + DocEntry + " ";
                }
                switch (Menu)
                {
                    case "519": //Imprimir Por Pantalla:
                        Informe(StrReport, StrFormula);
                        break;
                    case "520": //Imprimir Por Impresora:
                        Imprimir(StrReport, StrFormula, DeterminarNumeroCopias(TypeEx), true);
                        break;
                    case "7176": //Exportar A Pdf:
                        ExportarAFormato(StrReport, StrFormula, "PDF", NombreParaExportar);
                        break;
                    case "7169": //Exportar A Excel:
                        ExportarAFormato(StrReport, StrFormula, "Excel", NombreParaExportar);
                        break;
                    case "7170": //Exportar A Word:
                        ExportarAFormato(StrReport, StrFormula, "Word", NombreParaExportar);
                        break;
                    case "6657": //Enviar E-mail:
                    case "6659": //Enviar Fax:
                        ExportarAFormato(StrReport, StrFormula, "PDF", NombreParaExportar);
                        break;
                }
                return true;
            }
            catch (Exception ex)
            {
                SBOApp.MessageBox("Error lanzando impresion en CR: " + ex.Message, 1, "Ok", "", "");
                return false;
            }
        }

        public short DeterminarNumeroCopias(string TypeEx)
        {
            short NumeroCopias = 1;
            string Condicion = "";
            string Seleccion = "";
            try
            {
                oForm = csVariablesGlobales.SboApp.Forms.Item(csVariablesGlobales.InstanciaFormularioSAP);
                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("4").Specific;
                Condicion = "CardCode = '" + oEditText.String + "'";

                //csUtilidades csUtilidades = new csUtilidades();
                switch (TypeEx)
                {
                    case "133":
                    case "179":
                    case "60090":
                    case "60091":
                    case "65300":
                        Seleccion = "U_NumCopFacVen";
                        break;
                    case "140":
                    case "180":
                        Seleccion = "U_NumCopAlbVen";
                        break;
                }
                if (Seleccion != "")
                {
                    NumeroCopias = Convert.ToInt16(csUtilidades.DameValor("OCRD", Seleccion, Condicion));
                }
                return NumeroCopias;
            }
            catch (Exception ex)
            {
                csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "Ok", "", "");
                return NumeroCopias;
            }
        }

        public void Impresion(ref SAPbouiCOM.Application SBOApp, ref SAPbouiCOM.MenuEvent pVal,
                                          out bool BubbleEvent)
        {
            BubbleEvent = true;
            string StrTabla;
            bool BolResult;
            string StrDocEntry;
            string StrAbsEntry;
            //csUtilidades csUtilidades = new csUtilidades();

            StrTabla = "";
            csVariablesGlobales.StrReport = "";
            switch (SBOApp.Forms.ActiveForm.TypeEx)
            {
                case "133":
                case "140":
                case "149":
                case "142":
                case "139":
                case "3002":
                case "SIASL10005":
                case "179":
                case "180":
                case "188":
                    //Documentos de ventas
                    StrDocEntry = csUtilidades.docEntry(csVariablesGlobales.SboApp);
                    if (csUtilidades.ExisteForm(SBOApp.Forms.ActiveForm.TypeEx, SBOApp, 
                        ref StrTabla, ref csVariablesGlobales.StrReport, ref StrDocEntry, pVal.MenuUID))
                    {
                        try
                        {
                            if (StrDocEntry != "")
                            {
                                if (StrDocEntry != "0")
                                {
                                    BolResult = ImprimirDocumento(ref SBOApp, pVal.MenuUID, StrTabla, 
                                                StrDocEntry, SBOApp.Forms.ActiveForm.TypeEx, 
                                                csVariablesGlobales.StrReport, "DocEntry", 1, "");
                                    BubbleEvent = false;
                                    SBOApp.SetStatusBarMessage("Documento " + StrDocEntry, 
                                                SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                }
                                else
                                {
                                    BubbleEvent = false;
                                }
                            }
                            else
                            {
                                SBOApp.SetStatusBarMessage("Error lanzando impresion en CR: Documento no encontrado ", BoMessageTime.bmt_Short, true);
                            }
                        }
                        catch (Exception ex)
                        {
                            SBOApp.MessageBox("Error lanzando impresion en CR: " + ex.Message, 1, "Ok", "", "");
                        }
                    }
                    break;
                case "60052": //Operaciones de Efectos
                    StrAbsEntry = csUtilidades.absEntry(csVariablesGlobales.SboApp);
                    StrDocEntry = StrAbsEntry;
                    if ((pVal.MenuUID == csVariablesGlobales.MenuImprimirPorPantalla ||
                        pVal.MenuUID == csVariablesGlobales.MenuImprimirPorImpresora) & pVal.BeforeAction)
                    {
                        if (csUtilidades.ExisteForm(SBOApp.Forms.ActiveForm.TypeEx, SBOApp, 
                            ref StrTabla, ref csVariablesGlobales.StrReport, ref StrDocEntry, pVal.MenuUID))
                        {
                            try
                            {
                                if (StrDocEntry != "")
                                {
                                    if (StrDocEntry != "0")
                                    {
                                        BolResult = ImprimirDocumento(ref SBOApp, pVal.MenuUID, StrTabla, 
                                                    StrDocEntry, SBOApp.Forms.ActiveForm.TypeEx, 
                                                    csVariablesGlobales.StrReport, "AbsEntry", 1, "");
                                        BubbleEvent = false;
                                        SBOApp.SetStatusBarMessage("Documento " + StrDocEntry, 
                                                    SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                    }
                                    else
                                    {
                                        BubbleEvent = false;
                                    }
                                }
                                else
                                {
                                    SBOApp.SetStatusBarMessage("Error lanzando impresion en CR: Documento no encontrado ", BoMessageTime.bmt_Short, true);
                                }
                            }
                            catch (Exception ex)
                            {
                                SBOApp.MessageBox("Error lanzando impresion en CR: " + ex.Message, 1, "Ok", "", "");
                            }
                        }
                    }
                    break;
                default:
                    break;
            }
        }

        public void Imprimir(string pReport, string pSeleccion, short NumeroCopias, bool SeleccionarImpresora)
        {
            string Impresora = csVariablesGlobales.ImpresoraPorDefecto;
            if (SeleccionarImpresora)
            {
                csPrintDialog oClsDialog = new csPrintDialog();
                System.Windows.Forms.DialogResult dialogResult;
                dialogResult = oClsDialog.Impresora(NumeroCopias);
                if (dialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    Impresora = oClsDialog.StrNomImp;
                    NumeroCopias = oClsDialog.NumCop;
                }
                else
                {
                    return;
                }
            }
            try
            {
                //csUtilidades csUtilidades = new csUtilidades();
                csUtilidades.LeerConexion(false);
                //ConInfo.ConnectionInfo.ServerName = csVariablesGlobales.oCompany.Server;
                //ConInfo.ConnectionInfo.DatabaseName = csVariablesGlobales.oCompany.CompanyDB;
                //ConInfo.ConnectionInfo.UserID = csVariablesGlobales.oCompany.DbUserName;
                //ConInfo.ConnectionInfo.Password = csVariablesGlobales.DBPassword;
                if (csVariablesGlobales.crReport != null && csVariablesGlobales.crReport.IsLoaded)
                {
                    csVariablesGlobales.crReport.Close();
                    GC.Collect();
                }
                CargarReport(pReport, pSeleccion);
                csVariablesGlobales.crReport.PrintOptions.PrinterName = Impresora;
                csVariablesGlobales.crReport.PrintToPrinter(NumeroCopias, false, 0, 0);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public bool Informe(string pReport, string pSeleccion)
        {
            try
            {
                //csUtilidades csUtilidades = new csUtilidades();
                //csUtilidades.LeerConexion(false);
                //ConInfo.ConnectionInfo.ServerName = csVariablesGlobales.oCompany.Server;
                //ConInfo.ConnectionInfo.DatabaseName = csVariablesGlobales.oCompany.CompanyDB;
                //ConInfo.ConnectionInfo.UserID = csVariablesGlobales.oCompany.DbUserName;
                //ConInfo.ConnectionInfo.Password = csVariablesGlobales.DBPassword;
                if (csVariablesGlobales.crReport != null && csVariablesGlobales.crReport.IsLoaded)
                {
                    csVariablesGlobales.crReport.Close();
                    GC.Collect();
                }
                if (!CargarReport(pReport, pSeleccion))
                {
                    return false;
                }

                Object thisForm = this;
                frmReportViewer formReportViewer = new frmReportViewer(false, ref thisForm);
                csVariablesGlobales.LanzarImpresionCrystal = true;
                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return false;
            }
        }

        public void ExportarAFormato(string pReport, string pSeleccion, string Formato, string NombreParaExportar)
        {
            try
            {
                //csUtilidades csUtilidades = new csUtilidades();
                csUtilidades.LeerConexion(false);
                //ConInfo.ConnectionInfo.ServerName = csVariablesGlobales.oCompany.Server;
                //ConInfo.ConnectionInfo.DatabaseName = csVariablesGlobales.oCompany.CompanyDB;
                //ConInfo.ConnectionInfo.UserID = csVariablesGlobales.oCompany.DbUserName;
                //ConInfo.ConnectionInfo.Password = csVariablesGlobales.DBPassword;

                //Para liberar la memoria, que crystal la deja
                if (csVariablesGlobales.crReport != null && csVariablesGlobales.crReport.IsLoaded)
                {
                    csVariablesGlobales.crReport.Close();
                    GC.Collect();
                }
                CargarReport(pReport, pSeleccion);
                string NombreArchivo = csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_PARAM]",
                                                              "U_Ruta", "U_TipoRuta = 'Destino Ficheros Varios'") + @"\" +
                                                              NombreParaExportar;
                switch (Formato)
                {
                    case "Word":
                        csVariablesGlobales.crReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.WordForWindows,
                                                                  NombreArchivo + ".doc");
                        break;
                    case "Excel":
                        csVariablesGlobales.crReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.Excel,
                                                                  NombreArchivo + ".xls");
                        break;
                    case "PDF":
                        csVariablesGlobales.crReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat,
                                                                  NombreArchivo + ".pdf");
                        csVariablesGlobales.NombreArchivoEmail = NombreParaExportar + ".pdf";
                        break;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public bool CargarReport(string pReport, string pSeleccion)
        {
            try
            {
                csVariablesGlobales.crReport = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                csVariablesGlobales.crReport.Load(pReport);
                csVariablesGlobales.crReport.Refresh();

                csVariablesGlobales.crReport.DataSourceConnections[0].
                    SetConnection(csVariablesGlobales.oCompany.Server, csVariablesGlobales.oCompany.CompanyDB, 
                    csVariablesGlobales.oCompany.DbUserName, csVariablesGlobales.DBPassword);

                if (pSeleccion != "")
                {
                    csVariablesGlobales.crReport.RecordSelectionFormula = pSeleccion;
                }
                return true;
            }
            catch (Exception ex)
            {
                csVariablesGlobales.SboApp.MessageBox("Error lanzando impresion en CR: " + ex.Message, 1, "Ok", "", "");
                return false;
            }
        }
    }
}
