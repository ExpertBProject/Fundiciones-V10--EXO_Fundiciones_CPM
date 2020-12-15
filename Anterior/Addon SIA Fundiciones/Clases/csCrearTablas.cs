using System;
using System.Collections.Generic;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Addon_SIA
{
    public class csCrearTablas
    {
        SAPbobsCOM.UserTablesMD oUserTablesMD;
        SAPbobsCOM.UserFieldsMD oUserFieldsMD;
        SAPbobsCOM.UserKeysMD oUserKeysMD;
        SAPbobsCOM.UserQueries oUserQueries;
        SAPbobsCOM.SBObob oSBObob;
        SAPbobsCOM.Recordset oRecordSet;
        SAPbobsCOM.ValidValuesMD oValidValuesMD;

        public bool CrearTablasXml(string sFichero)
        {
            int iElementos, i;
            UserQueries oUserQuery2;
            csCrearConsultas CrearConsultas = new csCrearConsultas();

            csVariablesGlobales.oCompany.StartTransaction();

            try
            {
                iElementos = csVariablesGlobales.oCompany.GetXMLelementCount(sFichero);

                for (i = 0; i < iElementos; i++)
                {
                    switch (csVariablesGlobales.oCompany.GetXMLobjectType(sFichero, i))
                    {
                        case SAPbobsCOM.BoObjectTypes.oUserTables:
                            oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(csVariablesGlobales.oCompany.GetBusinessObjectFromXML(sFichero, i)));
                            System.Diagnostics.Debug.Print(oUserTablesMD.TableName);
                            if (ExisteTabla(oUserTablesMD.TableName) == false)
                            {
                                if (oUserTablesMD.Add() != 0)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                                    csVariablesGlobales.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    return false;
                                }
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                            oUserTablesMD = null;
                            GC.Collect();
                            break;
                        case SAPbobsCOM.BoObjectTypes.oUserFields:
                            oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)(csVariablesGlobales.oCompany.GetBusinessObjectFromXML(sFichero, i));
                            if (ExisteCampo(oUserFieldsMD.TableName, oUserFieldsMD.Name) == false)
                            {
                                if (oUserFieldsMD.Add() != 0)
                                {
                                    MessageBox.Show(csVariablesGlobales.oCompany.GetLastErrorDescription());
                                    csVariablesGlobales.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                                    return false;
                                }
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                            oUserFieldsMD = null;
                            GC.Collect();

                            break;
                        case BoObjectTypes.oUserKeys:
                            oUserKeysMD = (SAPbobsCOM.UserKeysMD)(csVariablesGlobales.oCompany.GetBusinessObjectFromXML(sFichero, i));
                            System.Diagnostics.Debug.Print(oUserKeysMD.KeyName);
                            if (ExisteTabla(oUserKeysMD.KeyName) == false)
                            {
                                if (oUserKeysMD.Add() != 0)
                                {
                                    csVariablesGlobales.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD);
                                    return false;
                                }
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD);
                            oUserKeysMD = null;
                            GC.Collect();

                            break;
                        case BoObjectTypes.oUserQueries:
                            oUserQueries = (SAPbobsCOM.UserQueries)(csVariablesGlobales.oCompany.GetBusinessObjectFromXML(sFichero, i));
                            System.Diagnostics.Debug.Print(oUserQueries.QueryDescription);
                            oUserQuery2 = CrearConsultas.ExisteConsulta(oUserQueries.QueryDescription);
                            if (oUserQuery2 == null)
                            {
                                if (oUserQueries.Add() != 0)
                                {
                                    csVariablesGlobales.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserQueries);
                                    try
                                    {
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserQuery2);
                                    }
                                    catch
                                    {
                                    }
                                    return false;
                                }
                            }
                            else
                            {
                                oUserQuery2.Query = oUserQueries.Query;
                                if (oUserQuery2.Update() != 0)
                                {
                                    csVariablesGlobales.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserQueries);
                                    try
                                    {
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserQuery2);
                                    }
                                    catch
                                    {
                                    }
                                    return false;
                                }
                            }
                            break;
                    }
                    System.Diagnostics.Debug.Print(i.ToString());
                }
                csVariablesGlobales.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                return true;
            }
            catch
            {
                if (csVariablesGlobales.oCompany.InTransaction == true)
                {
                    csVariablesGlobales.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                return false;
            }
        }

        public bool ExisteTabla(string sTabla)
        {
            try
            {
                oSBObob = ((SAPbobsCOM.SBObob)(csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoBridge)));
                oRecordSet = oSBObob.GetTableList();
                oRecordSet.MoveFirst();
                while (!oRecordSet.EoF)
                {
                    if (Convert.ToString(oRecordSet.Fields.Item(0).Value).ToUpper() == Convert.ToString(sTabla).ToUpper())
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oSBObob);
                        return true;
                    }
                    oRecordSet.MoveNext();
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSBObob);
                return false;
            }
            catch
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSBObob);
                return false;
            }
        }

        public bool ExisteCampo(string sTabla, string sNombre)
        {
            try
            {
                oSBObob = ((SAPbobsCOM.SBObob)(csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoBridge)));
                oRecordSet = oSBObob.GetTableFieldList(sTabla);
                oRecordSet.MoveFirst();
                while (!oRecordSet.EoF)
                {
                    if (Convert.ToString(oRecordSet.Fields.Item(0).Value).ToUpper() == Convert.ToString("U_" + sNombre).ToUpper())
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oSBObob);
                        return true;
                    }
                    oRecordSet.MoveNext();
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSBObob);
                return false;
            }
            catch
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSBObob);
                return false;
            }
        }

        public bool ExisteIndice(SAPbobsCOM.Company oCompany, string sKeyName)
        {
            oSBObob = ((SAPbobsCOM.SBObob)(oCompany.GetBusinessObject(BoObjectTypes.BoBridge)));
            oRecordSet = oSBObob.GetObjectKeyBySingleValue(BoObjectTypes.oUserKeys, "KeyName", sKeyName, BoQueryConditions.bqc_Equal);
            try
            {
                while (!oRecordSet.EoF)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSBObob);
                    return true;
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSBObob);
                return false;
            }
            catch
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSBObob);
                return false;
            }
        }

        public bool CrearCampo(string Tabla, string DescripcionTabla, string Campo,
                                string DescripcionCampo, SAPbobsCOM.BoFieldTypes FieldTypes, int LongCampo,
                                SAPbobsCOM.BoFldSubTypes FldSubType, string TablaRelacionada, bool Boleano,
                                string ValorPorDefecto, ref bool TablaModiFicada)
        {
            int Error;
            csVariablesGlobales.SboApp.SetStatusBarMessage("Campo -> " + DescripcionCampo + ", Tabla -> " +
                                                           DescripcionTabla, BoMessageTime.bmt_Short, false);
            if (!ExisteCampo(Tabla, Campo))
            {
                try
                {
                    //csUtilidades csUtilidades = new csUtilidades();
                    csUtilidades.LeerConexion(true);
                    oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.oUserFields));
                    oUserFieldsMD.TableName = Tabla;
                    oUserFieldsMD.Name = Campo;
                    oUserFieldsMD.Description = DescripcionCampo;
                    oUserFieldsMD.Type = FieldTypes;
                    oUserFieldsMD.EditSize = LongCampo;
                    oUserFieldsMD.DefaultValue = ValorPorDefecto;
                    if (FldSubType.ToString() != "")
                    {
                        oUserFieldsMD.SubType = FldSubType;
                    }
                    if (TablaRelacionada != "")
                    {
                        oUserFieldsMD.LinkedTable = TablaRelacionada;
                    }
                    if (Boleano == true)
                    {
                        oValidValuesMD = oUserFieldsMD.ValidValues;
                        oValidValuesMD.Description = "Sí";
                        oValidValuesMD.Value = "Y";
                        oValidValuesMD.Add();
                        oValidValuesMD = oUserFieldsMD.ValidValues;
                        oValidValuesMD.Description = "No";
                        oValidValuesMD.Value = "N";
                        oValidValuesMD.Add();
                        oUserFieldsMD.DefaultValue = ValorPorDefecto;
                    }
                    Error = oUserFieldsMD.Add();
                    if (Error != 0)
                    {
                        csVariablesGlobales.oCompany.GetLastError(out Error, out csVariablesGlobales.StrMsError);
                        csVariablesGlobales.oCompany.Disconnect();
                        csVariablesGlobales.oCompany = null;
                        GC.Collect();
                        return false;
                    }
                    else
                    {
                        csVariablesGlobales.oCompany.Disconnect();
                        csVariablesGlobales.oCompany = null;
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        TablaModiFicada = true;
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    if (csVariablesGlobales.oCompany != null)
                    {
                        if (csVariablesGlobales.oCompany.Connected == true)
                        {
                            csVariablesGlobales.oCompany.Disconnect();
                        }
                    }
                    csVariablesGlobales.oCompany = null;
                    GC.Collect();
                    csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "OK", "", "");
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        public bool CrearTabla(string Tabla, string DescripcionTabla,
                               BoUTBTableType UtbTableType, string TablaUsuarioOSistema)
        {
            int Error;
            
            csVariablesGlobales.SboApp.SetStatusBarMessage("Tabla " + DescripcionTabla, BoMessageTime.bmt_Short, false);
            if (!ExisteTabla(TablaUsuarioOSistema + Tabla))
            {
                try
                {
                    //csUtilidades csUtilidades = new csUtilidades();
                    csUtilidades.LeerConexion(true);
                    oUserTablesMD = ((SAPbobsCOM.UserTablesMD)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.oUserTables));
                    oUserTablesMD.TableName = Tabla;
                    oUserTablesMD.TableType = UtbTableType;
                    oUserTablesMD.TableDescription = DescripcionTabla;
                    Error = oUserTablesMD.Add();
                    if (Error != 0)
                    {
                        csVariablesGlobales.oCompany.GetLastError(out Error, out csVariablesGlobales.StrMsError);
                        csVariablesGlobales.oCompany.Disconnect();
                        csVariablesGlobales.oCompany = null;
                        //GC.WaitForPendingFinalizers();
                        GC.Collect();
                        return false;
                    }
                    else
                    {
                        csVariablesGlobales.oCompany.Disconnect();
                        csVariablesGlobales.oCompany = null;
                        GC.Collect();
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    if (csVariablesGlobales.oCompany.Connected == true)
                    {
                        csVariablesGlobales.oCompany.Disconnect();
                    }
                    csVariablesGlobales.oCompany = null;
                    GC.Collect();
                    csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "OK", "", "");
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        public void TablaReports() //Tabla de Usuario
        {
            bool TablaModiFicada = false;
            CrearTabla(csVariablesGlobales.Prefijo + "_REPORT", "Report", BoUTBTableType.bott_NoObject, "@");
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_REPORT", "Report", "Report", "Report", BoFieldTypes.db_Alpha, 250, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_REPORT", "Report", "Descrip", "Descripción", BoFieldTypes.db_Alpha, 250, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_REPORT", "Report", "TipDoc", "Tipo Documento", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_REPORT", "Report", "Borrador", "Borrador", BoFieldTypes.db_Alpha, 1, BoFldSubTypes.st_None, "", true, "N", ref TablaModiFicada);
            if (TablaModiFicada)
            {
                #region Script Correción Tabla
                string Script;
                Script = "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_REPORT] ALTER COLUMN U_Report nvarchar(250); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_REPORT] ALTER COLUMN U_Descrip nvarchar(250); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_REPORT] ALTER COLUMN U_TipDoc nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_REPORT]	DROP CONSTRAINT DF_@" + csVariablesGlobales.Prefijo + "_REPORT_U_Borrador;" +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_REPORT] ALTER COLUMN U_Borrador nvarchar(1); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_REPORT] ADD CONSTRAINT DF_@" + csVariablesGlobales.Prefijo + "_REPORT_U_Borrador DEFAULT ('N') FOR U_Borrador;";
                SqlCommand cmdCorrecionFormularios = new SqlCommand(Script, csVariablesGlobales.conAddon);
                cmdCorrecionFormularios.ExecuteNonQuery();
                #endregion
            }
        }

        public void TablaParametros() //Tabla de Usuario
        {
            bool TablaModiFicada = false;
            CrearTabla(csVariablesGlobales.Prefijo + "_PARAM", "Parámetros", BoUTBTableType.bott_NoObject, "@");
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_PARAM", "Parámetros", "Concepto", "Concepto", BoFieldTypes.db_Alpha, 250, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_PARAM", "Parámetros", "Valor", "Valor", BoFieldTypes.db_Alpha, 250, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_PARAM", "Parámetros", "Ruta", "Ruta", BoFieldTypes.db_Alpha, 250, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_PARAM", "Parámetros", "TipoRuta", "Tipo Ruta", BoFieldTypes.db_Alpha, 250, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            //CrearCampo("@" + csVariablesGlobales.Prefijo + "_PARAM", "Parámetros", "CryVen", "Doc. Ventas Crystal", BoFieldTypes.db_Alpha, 1, BoFldSubTypes.st_None, "", true, "N", ref TablaModiFicada);
            //CrearCampo("@" + csVariablesGlobales.Prefijo + "_PARAM", "Parámetros", "CryCom", "Doc. Compras Crystal", BoFieldTypes.db_Alpha, 1, BoFldSubTypes.st_None, "", true, "N", ref TablaModiFicada);
            //CrearCampo("@" + csVariablesGlobales.Prefijo + "_PARAM", "Parámetros", "Per349", "Periodo 349", BoFieldTypes.db_Alpha, 1, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);

            if (TablaModiFicada)
            {
                #region Script Correción Tabla
                string Script;
                Script = //"ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PARAM] DROP CONSTRAINT DF_@" + csVariablesGlobales.Prefijo + "_PARAM_U_CryVen; " +
                         //"ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PARAM] DROP CONSTRAINT DF_@" + csVariablesGlobales.Prefijo + "_PARAM_U_CryCom; " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PARAM] ALTER COLUMN U_Concepto nvarchar(250); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PARAM] ALTER COLUMN U_Valor nvarchar(250); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PARAM] ALTER COLUMN U_Ruta nvarchar(250); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PARAM] ALTER COLUMN U_TipoRuta nvarchar(250); " ; //+
                         //"ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PARAM] ALTER COLUMN U_CryVen nvarchar(1); " +
                         //"ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PARAM] ALTER COLUMN U_CryCom nvarchar(1); " +
                         //"ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PARAM] ALTER COLUMN U_Per349 nvarchar(1); " +
                         //"ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PARAM] ADD CONSTRAINT DF_@" + csVariablesGlobales.Prefijo + "_PARAM_U_CryVen DEFAULT ('N') FOR U_CryVen; " +
                         //"ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PARAM] ADD CONSTRAINT DF_@" + csVariablesGlobales.Prefijo + "_PARAM_U_CryCom DEFAULT ('N') FOR U_CryCom; ";
                SqlCommand cmdCorrecionFormularios = new SqlCommand(Script, csVariablesGlobales.conAddon);
                cmdCorrecionFormularios.ExecuteNonQuery();
                #endregion
            }
        }

        public void TablaFormularios() //Tabla de Usuario
        {
            bool TablaModiFicada = false;
            CrearTabla(csVariablesGlobales.Prefijo + "_FORMS", "Formularios", BoUTBTableType.bott_NoObject, "@");
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_FORMS", "Formularios", "formSAP", "Formulario SAP", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_FORMS", "Formularios", "formSIA", "Formulario SIA", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_FORMS", "Formularios", "descrip", "Descripción Formulario", BoFieldTypes.db_Alpha, 250, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);

            if (TablaModiFicada)
            {
                #region Script Correción Tabla
                string Script;
                Script = "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_FORMS] ALTER COLUMN U_formSAP nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_FORMS] ALTER COLUMN U_formSIA nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_FORMS] ALTER COLUMN U_descrip nvarchar(250); ";
                SqlCommand cmdCorrecionFormularios = new SqlCommand(Script, csVariablesGlobales.conAddon);
                cmdCorrecionFormularios.ExecuteNonQuery();
                #endregion
            }
        }

        public void TablaLineasDocumentosMarketing() //Tabla de SAP
        {
            bool TablaModiFicada = false;
            #region Descuentos
            if (csVariablesGlobales.CrearDtosEnDocumentos == "S")
            {
                CrearCampo("DLN1", "Tabla Líneas de Documentos", "Dto1", "% Dto. 1", BoFieldTypes.db_Float, 10, BoFldSubTypes.st_Percentage, "", false, "", ref TablaModiFicada);
                CrearCampo("DLN1", "Tabla Líneas de Documentos", "Dto2", "% Dto. 2", BoFieldTypes.db_Float, 10, BoFldSubTypes.st_Percentage, "", false, "", ref TablaModiFicada);
                CrearCampo("DLN1", "Tabla Líneas de Documentos", "Dto3", "% Dto. 3", BoFieldTypes.db_Float, 10, BoFldSubTypes.st_Percentage, "", false, "", ref TablaModiFicada);
                CrearCampo("DLN1", "Tabla Líneas de Documentos", "Dto4", "% Dto. 4", BoFieldTypes.db_Float, 10, BoFldSubTypes.st_Percentage, "", false, "", ref TablaModiFicada);
                CrearCampo("DLN1", "Tabla Líneas de Documentos", "Dto5", "% Dto. 5", BoFieldTypes.db_Float, 10, BoFldSubTypes.st_Percentage, "", false, "", ref TablaModiFicada);
                CrearCampo("DLN1", "Tabla Líneas de Documentos", "PreTar", "Precio Tarifa", BoFieldTypes.db_Float, 10, BoFldSubTypes.st_Percentage, "", false, "", ref TablaModiFicada);
            }
            #endregion
        }

        public void TablaInterlocutorComercial() //Tabla de SAP
        {
            bool TablaModiFicada = false;
            CrearCampo("OCRD", "Tabla Interlocutores Comerciales", "NumCopFacVen", "Nº Copias Facturas Ventas", BoFieldTypes.db_Numeric, 10, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("OCRD", "Tabla Interlocutores Comerciales", "NumCopAlbVen", "Nº Copias Albaranes Ventas", BoFieldTypes.db_Numeric, 10, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
        }

        public void TablaUbicaciones() //Tabla de Usuario
        {
            bool TablaModiFicada = false;
            CrearTabla(csVariablesGlobales.Prefijo + "_UBIC", "Ubicaciones", BoUTBTableType.bott_NoObject, "@");
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_UBIC", "Ubicaciones", "cod_ubic", "Código Ubicación", BoFieldTypes.db_Alpha, 8, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_UBIC", "Ubicaciones", "des", "Descripción", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_UBIC", "Ubicaciones", "almacen", "Almacén", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_UBIC", "Ubicaciones", "calle", "Calle", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_UBIC", "Ubicaciones", "hueco", "Hueco", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_UBIC", "Ubicaciones", "altura", "Altura", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            if (TablaModiFicada)
            {
                #region Script Correción Tabla
                string Script;
                Script = "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_UBIC] ALTER COLUMN U_cod_ubic nvarchar(8); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_UBIC] ALTER COLUMN U_des nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_UBIC] ALTER COLUMN U_almacen nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_UBIC] ALTER COLUMN U_calle nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_UBIC] ALTER COLUMN U_hueco nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_UBIC] ALTER COLUMN U_altura nvarchar(20); ";
                SqlCommand cmdCorrecionFormularios = new SqlCommand(Script, csVariablesGlobales.conAddon);
                cmdCorrecionFormularios.ExecuteNonQuery();
                #endregion
            }
        }

        public void TablaPlanos() //Tabla de Usuario
        {
            bool TablaModiFicada = false;
            CrearTabla(csVariablesGlobales.Prefijo + "_PLANO", "Planos", BoUTBTableType.bott_NoObject, "@");
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_PLANO", "Planos", "cod_plano", "Código Plano", BoFieldTypes.db_Alpha, 15, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_PLANO", "Planos", "des", "Descripción", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_PLANO", "Planos", "cod_artic", "Código Artículo", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_PLANO", "Planos", "SIA_UltRev", "Última Revisión", BoFieldTypes.db_Date, 10, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            if (TablaModiFicada)
            {
                #region Script Correción Tabla
                string Script;
                Script = "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PLANO] ALTER COLUMN U_cod_plano nvarchar(15); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PLANO] ALTER COLUMN U_des nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PLANO] ALTER COLUMN U_cod_artic nvarchar(20); ";
                SqlCommand cmdCorrecionFormularios = new SqlCommand(Script, csVariablesGlobales.conAddon);
                cmdCorrecionFormularios.ExecuteNonQuery();
                #endregion
            }
        }

        public void TablaModelo347() //Tabla de Usuario
        {
            bool TablaModiFicada = false;
            CrearTabla(csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", BoUTBTableType.bott_NoObject, "@");
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "CodIC", "Código IC", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "NomIC", "Nombre IC", BoFieldTypes.db_Alpha, 50, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "TipIC", "Tipo IC", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "NifIC", "N.I.F. IC", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "Base", "Base", BoFieldTypes.db_Float, 19, BoFldSubTypes.st_Sum, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "Prov", "Provincia", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "Pais", "País", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "Indic", "Indicador", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "CP", "Código Postal", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "CodInf", "Código Informe", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "Iva", "I.V.A.", BoFieldTypes.db_Float, 19, BoFldSubTypes.st_Sum, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "Total", "Total", BoFieldTypes.db_Float, 19, BoFldSubTypes.st_Sum, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "TipDoc", "Tipo Documento", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "Anticip", "Anticipo", BoFieldTypes.db_Float, 19, BoFldSubTypes.st_Sum, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD347", "Modelo 347", "Adquisi", "Adquisición", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);

            if (TablaModiFicada)
            {
                #region Script Correción Tabla
                string Script;
                Script = "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD347] ALTER COLUMN U_CodIC nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD347] ALTER COLUMN U_NomIC nvarchar(50); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD347] ALTER COLUMN U_TipIC nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD347] ALTER COLUMN U_NifIC nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD347] ALTER COLUMN U_Prov nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD347] ALTER COLUMN U_Pais nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD347] ALTER COLUMN U_Indic nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD347] ALTER COLUMN U_CP nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD347] ALTER COLUMN U_CodInf nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD347] ALTER COLUMN U_TipDoc nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD347] ALTER COLUMN U_Adquisi nvarchar(20); ";
                SqlCommand cmdCorrecionFormularios = new SqlCommand(Script, csVariablesGlobales.conAddon);
                cmdCorrecionFormularios.ExecuteNonQuery();
                #endregion
            }
        }

        public void TablaEstructuraConsultas() //Tabla de Usuario
        {
            bool TablaModiFicada = false;
            CrearTabla(csVariablesGlobales.Prefijo + "_ESTCON", "Estructura Consultas", BoUTBTableType.bott_NoObject, "@");
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_ESTCON", "Estructura Consultas", "des", "Descripción", BoFieldTypes.db_Alpha, 250, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_ESTCON", "Estructura Consultas", "ident", "Identificador", BoFieldTypes.db_Alpha, 10, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_ESTCON", "Estructura Consultas", "nom_sql", "Nombre SQL", BoFieldTypes.db_Alpha, 250, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_ESTCON", "Estructura Consultas", "cat_sql", "Categoría SQL", BoFieldTypes.db_Alpha, 250, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_ESTCON", "Estructura Consultas", "tipo", "Tipo Consulta", BoFieldTypes.db_Alpha, 250, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada); if (TablaModiFicada)
            if (TablaModiFicada)
            {
                #region Script Correción Tabla
                string Script;
                Script = "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_ESTCON] ALTER COLUMN U_des nvarchar(250); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_ESTCON] ALTER COLUMN U_ident nvarchar(10); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_ESTCON] ALTER COLUMN U_nom_sql nvarchar(250); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_ESTCON] ALTER COLUMN U_cat_sql nvarchar(250); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_ESTCON] ALTER COLUMN U_tipo nvarchar(250); ";
                SqlCommand cmdCorrecionFormularios = new SqlCommand(Script, csVariablesGlobales.conAddon);
                cmdCorrecionFormularios.ExecuteNonQuery();
                #endregion
            }
        }

        public void TablaDestinoMemoria() //Tabla de Usuario
        {
            bool TablaModiFicada = false;
            CrearTabla(csVariablesGlobales.Prefijo + "_DESMEM", "Destino Memoria", BoUTBTableType.bott_NoObject, "@");
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_DESMEM", "Destino Memoria", "ident", "Identificador", BoFieldTypes.db_Alpha, 10, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_DESMEM", "Destino Memoria", "valor", "Valor", BoFieldTypes.db_Alpha, 250, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            if (TablaModiFicada)
            {
                #region Script Correción Tabla
                string Script;
                Script = "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_DESMEM] ALTER COLUMN U_ident nvarchar(10); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_DESMEM] ALTER COLUMN U_valor nvarchar(250); ";
                SqlCommand cmdCorrecionFormularios = new SqlCommand(Script, csVariablesGlobales.conAddon);
                cmdCorrecionFormularios.ExecuteNonQuery();
                #endregion
            }
        }

        public void TablaBalances() //Tabla de Usuario
        {
            bool TablaModiFicada = false;
            CrearTabla(csVariablesGlobales.Prefijo + "_BALAN", "Grupos Nivel 3 para Balance", BoUTBTableType.bott_NoObject, "@");
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_BALAN", "Destino Memoria", "SIA_Grupo", "Grupo", BoFieldTypes.db_Alpha, 10, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_BALAN", "Destino Memoria", "SIA_Nomgrupo", "Descripción", BoFieldTypes.db_Alpha, 254, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            if (TablaModiFicada)
            {
                #region Script Correción Tabla
                string Script;
                Script = "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_BALAN] ALTER COLUMN U_SIA_Grupo nvarchar(10); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_BALAN] ALTER COLUMN U_SIA_Nomgrupo nvarchar(254); ";
                SqlCommand cmdCorrecionFormularios = new SqlCommand(Script, csVariablesGlobales.conAddon);
                cmdCorrecionFormularios.ExecuteNonQuery();
                #endregion
            }
        }

        public void TablaPedidoComprasDevuelta() //Tabla de Usuario
        {
            bool TablaModiFicada = false;
            CrearTabla(csVariablesGlobales.Prefijo + "_PEDC", "Pedido Compras Devuelta", BoUTBTableType.bott_NoObject, "@");
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_PEDC", "Pedido Compras Devuelta", "SIA_NumPed", "Nº Pedido", BoFieldTypes.db_Numeric, 10, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_PEDC", "Pedido Compras Devuelta", "SIA_NumLin", "Nº Linea", BoFieldTypes.db_Numeric, 10, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_PEDC", "Pedido Compras Devuelta", "SIA_CantDev", "Cantidad Devuelta", BoFieldTypes.db_Numeric, 10, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_PEDC", "Pedido Compras Devuelta", "SIA_YaServ", "¿Ya Servida?", BoFieldTypes.db_Alpha, 10, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_PEDC", "Pedido Compras Devuelta", "SIA_FecEnt", "Fecha Entrega", BoFieldTypes.db_Date, 10, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            if (TablaModiFicada)
            {
                #region Script Correción Tabla
                string Script;
                Script = "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_PEDC] ALTER COLUMN U_SIA_YaServ nvarchar(10); ";
                SqlCommand cmdCorrecionFormularios = new SqlCommand(Script, csVariablesGlobales.conAddon);
                cmdCorrecionFormularios.ExecuteNonQuery();
                #endregion
            }
        }

        public void TablaModelo349() //Tabla de Usuario
        {
            bool TablaModiFicada = false;
            CrearTabla(csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", BoUTBTableType.bott_NoObject, "@");
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "CodIC", "Código IC", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "NomIC", "Nombre IC", BoFieldTypes.db_Alpha, 50, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "TipIC", "Tipo IC", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "NifIC", "N.I.F. IC", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "Base", "Base", BoFieldTypes.db_Float, 19, BoFldSubTypes.st_Sum, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "Prov", "Provincia", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "Pais", "País", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "Indic", "Indicador", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "CP", "Código Postal", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "CodInf", "Código Informe", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "Iva", "I.V.A.", BoFieldTypes.db_Float, 19, BoFldSubTypes.st_Sum, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "Total", "Total", BoFieldTypes.db_Float, 19, BoFldSubTypes.st_Sum, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "TipDoc", "Tipo Documento", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "Anticip", "Anticipo", BoFieldTypes.db_Float, 19, BoFldSubTypes.st_Sum, "", false, "", ref TablaModiFicada);
            CrearCampo("@" + csVariablesGlobales.Prefijo + "_MOD349", "Modelo 349", "Adquisi", "Adquisición", BoFieldTypes.db_Alpha, 20, BoFldSubTypes.st_None, "", false, "", ref TablaModiFicada);

            if (TablaModiFicada)
            {
                #region Script Correción Tabla
                string Script;
                Script = "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD349] ALTER COLUMN U_CodIC nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD349] ALTER COLUMN U_NomIC nvarchar(50); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD349] ALTER COLUMN U_TipIC nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD349] ALTER COLUMN U_NifIC nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD349] ALTER COLUMN U_Prov nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD349] ALTER COLUMN U_Pais nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD349] ALTER COLUMN U_Indic nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD349] ALTER COLUMN U_CP nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD349] ALTER COLUMN U_CodInf nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD349] ALTER COLUMN U_TipDoc nvarchar(20); " +
                         "ALTER TABLE [@" + csVariablesGlobales.Prefijo + "_MOD349] ALTER COLUMN U_Adquisi nvarchar(20); ";
                SqlCommand cmdCorrecionFormularios = new SqlCommand(Script, csVariablesGlobales.conAddon);
                cmdCorrecionFormularios.ExecuteNonQuery();
                #endregion
            }
        }
    }
}