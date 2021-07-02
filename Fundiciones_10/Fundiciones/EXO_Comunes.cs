using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Text.RegularExpressions;



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

      public static double StringSAPToDoubleSistema(string Texto, string cMoneda)
      {
          double nRetorno = 0;          
          string  Cadena = Texto.Replace((cMoneda != "") ? cMoneda : "EUR", ""); 
          
          nRetorno = StringSAPToDoubleSistema(Cadena);

          return nRetorno;
      }

      public static string DoubleStringSAP(double Valor, BoFldSubTypes BoTipo, bool ForMatrix = false)
            {

                string cRetorno = "";

                switch (BoTipo)
                {
                    case BoFldSubTypes.st_Quantity:
                        Valor = Math.Round(Valor, VarGlobal.QtyDec);
                        break;
                    case BoFldSubTypes.st_Sum:
                        Valor = Math.Round(Valor, VarGlobal.SumDec);
                        break;
                    case BoFldSubTypes.st_Percentage:
                        Valor = Math.Round(Valor, VarGlobal.PercentDec);
                        break;
                    case BoFldSubTypes.st_Price:
                        Valor = Math.Round(Valor, VarGlobal.PriceDec);
                        break;
                    case BoFldSubTypes.st_Measurement:
                        Valor = Math.Round(Valor, VarGlobal.MeasureDec);
                        break;
                    case BoFldSubTypes.st_Rate:
                        Valor = Math.Round(Valor, VarGlobal.RateDec);
                        break;
                    default:
                        Valor = Math.Round(Valor, 2);
                        break;
                }

                cRetorno = Valor.ToString();
                if (!ForMatrix) cRetorno = cRetorno.Replace(',', '.');

                return cRetorno;

                //string cRetorno = "";
                //string cAux = Valor.ToString();

                //cRetorno = cAux.Replace(',', csVariablesGlobales.cSepDecimal);

                ////System.Globalization.NumberFormatInfo nfi = new System.Globalization.NumberFormatInfo();

                ////string hh = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
                ////string hh1 = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;

                ////string SepDecSAP = DevuelveValor("OADM", "DecSep", "");
                ////string SepMilSAP = DevuelveValor("OADM", "ThousSep", ""); 
                ////string SepDecSistema = nfi.NumberGroupSeparator;
                ////string SepMilSistema = nfi.NumberDecimalSeparator;
                ////string ValorDevuelto = Valor.ToString();
                ////if (SepDecSAP != SepDecSistema)
                ////{
                ////    ValorDevuelto = ValorDevuelto.Replace(SepDecSistema, SepDecSAP);
                ////}


            }

      public static string DateStringSAP(DateTime dFecha)
      {
          string cRetorno = dFecha.Year.ToString("0000") + dFecha.Month.ToString("00") + dFecha.Day.ToString("00");

          return cRetorno;
      }

      public static DateTime StringSAPDate(string cFechaSAP)
      {
          int nYear = Convert.ToInt32(cFechaSAP.Substring(0, 4));
          int nMes = Convert.ToInt32(cFechaSAP.Substring(4, 2));
          int nDia = Convert.ToInt32(cFechaSAP.Substring(6, 2));

          DateTime dRetorno = new DateTime(nYear, nMes, nDia);

          return dRetorno;
      }
    }

    public class TratamientoFicheros
    {
       public static string EscojoFichero(string cCadenaDefect)
            {
                string Ruta = "";

                EXO_SaveFileDialog oFichero = new EXO_SaveFileDialog();
                oFichero.Filter = "All Files (*)|*|Dat (*.dat)|*.dat|Text Files (*.txt)|*.txt";
                oFichero.FileName = cCadenaDefect;
                string DirectorioActual = Environment.CurrentDirectory;
                Thread threadGetFile = new Thread(new ThreadStart(oFichero.GetFileName));
                threadGetFile.TrySetApartmentState(ApartmentState.STA);
                threadGetFile.Start();
                try
                {
                    while (!threadGetFile.IsAlive) ; // Wait for thread to get started
                    Thread.Sleep(1);  // Wait a sec more
                    threadGetFile.Join();    // Wait for thread to end

                    Ruta = oFichero.FileName;
                }
                catch (Exception ex)
                {
                    Matriz.oGlobal.SBOApp.MessageBox(ex.Message, 1, "OK", "", "");
                }
                threadGetFile = null;
                oFichero = null;

                return Ruta;
            }

       public static string SeleccionoFichero()
            {
                string cRetorno = "";

                EXO_OpenFileDialog OpenFileDialog = new EXO_OpenFileDialog();
                OpenFileDialog.Filter = "Todos los ficheros|*.*";
                OpenFileDialog.InitialDirectory = "";
                Thread threadGetFile = new Thread(new ThreadStart(OpenFileDialog.GetFileName));
                threadGetFile.TrySetApartmentState(ApartmentState.STA);
                try
                {
                    threadGetFile.Start();
                    while (!threadGetFile.IsAlive) ; // Wait for thread to get started
                    Thread.Sleep(1);  // Wait a sec more
                    threadGetFile.Join();    // Wait for thread to end

                    // Use file name as you will here
                    cRetorno = OpenFileDialog.FileName;
                    threadGetFile.Abort();
                    threadGetFile = null;
                    OpenFileDialog.InitialDirectory = "";
                    OpenFileDialog = null;
                }
                catch (Exception ex)
                {
                    Matriz.oGlobal.SBOApp.MessageBox(ex.Message, 1, "OK", "", "");
                    threadGetFile.Abort();
                    threadGetFile = null;
                    OpenFileDialog.InitialDirectory = "";
                    OpenFileDialog = null;

                }

                return cRetorno;
            }
    }

    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        private IntPtr _hwnd;

        // Property
        public virtual IntPtr Handle
        {
            get { return _hwnd; }
        }

        // Constructor
        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }
    }

    public class Utilidades
    {        

        #region Lleno Combo de Series
        public static bool LlenoComboSeries(ref SAPbouiCOM.ComboBox oCombo,  BoObjectTypes oTipoObjeto, bool lSinBloqueados, bool lConIndicador, string cFecha)
        {
            string cNumeroObjeto = "", cIndicador = "", sql = "";
            SAPbobsCOM.Recordset oRec;

            switch (oTipoObjeto)
            {
                case BoObjectTypes.oDeliveryNotes:
                      cNumeroObjeto = "15";
                    break;
                case BoObjectTypes.oInvoices:
                      cNumeroObjeto = "13";
                    break;

            }

            if (lConIndicador)
            {
                sql = "SELECT T0.Indicator FROM OFPR T0 WHERE Convert(varchar(10), T0.F_RefDate, 112) <= '" + cFecha + "'  AND ";
                sql += " Convert(varchar(10), T0.T_RefDate, 112) >= '" + cFecha + "'";
                oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
                cIndicador = oRec.Fields.Item(0).Value.ToString().Trim();
            }
           
            sql = "SELECT T0.Series, T0.SeriesName FROM NNM1 T0  ";
            sql += " WHERE T0.ObjectCode = '" + cNumeroObjeto + "'";

            if (lConIndicador) sql += " AND T0.Indicator = '" + cIndicador + "'";
            if (lSinBloqueados) sql += " AND T0.Locked = 'N'";



            //Borro lo que hubiera
            int nCount = oCombo.ValidValues.Count;
            for (int i = 0; i < nCount; i++)
            {
                oCombo.ValidValues.Remove(0, BoSearchKey.psk_Index);
            }
            
            oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
            while (!oRec.EoF)
            {
                oCombo.ValidValues.Add(oRec.Fields.Item(0).Value.ToString().Trim(), oRec.Fields.Item(1).Value.ToString().Trim());
                oRec.MoveNext();
            }
            

            return true;
        }

        public static bool LlenoComboSeries(ref SAPbouiCOM.Column oColumnCombo, BoObjectTypes oTipoObjeto, bool lSinBloqueados, bool lConIndicador, string cFecha)
        {
            SAPbobsCOM.Recordset oRec;
            oRec = SubLlenoComboSeries(oTipoObjeto, lSinBloqueados, lConIndicador, cFecha);

            //Borro lo que hubiera
            int nCount = oColumnCombo.ValidValues.Count;
            for (int i = 0; i < nCount; i++)
            {
                oColumnCombo.ValidValues.Remove(0, BoSearchKey.psk_Index);
            }

            while (!oRec.EoF)
            {
                oColumnCombo.ValidValues.Add(oRec.Fields.Item(0).Value.ToString().Trim(), oRec.Fields.Item(1).Value.ToString().Trim());
                oRec.MoveNext();
            }


            return true;
        }

        private static SAPbobsCOM.Recordset SubLlenoComboSeries(BoObjectTypes oTipoObjeto, bool lSinBloqueados, bool lConIndicador, string cFecha)
        {
            string cNumeroObjeto = "", cIndicador = "", sql = "";
            SAPbobsCOM.Recordset oRec = null;

            switch (oTipoObjeto)
            {
                //Albaran de ventas
                case BoObjectTypes.oDeliveryNotes:
                    cNumeroObjeto = "15";
                    break;

                //Facturas de venta
                case BoObjectTypes.oInvoices:
                    cNumeroObjeto = "13";
                    break;

                //Devoluciones de ventas
                case BoObjectTypes.oReturns:
                    cNumeroObjeto = "16";
                    break;
                    
                //Abonos
                case BoObjectTypes.oCreditNotes:
                    cNumeroObjeto = "14";
                    break;

            }

            if (lConIndicador)
            {
                sql = "SELECT T0.Indicator FROM OFPR T0 WHERE Convert(varchar(10), T0.F_RefDate, 112) <= '" + cFecha + "'  AND ";
                sql += " Convert(varchar(10), T0.T_RefDate, 112) >= '" + cFecha + "'";
                oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
                cIndicador = oRec.Fields.Item(0).Value.ToString().Trim();
            }

            sql = "SELECT T0.Series, T0.SeriesName FROM NNM1 T0  ";
            sql += " WHERE T0.ObjectCode = '" + cNumeroObjeto + "'";

            if (lConIndicador) sql += " AND T0.Indicator = '" + cIndicador + "'";
            if (lSinBloqueados) sql += " AND T0.Locked = 'N'";

            oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
            return oRec;

            ////Borro lo que hubiera
            //int nCount = oCombo.ValidValues.Count;
            //for (int i = 0; i < nCount; i++)
            //{
            //    oCombo.ValidValues.Remove(0, BoSearchKey.psk_Index);
            //}

            //oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
            //while (!oRec.EoF)
            //{
            //    oCombo.ValidValues.Add(oRec.Fields.Item(0).Value.ToString().Trim(), oRec.Fields.Item(1).Value.ToString().Trim());
            //    oRec.MoveNext();
            //}


            
        }        
        #endregion

        public static int CategoriaQuery(string cNomCategoria, bool lCrear = false)
        {
            int nRetorno = 0;
        
            string strSQL = "SELECT CategoryId FROM OQCN WHERE CatName = '" + cNomCategoria + "'";
            nRetorno = (int) Matriz.oGlobal.refDi.SQL.sqlNumericaB1(strSQL);
            if (nRetorno == 0 && lCrear )
            {
                SAPbobsCOM.QueryCategories oCategorias = (SAPbobsCOM.QueryCategories) Matriz.oGlobal.compañia.GetBusinessObject(BoObjectTypes.oQueryCategories);
                oCategorias.Name = cNomCategoria;

                if( oCategorias.Add() != 0)
                {
                    Matriz.oGlobal.SBOApp.MessageBox("ERROR !!" + Matriz.oGlobal.compañia.GetLastErrorDescription(), 1, "Ok", "", "");
                }
                else
                {
                    Matriz.oGlobal.SBOApp.SetStatusBarMessage("Categoría Query " + cNomCategoria + " creada ", BoMessageTime.bmt_Short, false);
                    nRetorno = Convert.ToInt32(Matriz.oGlobal.compañia.GetNewObjectKey());
                }

                Object Ob = (Object) oCategorias;
                EXO_CleanCOM.CLiberaCOM.liberaCOM(ref Ob);
            }

            return nRetorno;        
        }

        public static int VerificoUserQuery(string cNomQuery, string cStringQuery, int nCategoriaID, bool lCrear = false)
        {                        
            string sql = "SELECT IntrnalKey FROM OUQR WHERE QName = '" + cNomQuery + "'";

            int nRetorno = Convert.ToInt32(Matriz.oGlobal.refDi.SQL.sqlNumericaB1(sql));

            if (nRetorno == 0 && lCrear)
            {
                SAPbobsCOM.QueryCategories oCategoria = Matriz.oGlobal.compañia.GetBusinessObject(BoObjectTypes.oQueryCategories);
                SAPbobsCOM.UserQueries oUserQueries = Matriz.oGlobal.compañia.GetBusinessObject(BoObjectTypes.oUserQueries);

                oUserQueries.QueryCategory = nCategoriaID;
                oUserQueries.QueryDescription = cNomQuery; 
                oUserQueries.Query = cStringQuery;
                

                if ( oUserQueries.Add() != 0)
                {
                    Matriz.oGlobal.SBOApp.MessageBox("ERROR !!" + Matriz.oGlobal.compañia.GetLastErrorDescription(), 1, "Ok", "", "");
                }
                else
                {
                    Matriz.oGlobal.SBOApp.SetStatusBarMessage("Query " + cNomQuery + " creada ", BoMessageTime.bmt_Short, false);
                    nRetorno = Convert.ToInt32(Matriz.oGlobal.refDi.SQL.sqlNumericaB1(sql));
                }

                Object Ob = (Object)oUserQueries;
                EXO_CleanCOM.CLiberaCOM.liberaCOM(ref Ob);

                Ob = (Object)oCategoria;
                EXO_CleanCOM.CLiberaCOM.liberaCOM(ref Ob);
            }

            return nRetorno;       
        }

        public static void CrearBusquedaFormateada(string FormId, string ItemId, string ColId, BoFormattedSearchActionEnum Accion,
                                                   int Consulta, BoYesNoEnum Refrescar, string FieldId, BoYesNoEnum ForzarRefrescar, BoYesNoEnum PorCampo)
        {
            bool lExiste = false;
            int nRet;

            SAPbobsCOM.FormattedSearches oFormattedSearches = (SAPbobsCOM.FormattedSearches)Matriz.oGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);

            string sql = "SELECT IndexID FROM CSHS WHERE FormId = '" + FormId + "' AND ItemID = '" + ItemId + "' AND ColID = '" + ColId + "'";
            SAPbobsCOM.Recordset oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
            string cAux = oRec.Fields.Item(0).Value.ToString().Trim();
            if (cAux == "") cAux = "0";
                        
            if (Convert.ToInt16(cAux) != 0)
            {
                
                oFormattedSearches.GetByKey(Convert.ToInt32(Convert.ToInt16(cAux)));
                lExiste = true;
            }

            oFormattedSearches.FormID = FormId;
            oFormattedSearches.ItemID = ItemId;
            oFormattedSearches.ColumnID = ColId;
            oFormattedSearches.Action = Accion;
            oFormattedSearches.QueryID = Consulta;
            oFormattedSearches.Refresh = Refrescar;
            oFormattedSearches.FieldID = FieldId;
            oFormattedSearches.ForceRefresh = ForzarRefrescar;
            oFormattedSearches.ByField = PorCampo;

            nRet = lExiste ? oFormattedSearches.Update() : oFormattedSearches.Add();
            if (nRet != 0)
            {
                Matriz.oGlobal.SBOApp.MessageBox(Matriz.oGlobal.compañia.GetLastErrorDescription(), 1, "", "", "");
            }
            else
            {
                Matriz.oGlobal.SBOApp.SetStatusBarMessage((lExiste ? "Actualizada" : "Creada") + " busqueda formateada para Form. " + FormId + " Item " + ItemId + " Col " + ColId + " - consulta " + Consulta.ToString(), BoMessageTime.bmt_Short, false);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oFormattedSearches);
            oFormattedSearches = null;
        }

        public static void BorroLineaMatrix(ref SAPbouiCOM.Matrix oMatrix, ref SAPbouiCOM.Form oFormulario)
        {

            oFormulario.Freeze(true);
            try
            {
                oMatrix.FlushToDataSource();
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    if (oMatrix.IsRowSelected(i))
                    {
                        oMatrix.DeleteRow(i);
                        if (oFormulario.Mode == BoFormMode.fm_OK_MODE) oFormulario.Mode = BoFormMode.fm_UPDATE_MODE;
                        break;
                    }
                }

                oMatrix.FlushToDataSource();
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                Matriz.oGlobal.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
            }

            oFormulario.Freeze(false);
        }

        public static bool GraboDeDataSource(SAPbouiCOM.DBDataSource oDBDataSource, string cTabla, int Digitos = 0)
        {
            string cCode;
            int nOK;
            DateTime dAuxiliar;
            SAPbobsCOM.UserTable oUserTable = Matriz.oGlobal.compañia.UserTables.Item(cTabla);

            try
            {
                for (int i = 0; i <= oDBDataSource.Size - 1; i++)
                {
                    cCode = oDBDataSource.GetValue("Code", i).Trim();

                    bool lCrearNuevo = true;
                    if (cCode != "")
                    {
                        lCrearNuevo = !oUserTable.GetByKey(cCode);
                    }

                    if (!lCrearNuevo)
                    {
                        oUserTable.GetByKey(cCode);
                        #region si existe...

                        foreach (SAPbouiCOM.Field oField in oDBDataSource.Fields)
                        {
                            if (oField.Name != "Code" && oField.Name != "Name")
                            {
                                switch (oField.Type)
                                {
                                    case (BoFieldsType.ft_Date):
                                        {
                                            if (DateTime.TryParseExact(oDBDataSource.GetValue(oField.Name, i), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dAuxiliar))
                                            {
                                                oUserTable.UserFields.Fields.Item(oField.Name).Value = dAuxiliar;
                                            }
                                            else
                                            {
                                                oUserTable.UserFields.Fields.Item(oField.Name).Value = "";
                                            }
                                        }
                                        break;
                                    case BoFieldsType.ft_Float:
                                    case BoFieldsType.ft_Percent:
                                    case BoFieldsType.ft_Measure:
                                    case BoFieldsType.ft_Price:
                                    case BoFieldsType.ft_Quantity:
                                    case BoFieldsType.ft_Rate:
                                    case BoFieldsType.ft_Sum:
                                        oUserTable.UserFields.Fields.Item(oField.Name).Value = Conversiones.ValueSAPToDoubleSistema(oDBDataSource.GetValue(oField.Name, i));
                                        break;

                                    default:
                                        {
                                            oUserTable.UserFields.Fields.Item(oField.Name).Value = oDBDataSource.GetValue(oField.Name, i);
                                        }
                                        break;
                                }
                            }
                        }

                        nOK = oUserTable.Update();
                        if (nOK != 0)
                        {
                            Matriz.oGlobal.SBOApp.SetStatusBarMessage(Matriz.oGlobal.compañia.GetLastErrorDescription(), BoMessageTime.bmt_Short, true);
                            return false;
                        }
                        #endregion
                    }
                    else
                    {
                        #region si no existe
                        oUserTable = Matriz.oGlobal.compañia.UserTables.Item(cTabla);

                        foreach (SAPbouiCOM.Field oField in oDBDataSource.Fields)
                        {
                            if (oField.Name != "Code" && oField.Name != "Name")
                            {
                                switch (oField.Type)
                                {
                                    case (BoFieldsType.ft_Date):
                                        {
                                            if (DateTime.TryParseExact(oDBDataSource.GetValue(oField.Name, i), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dAuxiliar))
                                            {
                                                oUserTable.UserFields.Fields.Item(oField.Name).Value = dAuxiliar;
                                            }
                                            else
                                            {
                                                oUserTable.UserFields.Fields.Item(oField.Name).Value = "";
                                            }
                                        }
                                        break;

                                    case BoFieldsType.ft_Float:
                                    case BoFieldsType.ft_Percent:
                                    case BoFieldsType.ft_Measure:
                                    case BoFieldsType.ft_Price:
                                    case BoFieldsType.ft_Quantity:
                                    case BoFieldsType.ft_Rate:
                                    case BoFieldsType.ft_Sum:
                                        oUserTable.UserFields.Fields.Item(oField.Name).Value = Conversiones.ValueSAPToDoubleSistema(oDBDataSource.GetValue(oField.Name, i));
                                        break;

                                    default:
                                        oUserTable.UserFields.Fields.Item(oField.Name).Value = oDBDataSource.GetValue(oField.Name, i);
                                        break;
                                }
                            }
                            else if (oField.Name == "Code")
                            {
                                #region Ultimo code
                                string sqlUlt = "SELECT MAX( CAST(T0.Code AS NUMERIC(10))) FROM [@" + cTabla + "] T0";
                                double nMaxCode = Matriz.oGlobal.refDi.SQL.sqlNumericaB1(sqlUlt);
                                string cNuevoCode = "";
                                if (Digitos == -1)
                                {
                                    cNuevoCode = Convert.ToString((Convert.ToInt32(nMaxCode) + 1));
                                }
                                else if (Digitos == 0)
                                {
                                    string cNumCeros = Convert.ToString(nMaxCode);
                                    string cFormatAux = "".PadRight(cNumCeros.Length, '0');
                                    cNuevoCode = (Convert.ToInt32(nMaxCode) + 1).ToString(cFormatAux);
                                }
                                else
                                {
                                    string cFormatAux = "".PadRight(Digitos, '0');
                                    cNuevoCode = (Convert.ToInt32(nMaxCode) + 1).ToString(cFormatAux);
                                }
                                #endregion
                                oUserTable.Code = cNuevoCode;
                                oUserTable.Name = oUserTable.Code;
                            }
                        }

                        nOK = oUserTable.Add();
                        if (nOK != 0)
                        {
                            Matriz.oGlobal.SBOApp.SetStatusBarMessage(Matriz.oGlobal.compañia.GetLastErrorDescription(), BoMessageTime.bmt_Short, true);
                            return false;
                        }
                        #endregion
                    }
                }
            }
            catch (Exception ex)
            {
                Matriz.oGlobal.SBOApp.SetStatusBarMessage(Matriz.oGlobal.compañia.GetLastErrorDescription(), BoMessageTime.bmt_Short, true);
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return true;

        }

        public static void BorroDataTable(ref SAPbouiCOM.DataTable oTablaInf)
        {            
            if (!oTablaInf.IsEmpty)
            {
              int nNumReg = oTablaInf.Rows.Count;
              for (int i = 0; i < nNumReg; i++)
              {
                  oTablaInf.Rows.Remove(0);
              }
            }            
        }

        public static void LLenoComboGenerico(ref SAPbouiCOM.Item oItemCombo, string cTabla, string cWhere = "")
        {
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox) oItemCombo.Specific;
            string sql = "SELECT T0.Code, T0.Name FROM [" + cTabla + "] T0 "  + cWhere + " ORDER BY T0.Name ";
            SAPbobsCOM.Recordset oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
            while (!oRec.EoF)
            {
                oCombo.ValidValues.Add(oRec.Fields.Item(0).Value, oRec.Fields.Item(1).Value);
                oRec.MoveNext();
            }
            oItemCombo.DisplayDesc = true;
            oCombo.ExpandType = BoExpandType.et_DescriptionOnly;
        }

        public static void LLenoComboGenerico(ref SAPbouiCOM.Column  oColumCombo, string cTabla)
        {            
            string sql = "SELECT T0.Code, T0.Name FROM [" + cTabla + "] T0 ORDER BY T0.Name ";
            SAPbobsCOM.Recordset oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1(sql);
            while (!oRec.EoF)
            {
                oColumCombo.ValidValues.Add(oRec.Fields.Item(0).Value, oRec.Fields.Item(1).Value);
                oRec.MoveNext();
            }
            oColumCombo.DisplayDesc = true;
            oColumCombo.ExpandType = BoExpandType.et_DescriptionOnly;
        }

        public static SAPbouiCOM.Form BuscoFormLanzado(string TypeEx)
        {
            SAPbouiCOM.Form oFORMORET = null;
            for (int i = 0; i < Matriz.oGlobal.SBOApp.Forms.Count; i++)
            {
                if (Matriz.oGlobal.SBOApp.Forms.Item(i).TypeEx == TypeEx)
                {
                    oFORMORET = Matriz.oGlobal.SBOApp.Forms.GetForm(Matriz.oGlobal.SBOApp.Forms.Item(i).TypeEx, Matriz.oGlobal.SBOApp.Forms.Item(i).TypeCount);                                        
                    break;
                }
            }

            return oFORMORET;
        }

        public static bool LanzoMenuUserTable(string cTablaSinArroba)
        {
           bool lRetorno = false;
           SAPbouiCOM.Menus oMenus = Matriz.oGlobal.SBOApp.Menus.Item("51200").SubMenus;
           for (int i = 0; i <= oMenus.Count - 1; i++)
           {
               if (oMenus.Item(i).String.IndexOf(cTablaSinArroba) == 0)
              {
                 Matriz.oGlobal.SBOApp.ActivateMenuItem(oMenus.Item(i).UID);
                 lRetorno = true;
                 break;                 
               }
            }


           EXO_CleanCOM.CLiberaCOM.Menus(oMenus);
           return lRetorno;
        }

        public static string LeoQueryFich(string cNomFichLargo)
        {
            string sql = "", cAux = "";                           
            System.IO.StreamReader Fichero = new System.IO.StreamReader(cNomFichLargo);
            while (Fichero.Peek() != -1)
            {
              cAux = Fichero.ReadLine();
              if (cAux.Length > 2 && cAux.Substring(0, 2) == "--") continue;

              sql += cAux.Replace("\t", " ") + " ";
            }
            Fichero.Close();
                        
            return sql;
        }

        public static string LeoQueryFich(string cNomQueryIncrustada, Type Tipo)

        {             
             string cQuery = Matriz.oGlobal.funciones.leerEmbebido(Tipo, cNomQueryIncrustada);
             cQuery = cQuery.Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");

             return cQuery;
        }

        public static void LlenoProvincias(ref SAPbouiCOM.Item oItem, string cPais, BoExpandType TipoExpan = BoExpandType.et_DescriptionOnly)
        {
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox) oItem.Specific;
            #region Borro las provincias de antes
            int nNumVal = oCombo.ValidValues.Count;
            for (int j = 0; j < nNumVal; j++)
            {
                oCombo.ValidValues.Remove(0, BoSearchKey.psk_Index);
            }
            #endregion
            
            SAPbobsCOM.Recordset oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1("SELECT T0.Code, T0.Name FROM OCST T0 WHERE T0.Country = '" + cPais + "' ORDER BY T0.Name");
            while (!oRec.EoF)
            {
                oCombo.ValidValues.Add(oRec.Fields.Item(0).Value, oRec.Fields.Item(1).Value);
                oRec.MoveNext();
            }
            oCombo.ExpandType = TipoExpan;
            
            Object ob = (Object)oRec;
            EXO_CleanCOM.CLiberaCOM.liberaCOM(ref ob);

        }

        public static void LLenoComboGrupoArt(ref SAPbouiCOM.Item oItem, string cWhere = "")
        {
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
                
            SAPbobsCOM.Recordset oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1("SELECT T0.ItmsGrpCod, T0.ItmsGrpNam FROM OITB T0 " + cWhere + " ORDER BY T0.ItmsGrpNam");
            while (!oRec.EoF)
            {
                oCombo.ValidValues.Add(oRec.Fields.Item(0).Value, oRec.Fields.Item(1).Value);
                oRec.MoveNext();
            }
            oCombo.ExpandType = BoExpandType.et_DescriptionOnly;

            Object ob = (Object)oRec;
            EXO_CleanCOM.CLiberaCOM.liberaCOM(ref ob);

        }

        public static void LLenoComboGrupoArt(ref SAPbouiCOM.Column oColumna, string cWhere = "")
        {
            SAPbobsCOM.Recordset oRec = Matriz.oGlobal.refDi.SQL.sqlComoRsB1("SELECT T0.ItmsGrpCod, T0.ItmsGrpNam FROM OITB T0 " + cWhere + " ORDER BY T0.ItmsGrpNam");
            while (!oRec.EoF)
            {
                oColumna.ValidValues.Add(oRec.Fields.Item(0).Value, oRec.Fields.Item(1).Value);
                oRec.MoveNext();
            }
            oColumna.ExpandType = BoExpandType.et_DescriptionOnly;

            Object ob = (Object)oRec;
            EXO_CleanCOM.CLiberaCOM.liberaCOM(ref ob);
        }

        public static void ActualizoFormUDO(string cUDO, string cFicheroSRF)
        {            
            string cStringXml = Matriz.oGlobal.funciones.leerEmbebido(Matriz.TypeMatriz, cFicheroSRF);
            if (cStringXml != "")
            {
                GC.Collect();
                SAPbobsCOM.UserObjectsMD oUserObjectMD = oUserObjectMD = (SAPbobsCOM.UserObjectsMD) Matriz.oGlobal.compañia.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
                oUserObjectMD.GetByKey(cUDO);
            
                oUserObjectMD.EnableEnhancedForm = BoYesNoEnum.tYES;
                oUserObjectMD.RebuildEnhancedForm = BoYesNoEnum.tNO;
                oUserObjectMD.FormSRF = cStringXml;
                int lRetCode = oUserObjectMD.Update();

                if (lRetCode != 0)
                {
                    Matriz.oGlobal.SBOApp.MessageBox("ERROR!!\n" + Matriz.oGlobal.compañia.GetLastErrorDescription(), 1, "Ok", "", "");                                        
                }
                else
                {
                    Matriz.oGlobal.SBOApp.SetStatusBarMessage("Actualizado Form UDO " + cUDO);                   
                }                
            }           
        }
        
        public static double RecojoPrecio(string cIC, string cItemCode, double nCantidad, DateTime dFecha)       
        {
            double nRetorno = 0;

            //Precio
            SAPbobsCOM.SBObob sboBob = (SAPbobsCOM.SBObob)Matriz.oGlobal.compañia.GetBusinessObject(BoObjectTypes.BoBridge);
            SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)Matriz.oGlobal.compañia.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRec = sboBob.GetItemPrice(cIC, cItemCode, nCantidad, dFecha);
            nRetorno = (double) oRec.Fields.Item(0).Value;


            Object oB = (Object) oRec;
            EXO_CleanCOM.CLiberaCOM.liberaCOM(ref oB);

            return nRetorno;            
        }

        public static string DameValorFUNDI(string StrTabla, string StrCampo, string StrCondicion)
        {
            SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)Matriz.oGlobal.compañia.GetBusinessObject(BoObjectTypes.BoRecordset));

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

                oRecordset.DoQuery(CadenaSelect);
                oRecordset.MoveFirst();
                if (oRecordset.RecordCount <= 0)
                {
                    Object oB = (Object)oRecordset;
                    EXO_CleanCOM.CLiberaCOM.liberaCOM(ref oB);
                    return "";
                }
                else
                {
                    if (oRecordset.Fields.Item(0).Value.ToString() != "")
                    {
                        string a = oRecordset.Fields.Item(0).Value.ToString();
                        string cRetorno = oRecordset.Fields.Item(0).Value.ToString();

                        Object oB = (Object)oRecordset;
                        EXO_CleanCOM.CLiberaCOM.liberaCOM(ref oB);

                        return cRetorno;
                    }
                    else
                    {
                        Object oB = (Object)oRecordset;
                        EXO_CleanCOM.CLiberaCOM.liberaCOM(ref oB);

                        return "";
                    }
                }
            }
            catch
            {
                Object oB = (Object)oRecordset;
                EXO_CleanCOM.CLiberaCOM.liberaCOM(ref oB);
                return "";
            }
            
            
        }

        public static double ConvertirCantidadFUNDI(string Valor)
        {
            string SeparadorDecimal = DameValorFUNDI("OADM", "DecSep", "");
            string SeparadorMiles = DameValorFUNDI("OADM", "ThousSep", "");

            Valor = Valor.Replace(SeparadorMiles, "");

            if (SeparadorDecimal != ",")
            {
                return Convert.ToDouble(Valor.Replace(SeparadorDecimal, ","));
            }
            return Convert.ToDouble(Valor);
        }

        public static bool TodasMayusculasFUNDI(string Cadena)
        {
            string Estructura = "[a-z,ñ]";
            Regex Estructura_Regex = new Regex(Estructura);

            if (Estructura_Regex.IsMatch(Cadena))
            {
                return false;
            }
            return true;
        }

        public static bool Informe(string pReport, string pSeleccion)
        {
            //try
            //{
            //    if (Matriz.crReport != null && Matriz.crReport.IsLoaded)
            //    {
            //        Matriz.crReport.Close();
            //        GC.Collect();
            //    }

            //    Matriz.crReport = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            //    Matriz.crReport.Load(pReport);

            //    string cPassDB = Matriz.oGlobal.refDi.SQL.sqlStringB1("select T0.U_EXO_BDPW FROM [@EXO_OGEN] T0");
            //    CrystalDecisions.Shared.DataSourceConnections conrepor = Matriz.crReport.DataSourceConnections;
            //    conrepor[0].SetConnection(Matriz.oGlobal.compañia.Server, Matriz.oGlobal.compañia.CompanyDB, Matriz.oGlobal.compañia.DbUserName, cPassDB);

            //    if (pSeleccion != "")
            //    {
            //        Matriz.crReport.RecordSelectionFormula = pSeleccion;
            //    }

                
            //    frmReportViewer formReportViewer = new frmReportViewer(false, ref Matriz.ThisMatriz);
            //    Matriz.LanzarImpresionCrystal = true;
            //    return true;
            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show(ex.Message);
                return false;
            //}
        }

        public static string DameValorEditTextFUNDI(string Item, SAPbouiCOM.Form Formulario)
        {
            //oForm = Formulario;
            SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)Formulario.Items.Item(Item).Specific;
            return oEditText.String;
        }

        public static double TextoADoubleFUNDI(string StrTexto)
        {
            string StrAux;
            if (!IsNumericFUNDI(StrTexto))
            {
                return 0;
            }
            StrAux = StrTexto.Replace(".", ",");
            double a = Convert.ToDouble(StrAux);
            return a; // Convert.ToDouble(StrAux);
        }

        public static bool IsNumericFUNDI(object Expression)
        {
            bool isNum;
            double retNum;

            isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }

        public static string UltimoCodeFUNDI(string StrTabla)
        {

            SAPbobsCOM.Recordset oRecordset = ((SAPbobsCOM.Recordset)Matriz.oGlobal.compañia.GetBusinessObject(BoObjectTypes.BoRecordset));
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

            Object oB = (Object)oRecordset;
            EXO_CleanCOM.CLiberaCOM.liberaCOM(ref oB);
        }

        public static string CompletaConCerosFUNDI(int NumeroCaracteres, string Valor, int AumentarContadorEn)
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


        public static string NuevoReportType(string cTypeName, string cAddonName, string cTipoForm, string cMenuID, bool lCrear = false)
        {
            SAPbobsCOM.ReportTypesService rptTypeService;
            rptTypeService = (SAPbobsCOM.ReportTypesService)Matriz.oGlobal.compañia.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService);
            SAPbobsCOM.ReportType newType = (SAPbobsCOM.ReportType)rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType);

            SAPbobsCOM.ReportTypesParams oReporParams;
            oReporParams = rptTypeService.GetReportTypeList();

            string cRetorno = "";

            //miro a ver si existe, y si es asi, lo fijo
            for (int i = 0; i < oReporParams.Count; i++)
            {
                if (oReporParams.Item(i).TypeName == cTypeName &&
                    oReporParams.Item(i).AddonName == cAddonName &&
                    oReporParams.Item(i).MenuID == cMenuID &&
                    oReporParams.Item(i).AddonFormType == cTipoForm
                    )
                {
                    cRetorno = oReporParams.Item(i).TypeCode;
                    return cRetorno;
                }
            }

            if (lCrear)
            {
                //Si llega aqui, creo el el nuevo tipo
                rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType);
                newType.TypeName = cTypeName;
                newType.AddonName = cAddonName;
                newType.AddonFormType = cTipoForm;
                newType.MenuID = cMenuID;
                SAPbobsCOM.ReportTypeParams newTypeParam = rptTypeService.AddReportType(newType);

                //para ahora fijar el report en el impreso
                cRetorno = newTypeParam.TypeCode;
            }
            return cRetorno;

            ////Creo la layout
            //SAPbobsCOM.ReportLayoutsService rptService = (SAPbobsCOM.ReportLayoutsService)
            //csVariablesGlobales.oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);
            //SAPbobsCOM.ReportLayout newReport = (SAPbobsCOM.ReportLayout)
            //rptService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout);
            //newReport.Author = csVariablesGlobales.oCompany.UserName;
            //newReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal;
            //newReport.Name = "Addon Demo Report 6";
            //newReport.TypeCode = newTypeParam.TypeCode;
            //SAPbobsCOM.ReportLayoutParams newReportParam = rptService.AddReportLayout(newReport);

            ////la pongo por defecto
            //newType = rptTypeService.GetReportType(newTypeParam);
            //newType.DefaultReportLayout = newReportParam.LayoutCode;
            //rptTypeService.UpdateReportType(newType);

        }



        


    }

}

