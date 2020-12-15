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
    class csModelos
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

        public static void GenerarFichero347(string Ruta, SAPbouiCOM.Form Formulario)
        {
            StreamWriter Fichero347 = new StreamWriter(Ruta);
            try
            {
                oForm = Formulario;
                string cNIFOK = "", cAuxiliar = "", cFecha = "", StrNifEmp, StrNomEmp, StrTelEmp, StrLinea;
                string cSiguiente = "", cNumFac, StrNifCli = "", StrNomCli = "", StrCpCli = "", StrCodCli = "", StrTipoIC = "", Suma = "";
                string cSignoSuma, cSignoPriTri, cSignoSegunTri, cSignoTer, cSignoCuar;
                string cTabla = "";
                int nMes = 0;
                double PrimeTri = 0, SegunTri = 0, TercerTri = 0, CuartoTri = 0;
                oRecordset = ((SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));

                StrNifEmp = csUtilidades.DameValor("OADM", "TaxIdNum", "").ToUpper();
                StrNifEmp = StrNifEmp.Substring(2, StrNifEmp.Length - 2);// UCase(Mid(StrNifEmp, 3, Len(StrNifEmp)));
                StrNomEmp = csUtilidades.DameValor("OADM", "CompnyName", "").ToUpper();
                StrTelEmp = csUtilidades.DameValor("OADM", "Phone1", "").ToUpper();
                StrLinea = "1" +
                            "347" + //Mid(StrLinea, 2, 3) = "347";
                            csVariablesGlobales.AñoModelo + //Mid(StrLinea, 5, 4) = StrEjercicio;
                            StrNifEmp +//Mid(StrLinea, 9, 9) = StrNifEmp; // /*NIF EMPRESA*/
                            StrNomEmp.PadRight(40).Substring(0, 40) +// Mid(StrLinea, 18, 40) = Mid(StrNomEmp, 1, 40); // /*NOMBRE EMPRESA*/
                            "T" +// Mid(StrLinea, 58, 1) = "D";
                            StrTelEmp.PadRight(9, '0').Substring(0, 9) +// Mid(StrLinea, 59, 9) = Mid(StrTelfEmp, 1, 9); // /*TELEFONO EMPRESA*/
                            StrNomEmp.PadRight(40).Substring(0, 40) +// Mid(StrLinea, 68, 40) = Mid(StrNomEmp, 1, 40); // /*NOMBRE EMPRESA*/
                            "3470000000001" + //Mid(StrLinea, 108, 13) = "3470000000001"; // /*EURO*/
                            "  " +//Mid(StrLinea, 121, 2) = "  ";
                            "0".PadLeft(13, '0') + //Nº identificativo de la declaracion
                            "0".PadLeft(9, '0') + //Nº total de personas y declarantes
                            " " + "0".PadLeft(15, '0') + //Importe total operaciones
                            "0".PadLeft(9, '0') + //Nº total inmuebles
                            "0".PadLeft(15, '0') + //Importe total 
                            " ".PadRight(316);
                Fichero347.WriteLine(StrLinea); //PrintLine(1, StrLinea);
                oForm = Formulario;
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                {                    
                    StrNifCli = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(i).Specific).String;
                    cNumFac = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("350000012").Cells.Item(i).Specific).String;

                    //Si esta en blanco o no hay nada, continue                      
                    if ( (StrNifCli == "" && cNumFac == "") || (cNumFac != "" && cNIFOK == "") ) continue;

                   
                    if (StrNifCli != "")
                    {
                        cNIFOK = StrNifCli;
                        StrCodCli = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("0").Cells.Item(i).Specific).String;                        
                        if (StrNifCli.Length > 2)
                        {
                            StrNifCli = StrNifCli.Substring(2, StrNifCli.Length - 2);
                        }
                        StrNomCli = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).String;
                        StrNomCli = StrNomCli.ToUpper().Replace("ª", " ").Replace("º", " ").PadRight(40).Substring(0, 40);

                        StrCpCli = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("5").Cells.Item(i).Specific).String;
                        if (StrCpCli.Length != 0)
                        {
                            StrCpCli = StrCpCli.ToUpper().Substring(0, 2);
                        }
                        else
                        {
                            StrCpCli = "  ";
                            csVariablesGlobales.SboApp.MessageBox("Falta el C.P. del I.C. " + StrCodCli, 1, "", "", "");
                        }

                        //StrTipoIC = csUtilidades.DameValor("OCRD", "CardType", "CardCode = '" + StrCodCli + "'");                        
                        oRecordset.DoQuery("SELECT CardType FROM OCRD WHERE OCRD.CardCode = '" + StrCodCli + "'");
                        oRecordset.MoveFirst();
                        StrTipoIC = oRecordset.Fields.Item("CardType").Value.ToString();                                                

                        if (StrTipoIC == "C")
                        {
                            StrTipoIC = "B"; //Cliente
                        }
                        else
                        {
                            StrTipoIC = "A"; //Proveedor
                        }
                        //Aqui esta la suma final
                        Suma = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("4").Cells.Item(i).Specific).String;
                    }

                    if (cNumFac != "")
                    {                        
                        cTabla = "";
                        if ( cNumFac.Replace("FA", "") != cNumFac )
                        {
                            cAuxiliar = cNumFac.Replace("FA", "");
                            cTabla = "OINV";
                        }
                        else if (cNumFac.Replace("RC", "") != cNumFac)
                        {
                            cAuxiliar = cNumFac.Replace("RC", "");
                            cTabla = "ORIN";
                        }
                        else if (cNumFac.Replace("AN", "") != cNumFac)
                        {
                            cAuxiliar = cNumFac.Replace("AN", "");
                            if ( StrTipoIC == "C" )
                            {
                                cTabla = "ODPI";
                            }
                            else
                            {
                                cTabla = "ODPO";
                            }                                                     
                        }
                        else if (cNumFac.Replace("TT", "") != cNumFac)
                        {
                            cAuxiliar = cNumFac.Replace("TT", "");
                            cTabla = "OPCH";
                        }
                        else if (cNumFac.Replace("AC", "") != cNumFac)
                        {
                            cAuxiliar = cNumFac.Replace("AC", "");
                            cTabla = "ORPC";
                        }
                        else
                        {
                            csVariablesGlobales.SboApp.MessageBox("Error en 347 para documento" + cNumFac, 0, "", "", "");
                            cTabla = "";
                            return;
                        }
                        if (cAuxiliar == "100000198")
                        {
                            cFecha = "uu";
                        }

                        oRecordset.DoQuery("SELECT TaxDate FROM " + cTabla + " WHERE DocNum = " + cAuxiliar);
                        oRecordset.MoveFirst();
                        cFecha = oRecordset.Fields.Item("TaxDate").Value.ToString();

                        nMes = Convert.ToDateTime(cFecha).Month;
                        if (nMes == 1 || nMes == 2 || nMes == 3)
                        {
                            cAuxiliar = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("4").Cells.Item(i).Specific).String;
                            cAuxiliar = cAuxiliar.Replace(",", "#");
                            cAuxiliar = cAuxiliar.Replace(".", ",");
                            cAuxiliar = cAuxiliar.Replace("#", ".");
                            PrimeTri = PrimeTri + Convert.ToDouble(cAuxiliar);
                        }
                        else if (nMes == 4 || nMes == 5 || nMes == 6)
                        {
                            cAuxiliar = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("4").Cells.Item(i).Specific).String;
                            cAuxiliar = cAuxiliar.Replace(",", "#");
                            cAuxiliar = cAuxiliar.Replace(".", ",");
                            cAuxiliar = cAuxiliar.Replace("#", ".");
                            SegunTri = SegunTri + Convert.ToDouble(cAuxiliar);
                        }
                        else if (nMes == 7 || nMes == 8 || nMes == 9)
                        {
                            cAuxiliar = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("4").Cells.Item(i).Specific).String;
                            cAuxiliar = cAuxiliar.Replace(",", "#");
                            cAuxiliar = cAuxiliar.Replace(".", ",");
                            cAuxiliar = cAuxiliar.Replace("#", ".");
                            TercerTri = TercerTri + Convert.ToDouble(cAuxiliar);
                        }
                        else if (nMes == 10 || nMes == 11 || nMes == 12)
                        {
                            cAuxiliar = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("4").Cells.Item(i).Specific).String;
                            cAuxiliar = cAuxiliar.Replace(",", "#");
                            cAuxiliar = cAuxiliar.Replace(".", ",");
                            cAuxiliar = cAuxiliar.Replace("#", ".");
                            CuartoTri = CuartoTri + Convert.ToDouble(cAuxiliar);
                        }                                              
                    }

                    //Si estoy en el ultimo, pinto
                    if (i < oMatrix.VisualRowCount)
                    {
                        cSiguiente = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(i + 1).Specific).String;
                    }
                    if (i == oMatrix.VisualRowCount || ( cSiguiente != "" && cSiguiente != cNIFOK ))
                       {

                          Suma = Suma.Replace(",", "#");
                          Suma = Suma.Replace(".", ",");
                          Suma = Suma.Replace("#", ".");

                          if ( Convert.ToDouble(Suma) >= 0 )   
                          { cSignoSuma = " "; }
                          else { cSignoSuma = "N"; }

                          if (Convert.ToDouble(PrimeTri) >= 0)
                          { cSignoPriTri = " "; }
                          else { cSignoPriTri = "N"; }

                          if (Convert.ToDouble(SegunTri) >= 0)
                          { cSignoSegunTri = " "; }
                          else { cSignoSegunTri = "N"; }

                          if (Convert.ToDouble(TercerTri) >= 0)
                          { cSignoTer = " "; }
                          else { cSignoTer = "N"; }

                          if (Convert.ToDouble(CuartoTri) >= 0)
                          { cSignoCuar = " "; }
                          else { cSignoCuar = "N"; }
                        
                           StrLinea = "2" +
                                       "347" +
                                       csVariablesGlobales.AñoModelo +
                                       StrNifEmp + // /*NIF EMPRESA*/
                                       cNIFOK.Replace("ES", "")  + /*NIF Cliente*/
                                       "         " +
                                       StrNomCli + /*Nombre Cliente*/
                                       "D" +
                                       StrCpCli + // /*CP CLIENTE*/
                                       "   " +
                                       StrTipoIC +  // /*Tipo IC*/
                                       cSignoSuma + 
                                       (Math.Abs(Convert.ToDouble(Suma) * 100)).ToString("000000000000000") +
                                       " ".PadLeft(2) +
                                       "0".PadLeft(15, '0') +
                                       " " + 
                                       "0".PadLeft(15, '0') +
                                       "0000" +
                                       cSignoPriTri +
                                       (100 * Math.Abs(PrimeTri)).ToString("000000000000000") +
                                       " " + "0".PadLeft(15, '0') +
                                       cSignoSegunTri +
                                       (100 * Math.Abs(SegunTri)).ToString("000000000000000") +
                                       " " + "0".PadLeft(15, '0') +
                                       cSignoTer +
                                       (100 * Math.Abs(TercerTri)).ToString("000000000000000") +
                                       " " + "0".PadLeft(15, '0') +
                                       cSignoCuar +
                                       (100 * Math.Abs(CuartoTri)).ToString("000000000000000") +
                                       " " + "0".PadLeft(15, '0') +
                                       " ".PadRight(237);
                        
                        StrLinea = StrLinea.Replace('(', ' ').Replace(')', ' ').Replace('[', ']').Replace('Á', 'A').
                                 Replace('É', 'E').Replace('Í', 'I').Replace('Ó', 'O').Replace('Ú', 'U').
                                 Replace('Ñ', 'N').Replace('"', ' ').Replace('º', ' ').Replace('ª', ' ');
                        Fichero347.WriteLine(StrLinea);
                        csVariablesGlobales.SboApp.SetStatusBarMessage("Generado IC " + StrCodCli + " NIF " + cNIFOK, BoMessageTime.bmt_Short, false);

                        PrimeTri = 0; SegunTri = 0; TercerTri = 0; CuartoTri = 0;
                        Suma = "";                                          
                    }
                }
                Fichero347.Close();
                csVariablesGlobales.SboApp.MessageBox("Fichero generado correctamente", 1, "Ok", "", "");
                csVariablesGlobales.AñoModelo = "";
            }
            catch (Exception ex)
            {
                Fichero347.Close();
                csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
    }
}
