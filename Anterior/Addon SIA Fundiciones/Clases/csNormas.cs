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
    class csNormas
    {
        #region Norma58
        public void GenerarFicheroNorma58(string Ruta, SAPbouiCOM.Form Formulario)
        {
            //csUtilidades csUtilidades = new csUtilidades();
            StreamWriter cFicheroNorma58 = new StreamWriter(Ruta);
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = (SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string cNumDeposito = csUtilidades.DameValorEditText("3", Formulario);
            string cEntidad = csUtilidades.DameValorEditText("20", Formulario);
            string cOficina = csUtilidades.DameValorEditText("12", Formulario);
            string cCuenta = csUtilidades.DameValorEditText("14", Formulario);
            string cNifEmpresa = "", cSufijo, cDatos, cDigito = "", cPobEmpresa = "", cProvEmpresa = "";
            string cCliente, cNombreCliente, cCuentaBancaria, cFechaVencimiento;
            decimal nImporte, nTotalRemesa = 0;
            int nNumRegistroIndi = 0, nNumRegistroTot = 0;

            /*********Datos de la empresa********/

            cNifEmpresa = (csUtilidades.DameValor("OADM", "TaxIdNum", "")).Substring(2).ToString();
            cPobEmpresa = csUtilidades.DameValor("ADM1", "City", "");
            cProvEmpresa = csUtilidades.DameValor("ADM1", "ZipCode", "");
            /*****************digito de control del banco**********/
            cDigito = csUtilidades.DameValor("DSC1", "ControlKey", "BankCode = " + cEntidad);
            cSufijo = "000";
            //Miro el efecto actual
            cDatos = CabeceraPresentador58(cNifEmpresa, cSufijo, cEntidad, cOficina);
            cFicheroNorma58.WriteLine(cDatos);
            cDatos = CabeceraOrdenante58(cNifEmpresa, cSufijo, cEntidad, cOficina, cDigito, cCuenta);
            cFicheroNorma58.WriteLine(cDatos);
            nNumRegistroTot++;
            /*Recorro los efectos*/
            oRecordSet.DoQuery("SELECT T2.[Cardname],T2.[CardCode],T2.[BoeNum], T2.[DueDate], T2.[BoeSum], " +
                               "T2.[BPBankCod],T2.[BPBankBrnc], T2.[ControlKey], T2.[BPBankAct] " +
                               "FROM ODPS T0 , DPS1 T1, OBOE T2 " +
                               "WHERE T0.[DeposId] =  T1.[DepositID] and T1.[CheckKey]=  T2.[BoeKey]" +
                               "AND T0.[DeposId] =" + cNumDeposito);

            while (!oRecordSet.EoF)
            {
                cCliente = Convert.ToString(oRecordSet.Fields.Item("CardCode").Value);
                cNombreCliente = Convert.ToString(oRecordSet.Fields.Item("CardName").Value);
                cCuentaBancaria = Convert.ToString(oRecordSet.Fields.Item("BPBankCod").Value) +
                                  Convert.ToString(oRecordSet.Fields.Item("BPBankBrnc").Value) +
                                  Convert.ToString(oRecordSet.Fields.Item("ControlKey").Value) +
                                  Convert.ToString(oRecordSet.Fields.Item("BPBankAct").Value);
                cFechaVencimiento = Convert.ToDateTime(oRecordSet.Fields.Item("DueDate").Value).ToString("d");
                nImporte = Convert.ToDecimal(oRecordSet.Fields.Item("BoeSum").Value);
                nTotalRemesa = nTotalRemesa + nImporte;
                cDatos = IndividualObligatorio58(cNifEmpresa, cSufijo, cCliente, cNombreCliente, cCuentaBancaria,
                                                   cFechaVencimiento, nImporte);
                nNumRegistroTot++;
                nNumRegistroIndi++;
                cFicheroNorma58.WriteLine(cDatos);
                cDatos = DomicilioObligatorio58(cNifEmpresa, cSufijo, cCliente, cFechaVencimiento, cPobEmpresa, cProvEmpresa);
                nNumRegistroTot++;
                cFicheroNorma58.WriteLine(cDatos);
                oRecordSet.MoveNext();

            }
            nNumRegistroTot++;
            cDatos = TotalOrdenante58(cNifEmpresa, cSufijo, nNumRegistroIndi, nNumRegistroTot, nTotalRemesa);
            cFicheroNorma58.WriteLine(cDatos);
            nNumRegistroTot++;
            nNumRegistroTot++;
            cDatos = TotalGeneral58(cNifEmpresa, cSufijo, nNumRegistroIndi, nNumRegistroTot, nTotalRemesa);
            cFicheroNorma58.WriteLine(cDatos);
            cFicheroNorma58.Close();
            csVariablesGlobales.SboApp.MessageBox("Fichero Generado", 1, "", "", "");
        }

        public string CabeceraPresentador58(string cNifEmpresa, string cSufijo, string cEntidad, string cOficina)
        {
            string cDatos, cRelleno;
            //Cod Registro
            cDatos = "51";
            //cod dato
            cDatos = cDatos + "70";
            //NIF presentador
            cDatos = cDatos + cNifEmpresa;
            //SUFIJO
            cDatos = cDatos + cSufijo;//??????????????????????????
            //Fecha confeccion
            string cFecha;
            DateTime dFecha = DateTime.Now;
            cFecha = dFecha.ToString("d");
            cDatos = cDatos + dFecha.Day.ToString("00") + dFecha.Month.ToString("00") + dFecha.Year.ToString("0000").Substring(2);
            //libre
            cRelleno = "";
            cDatos = cDatos + " ".PadRight(6);
            //Nombre del presentador
            cDatos = cDatos + csVariablesGlobales.SboApp.Company.Name.PadRight(40);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(20);
            //Entidad receptora, oficina
            cDatos = cDatos + cEntidad + cOficina;
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(12);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(40);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(14);
            return cDatos;
        }

        public string CabeceraOrdenante58(string cNifEmpresa, string cSufijo, string cEntidad,
                                          string cOficina, string cDigito, string cCuenta)
        {
            string cDatos, cRelleno;
            //Cod registro cod dato
            cDatos = "53" + "70";
            //cod ordenante
            cDatos = cDatos + cNifEmpresa;
            //sufijo
            cDatos = cDatos + cSufijo;//??????????????????????????
            //fecha confeccion
            string cFecha;
            DateTime dFecha = DateTime.Now;
            cFecha = dFecha.ToString("d");
            cDatos = cDatos + dFecha.Day.ToString("00") + dFecha.Month.ToString("00") + dFecha.Year.ToString("0000").Substring(2);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(6);
            //Nombre del ordenante
            cDatos = cDatos + csVariablesGlobales.SboApp.Company.Name.PadRight(40);
            // Entidad receptora oficina
            cDatos = cDatos + cEntidad + cOficina;
            //digitos de control
            cDatos = cDatos + cDigito;
            //cuenta del ordenante
            cDatos = cDatos + cCuenta;
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(8);
            //tipo procedimiento
            cDatos = cDatos + "06";
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(10);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(40);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(2);
            //codigo INE
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(9);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(3);
            return cDatos;
        }

        public string IndividualObligatorio58(string cNifEmpresa, string cSufijo, string cCliente, string cNombreCliente,
                                              string cCuentaBancaria, string cFechaVencimiento, decimal nImporte)
        {
            string cDatos, cRelleno;
            //Cod registro cod dato
            cDatos = "56" + "70";
            //cod ordenante
            cDatos = cDatos + cNifEmpresa;
            //sufijo
            cDatos = cDatos + cSufijo;//??????????????????????????
            //cod cliente
            cDatos = cDatos + cCliente.PadLeft(12, '0').Substring(0, 12);
            //Nombre cliente
            cDatos = cDatos + cNombreCliente.PadRight(40).Substring(0, 40);
            //entidad,oficina,cod control cuenta
            cDatos = cDatos + cCuentaBancaria.PadRight(20).Substring(0, 20);
            //importe
            nImporte = nImporte * 100;
            cDatos = cDatos + Convert.ToInt32(nImporte).ToString("0000000000");
            //cod devoluciones
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(6);
            // cod referencia interna
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(10);
            // primer campo concepto
            cRelleno = "";//????????????????????????????necesito el numero de documento
            cDatos = cDatos + cRelleno.PadRight(40);
            //fecha vencimiento
            cDatos = cDatos + cFechaVencimiento.Substring(0, 2) + cFechaVencimiento.Substring(3, 2) + cFechaVencimiento.Substring(8, 2);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(2);
            return cDatos;
        }

        public string DomicilioObligatorio58(string cNifEmpresa, string cSufijo, string cCliente, string cFechaVencimiento,
                                             string cPobEmpresa, string cProvEmpresa)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = (SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string cDatos, cRelleno;
            //Cod registro cod dato
            cDatos = "56" + "76";
            //cod ordenante
            cDatos = cDatos + cNifEmpresa;
            //sufijo
            cDatos = cDatos + cSufijo;//??????????????????????????
            /*********Datos del cliente********/
            //cod cliente
            cDatos = cDatos + cCliente.PadLeft(12, '0').Substring(0, 12);
            oRecordSet.DoQuery("SELECT * FROM OCRD T0 WHERE CardCode = '" + cCliente + "'");
            while (!oRecordSet.EoF)
            {
                //domicilio del deudor
                cDatos = cDatos + Convert.ToString(oRecordSet.Fields.Item("Address").Value).PadRight(40).Substring(0, 40);
                //plaza del deudor
                cDatos = cDatos + Convert.ToString(oRecordSet.Fields.Item("City").Value).PadRight(35).Substring(0, 35);
                // codigo postal
                cDatos = cDatos + Convert.ToString(oRecordSet.Fields.Item("ZipCode").Value).PadRight(5).Substring(0, 5);
                oRecordSet.MoveNext();
            }
            //poblacion ordenante
            cDatos = cDatos + cPobEmpresa.PadRight(38).Substring(0, 38);
            //provincia ordenante
            cDatos = cDatos + cProvEmpresa.PadRight(5).Substring(0, 2);
            //fecha en que se formalizo el contrato-fecha de remesa
            cDatos = cDatos + cFechaVencimiento.Substring(0, 2) + cFechaVencimiento.Substring(3, 2) + cFechaVencimiento.Substring(8, 2);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(8);
            return cDatos;

        }

        public string TotalOrdenante58(string cNifEmpresa, string cSufijo, int nNumRegistroIndi,
                                       int nNumRegistroTot, decimal nTotRemesa)
        {
            string cDatos, cRelleno;
            //Cod registro cod dato
            cDatos = "58" + "70";
            //cod ordenante
            cDatos = cDatos + cNifEmpresa;
            //sufijo
            cDatos = cDatos + cSufijo;//??????????????????????????
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(12);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(40);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(20);
            //suma importe
            cDatos = cDatos + (nTotRemesa * 100).ToString("0000000000");
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(6);
            //num reg individuales
            cDatos = cDatos + nNumRegistroIndi.ToString("0000000000");
            //numero registro
            cDatos = cDatos + nNumRegistroTot.ToString("0000000000");
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(20);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(18);
            return cDatos;
        }

        public string TotalGeneral58(string cNifEmpresa, string cSufijo, int nNumRegistroIndi,
                                     int nNumRegistroTot, decimal nTotRemesa)
        {
            string cDatos, cRelleno;
            //Cod registro cod dato
            cDatos = "59" + "70";
            //cod ordenante
            cDatos = cDatos + cNifEmpresa;
            //sufijo
            cDatos = cDatos + cSufijo;//??????????????????????????
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(12);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(40);
            //numero de ordenantes diferentes
            cDatos = cDatos + "0001";
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(16);
            //suma de importes
            cDatos = cDatos + (nTotRemesa * 100).ToString("0000000000");
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(6);
            //numero de creditos
            cDatos = cDatos + nNumRegistroIndi.ToString("0000000000");
            //nº de registros
            cDatos = cDatos + nNumRegistroTot.ToString("0000000000");
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(20);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(18);
            return cDatos;
        }
        #endregion
        #region Norma19
        public void GeneroFicheroNorma19(string Ruta, SAPbouiCOM.Form Formulario)
        {
            //csUtilidades csUtilidades = new csUtilidades();
            StreamWriter cFicheroNorma19 = new StreamWriter(Ruta);
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = (SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string cNumDeposito = csUtilidades.DameValorEditText("3", Formulario);
            string cEntidad = csUtilidades.DameValorEditText("20", Formulario);
            string cOficina = csUtilidades.DameValorEditText("12", Formulario);
            string cCuenta = csUtilidades.DameValorEditText("14", Formulario);
            string cFechaRemesa = csUtilidades.DameValorEditText("9", Formulario);
            string cNifEmpresa = "", cSufijo = "", cDatos, cDigito = "";
            string cCliente, cNombreCliente, cCuentaBancaria, cFechaVencimiento;
            decimal nImporte, nTotalRemesa = 0;
            int nNumRegistroIndi = 0, nNumRegistroTot = 0;

            /*********Datos de la empresa********/
            cNifEmpresa = (csUtilidades.DameValor("OADM", "TaxIdNum", "")).Substring(2).ToString();
            /*****************digito de control del banco**********/
            cDigito = csUtilidades.DameValor("DSC1", "ControlKey", "BankCode = " + cEntidad);
            cSufijo = csUtilidades.DameValor("DSC1", "UsrNumber2", "BankCode = " + cEntidad).PadLeft(3);
            //Miro el efecto actual
            cDatos = CabeceraPresentador19(cNifEmpresa, cSufijo, cEntidad, cOficina);
            cFicheroNorma19.WriteLine(cDatos);
            cDatos = CabeceraOrdenante19(cNifEmpresa, cSufijo, cEntidad, cOficina, cDigito, cCuenta, cFechaRemesa);
            cFicheroNorma19.WriteLine(cDatos);
            nNumRegistroTot++;
            /*Recorro los efectos*/
            oRecordSet.DoQuery("SELECT T2.[Cardname],T2.[CardCode],T2.[BoeNum], T2.[DueDate], T2.[BoeSum], " +
                               "T2.[BPBankCod],T2.[BPBankBrnc], T2.[ControlKey], T2.[BPBankAct] " +
                               "FROM ODPS T0 , DPS1 T1, OBOE T2 " +
                               "WHERE T0.[DeposId] =  T1.[DepositID] and T1.[CheckKey]=  T2.[BoeKey]" +
                               "AND T0.[DeposId] =" + cNumDeposito);
            while (!oRecordSet.EoF)
            {
                cCliente = Convert.ToString(oRecordSet.Fields.Item("CardCode").Value);
                cNombreCliente = Convert.ToString(oRecordSet.Fields.Item("CardName").Value);
                cCuentaBancaria = Convert.ToString(oRecordSet.Fields.Item("BPBankCod").Value) +
                                  Convert.ToString(oRecordSet.Fields.Item("BPBankBrnc").Value) +
                                  Convert.ToString(oRecordSet.Fields.Item("ControlKey").Value) +
                                  Convert.ToString(oRecordSet.Fields.Item("BPBankAct").Value);
                cFechaVencimiento = Convert.ToDateTime(oRecordSet.Fields.Item("DueDate").Value).ToString("d");
                nImporte = Convert.ToDecimal(oRecordSet.Fields.Item("BoeSum").Value);
                nTotalRemesa = nTotalRemesa + nImporte;
                cDatos = IndividualObligatorio19(cNifEmpresa, cSufijo, cCliente, cNombreCliente, cCuentaBancaria, nImporte);
                nNumRegistroTot++;
                nNumRegistroIndi++;
                cFicheroNorma19.WriteLine(cDatos);
                cDatos = IndividualOpcional19(cNifEmpresa, cSufijo, cCliente);
                nNumRegistroTot++;
                cFicheroNorma19.WriteLine(cDatos);
                oRecordSet.MoveNext();

            }
            nNumRegistroTot++;
            cDatos = TotalOrdenante19(cNifEmpresa, cSufijo, nNumRegistroIndi, nNumRegistroTot, nTotalRemesa);
            cFicheroNorma19.WriteLine(cDatos);
            nNumRegistroTot++;
            nNumRegistroTot++;
            cDatos = TotalGeneral19(cNifEmpresa, cSufijo, nNumRegistroIndi, nNumRegistroTot, nTotalRemesa);
            cFicheroNorma19.WriteLine(cDatos);
            cFicheroNorma19.Close();
            csVariablesGlobales.SboApp.MessageBox("Fichero Generado", 1, "", "", "");

        }

        public string CabeceraPresentador19(string cNifEmpresa, string cSufijo, string cEntidad, string cOficina)
        {

            string cDatos, cRelleno;
            //Cod registro cod dato
            cDatos = "51" + "80";
            //cod ordenante
            cDatos = cDatos + cNifEmpresa;
            //sufijo
            cDatos = cDatos + cSufijo;//??????????????????????????
            //fecha confeccion
            string cFecha;
            DateTime dFecha = DateTime.Now;
            cFecha = dFecha.ToString("d");
            cDatos = cDatos + dFecha.Day.ToString("00") + dFecha.Month.ToString("00") + dFecha.Year.ToString("0000").Substring(2);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(6);
            //Nombre del ordenante
            cDatos = cDatos + csVariablesGlobales.SboApp.Company.Name.PadRight(40);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(20);

            // Entidad receptora oficina
            cDatos = cDatos + cEntidad + cOficina;

            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(12);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(40);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(14);
            return cDatos;
        }

        public string CabeceraOrdenante19(string cNifEmpresa, string cSufijo, string cEntidad, string cOficina,
                                          string cDigito, string cCuenta, string cFechaRemesa)
        {

            string cDatos, cRelleno;
            //Cod registro cod dato
            cDatos = "53" + "80";
            //cod ordenante
            cDatos = cDatos + cNifEmpresa;
            //sufijo
            cDatos = cDatos + cSufijo;
            //fecha confeccion
            string cFecha;
            DateTime dFecha = DateTime.Now;
            cFecha = dFecha.ToString("d");
            cDatos = cDatos + dFecha.Day.ToString("00") + dFecha.Month.ToString("00") + dFecha.Year.ToString("0000").Substring(2);
            //Fecha de cargo Fecha de la remesa
            cDatos = cDatos + cFechaRemesa.Substring(0, 2) + cFechaRemesa.Substring(3, 2) + cFechaRemesa.Substring(8, 2);
            //Nombre del ordenante
            cDatos = cDatos + csVariablesGlobales.SboApp.Company.Name.PadRight(40);
            // Entidad receptora oficina
            cDatos = cDatos + cEntidad + cOficina;
            //digitos de control
            cDatos = cDatos + cDigito;
            //cuenta del ordenante
            cDatos = cDatos + cCuenta;
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(8);
            //tipo procedimiento
            cDatos = cDatos + "02";
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(10);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(40);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(14);
            return cDatos;
        }

        public string IndividualObligatorio19(string cNifEmpresa, string cSufijo, string cCliente, string cNombreCliente,
                        string cCuentaBancaria, decimal nImporte)
        {
            string cDatos, cRelleno;
            //Cod registro cod dato
            cDatos = "56" + "80";
            //cod ordenante
            cDatos = cDatos + cNifEmpresa;
            //sufijo
            cDatos = cDatos + cSufijo;//??????????????????????????
            //cod cliente
            cDatos = cDatos + cCliente.PadLeft(12, '0').Substring(0, 12);
            //Nombre cliente
            cDatos = cDatos + cNombreCliente.PadRight(40).Substring(0, 40);
            //entidad,oficina,cod control cuenta
            cDatos = cDatos + cCuentaBancaria.PadRight(20).Substring(0, 20);
            //importe
            nImporte = nImporte * 100;
            cDatos = cDatos + Convert.ToInt32(nImporte).ToString("0000000000");
            // libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(16);
            // primer campo concepto
            cRelleno = "";//????????????????????????????necesito el numero de documento
            cDatos = cDatos + cRelleno.PadRight(17);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(23);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(8);
            return cDatos;

        }

        public string IndividualOpcional19(string cNifEmpresa, string cSufijo, string cCliente)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = (SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string cDatos, cRelleno;
            //Cod registro cod dato
            cDatos = "56" + "86";
            //cod ordenante
            cDatos = cDatos + cNifEmpresa;
            //sufijo
            cDatos = cDatos + cSufijo;//??????????????????????????
            /*********Datos del cliente********/
            //cod cliente
            cDatos = cDatos + cCliente.PadLeft(12, '0').Substring(0, 12);
            oRecordSet.DoQuery("SELECT * FROM OCRD T0 WHERE CardCode = '" + cCliente + "'");
            while (!oRecordSet.EoF)
            {
                //nombre
                cDatos = cDatos + Convert.ToString(oRecordSet.Fields.Item("CardName").Value).PadRight(40).Substring(0, 40);
                //domicilio del deudor
                cDatos = cDatos + Convert.ToString(oRecordSet.Fields.Item("Address").Value).PadRight(40).Substring(0, 40);
                //plaza del deudor
                cDatos = cDatos + Convert.ToString(oRecordSet.Fields.Item("City").Value).PadRight(35).Substring(0, 35);
                // codigo postal
                cDatos = cDatos + Convert.ToString(oRecordSet.Fields.Item("ZipCode").Value).PadRight(5).Substring(0, 5);
                oRecordSet.MoveNext();
            }

            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(14);
            return cDatos;

        }

        public string TotalOrdenante19(string cNifEmpresa, string cSufijo, int nNumRegistroIndi,
                                       int nNumRegistroTot, decimal nTotRemesa)
        {
            string cDatos, cRelleno;
            //Cod registro cod dato
            cDatos = "58" + "80";
            //cod ordenante
            cDatos = cDatos + cNifEmpresa;
            //sufijo
            cDatos = cDatos + cSufijo;//??????????????????????????
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(12);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(40);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(20);
            //suma importe
            cDatos = cDatos + (nTotRemesa * 100).ToString("0000000000");
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(6);
            //num reg individuales
            cDatos = cDatos + nNumRegistroIndi.ToString("0000000000");
            //numero registro
            cDatos = cDatos + nNumRegistroTot.ToString("0000000000");
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(20);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(18);
            return cDatos;
        }

        public string TotalGeneral19(string cNifEmpresa, string cSufijo, int nNumRegistroIndi,
                                     int nNumRegistroTot, decimal nTotRemesa)
        {
            string cDatos, cRelleno;
            //Cod registro cod dato
            cDatos = "59" + "80";
            //cod ordenante
            cDatos = cDatos + cNifEmpresa;
            //sufijo
            cDatos = cDatos + cSufijo;//??????????????????????????
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(12);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(40);
            //numero de ordenantes diferentes
            cDatos = cDatos + "0001";
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(16);
            //suma de importes
            cDatos = cDatos + (nTotRemesa * 100).ToString("0000000000");
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(6);
            //numero de creditos
            cDatos = cDatos + nNumRegistroIndi.ToString("0000000000");
            //nº de registros
            cDatos = cDatos + nNumRegistroTot.ToString("0000000000");
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(20);
            //libre
            cRelleno = "";
            cDatos = cDatos + cRelleno.PadRight(18);
            return cDatos;
        }
        #endregion
    }
}
