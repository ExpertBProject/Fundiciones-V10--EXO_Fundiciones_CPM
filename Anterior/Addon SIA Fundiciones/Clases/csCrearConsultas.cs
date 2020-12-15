using System;
using System.Collections.Generic;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Addon_SIA
{
    class csCrearConsultas
    {
        Recordset oRecordSet;
        UserQueries oUserQueries;
        QueryCategories oQueryCategories;
        
        public UserQueries ExisteConsulta(string StrQuery)
        {
            oRecordSet = (SAPbobsCOM.Recordset)(csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));
            oUserQueries = (SAPbobsCOM.UserQueries)(csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.oUserQueries));

            try
            {
                oRecordSet.DoQuery("SELECT IntrnalKey, QCategory, QName FROM OUQR WHERE UPPER(LTRIM(RTRIM(QName))) = '" + 
                                    StrQuery.Trim().ToUpper() + "'");
                while (!oRecordSet.EoF)
                {
                    oUserQueries.GetByKey(Convert.ToInt32(oRecordSet.Fields.Item(0).Value), 
                                          Convert.ToInt32(oRecordSet.Fields.Item(1).Value));
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserQueries);
                    return oUserQueries;
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserQueries);
                return null;
            }
            catch
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserQueries);
                return null;
            }
        }

        public int ExisteCategoria(string StrCategoria)
        {
            oRecordSet = (SAPbobsCOM.Recordset)(csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));
            oQueryCategories = (SAPbobsCOM.QueryCategories)(csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.oQueryCategories));

            oRecordSet.DoQuery("SELECT CategoryID FROM OQCN WHERE UPPER(LTRIM(RTRIM(CatName))) = '" + 
                                StrCategoria.Trim().ToUpper() + "'");
            if (!oRecordSet.EoF)
            {
                return Convert.ToInt32((oRecordSet.Fields.Item(0).Value));
            }
            else
            {
                return 0;
            }
        }

        public void CrearCategoriaConsultas(string StrCategoria)
        {
            if (ExisteCategoria(StrCategoria) == 0)
            {
                oQueryCategories = (SAPbobsCOM.QueryCategories)(csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.oQueryCategories));
                oQueryCategories.Name = StrCategoria;
                oQueryCategories.Permissions = "YYYYYYYYYYYYYYY";
                int RetVal = oQueryCategories.Add();
                if (RetVal != 0)
                {
                    csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                }
            }
        }

        public void CrearConsulta(string Categoria, string Consulta, string Descripcion)
        {
            CrearCategoriaConsultas(Categoria);
            if (ExisteConsulta(Descripcion) == null)
            {
                oUserQueries = (SAPbobsCOM.UserQueries)(csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.oUserQueries));
                oUserQueries.QueryCategory = ExisteCategoria(Categoria);
                oUserQueries.Query = Consulta;
                oUserQueries.QueryDescription = Descripcion;
                int RetVal = oUserQueries.Add();
                if (RetVal != 0)
                {
                    csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
                }
            }
        }
    }
}
