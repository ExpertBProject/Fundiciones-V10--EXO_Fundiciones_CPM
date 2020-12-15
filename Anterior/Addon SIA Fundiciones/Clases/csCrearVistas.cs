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
    class csCrearVistas
    {
        public void V_SIA_TipoIvaFactVent()
        {
            try
            {
                //csUtilidades csUtilidades = new csUtilidades();
                csUtilidades.LeerConexion(true);
                string Script;
                Script = "USE [" + csVariablesGlobales.oCompany.CompanyDB + "] " +
                        "IF not EXISTS (SELECT * " +
                        "FROM sys.views " +
                        "WHERE object_id = OBJECT_ID(N'[dbo].[V_SIA_TipoIvaFactVent]')) " +
                        "SET ANSI_NULLS ON " +
                        "SET QUOTED_IDENTIFIER ON " +
                        "EXEC( " +
                        "'CREATE VIEW [dbo].[V_SIA_TipoIvaFactVent] AS  " +
                        "SELECT DISTINCT ''F'' AS Tipo, C.DocEntry, C.VatGroup, V.IsEC " +
                        "FROM dbo.INV1 AS C INNER JOIN dbo.OVTG AS V ON C.VatGroup = V.Code " +
                        "UNION " +
                        "SELECT DISTINCT ''A'' AS Tipo, C.DocEntry, C.VatGroup, V.IsEC " +
                        "FROM dbo.RIN1 AS C INNER JOIN dbo.OVTG AS V ON C.VatGroup = V.Code' " +
                        ")";
                SqlCommand cmdCreacionDeVistas = new SqlCommand(Script, csVariablesGlobales.conAddon);
                cmdCreacionDeVistas.ExecuteNonQuery();
                cmdCreacionDeVistas = null;
            }
            catch
            {
                
            }
        }

        public void V_SIA_TipoIvaFactComp()
        {
            try
            {
                //csUtilidades csUtilidades = new csUtilidades();
                csUtilidades.LeerConexion(true);
                string Script;
                Script = "USE [" + csVariablesGlobales.oCompany.CompanyDB + "] " +
                        "IF not EXISTS (SELECT * " +
                        "FROM sys.views " +
                        "WHERE object_id = OBJECT_ID(N'[dbo].[V_SIA_TipoIvaFactComp]')) " +
                        "SET ANSI_NULLS ON " +
                        "SET QUOTED_IDENTIFIER ON " +
                        "EXEC( " +
                        "'CREATE VIEW [dbo].[V_SIA_TipoIvaFactComp] AS  " +
                        "SELECT DISTINCT ''F'' AS Tipo, C.DocEntry, C.VatGroup, V.IsEC " +
                        "FROM dbo.PCH1 AS C INNER JOIN " +
                        "dbo.OVTG AS V ON C.VatGroup = V.Code " +
                        "UNION " +
                        "SELECT DISTINCT ''A'' AS Tipo, C.DocEntry, C.VatGroup, V.IsEC " +
                        "FROM dbo.RPC1 AS C INNER JOIN " +
                        "dbo.OVTG AS V ON C.VatGroup = V.Code' " +
                        ")";
                SqlCommand cmdCreacionDeVistas = new SqlCommand(Script, csVariablesGlobales.conAddon);
                cmdCreacionDeVistas.ExecuteNonQuery();
                cmdCreacionDeVistas = null;
            }
            catch
            {
                
            }
        }
    }
}
