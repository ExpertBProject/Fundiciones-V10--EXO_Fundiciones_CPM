using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cliente
{
    public static class Especifico
    {
        public static string sqlQueryCalculo(string cGrupoArticulo, string cDesdeArt, string cHastaArt, string cDesdeFecha, string cHastaFecha, string cListaAlmacenes, bool lDesglosada, string cTipoCalculo)
        {

            string sql = "EXECUTE EXO_CalculoPrecMedio '##DESDEARTICULO', '##HASTAARTICULO', ##GRUPOARTICULO, '##DESDEFECHA', '##HASTAFECHA', '##LISTAALMACENES'";
            sql += lDesglosada ? ", 'Y'" : ", 'N'";
            sql += ", '##TIPOCALCULO' ";

            sql = sql.Replace("##DESDEFECHA", cDesdeFecha).Replace("##HASTAFECHA", cHastaFecha).Replace("##DESDEARTICULO", cDesdeArt).Replace("##HASTAARTICULO", cHastaArt);
            sql = sql.Replace("##GRUPOARTICULO", cGrupoArticulo).Replace("##LISTAALMACENES", cListaAlmacenes).Replace("##GRUPOARTICULO", cGrupoArticulo).Replace("##TIPOCALCULO", cTipoCalculo);

            return sql;
        }

    }
}
