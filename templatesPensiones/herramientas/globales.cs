using SISPE_MIGRACION.codigo.herramientas.forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;



namespace templatesPensiones.herramientas
{


    public class globales
    {


        internal static string datosConexion;

        public static dynamic consulta(string consulta, bool tipoSelect = false, bool eliminando = false)
        {
            return baseDatos.consulta(consulta, tipoSelect, eliminando);
        }

        public static dynamic web(string consulta, bool tipoSelect = false, bool eliminando = false)
        {
            return baseDatos.web(consulta, tipoSelect, eliminando);
        }

        public static double convertDouble(string numero)
        {
            string strNumero = (string.IsNullOrWhiteSpace(numero)) ? "0" : numero;
            double dblNumero = 0;
            try
            {
                dblNumero = double.Parse(strNumero, System.Globalization.NumberStyles.Currency);
            }
            catch
            {
                dblNumero = 0;
            }
            return dblNumero;
        }



        public static int convertInt(string numero)
        {
            string strNumero = (string.IsNullOrWhiteSpace(numero)) ? "0" : numero;
            int numeroaux = 0;
            try
            {
                numeroaux = int.Parse(strNumero, System.Globalization.NumberStyles.Currency);
            }
            catch
            {
                numeroaux = 0;
            }
            return numeroaux;
        }
        public static double convtDouble(string numero)
        {
            string strNumero = (string.IsNullOrWhiteSpace(numero)) ? "0" : numero;
            double dblNumero = 0;
            try
            {
                dblNumero = double.Parse(strNumero, System.Globalization.NumberStyles.Currency);
            }
            catch
            {
                dblNumero = 0;
            }
            return dblNumero;
        }




        public static byte[] reportes(string nombreReporte, string tablaDataSet, object[] objeto, string mensaje = "", bool imprimir = false, object[] parametros = null, bool espdf = false, string nombrePdf = "", string subcarpeta = "")

        {
            frmReporte reporte = new frmReporte(nombreReporte, subcarpeta, tablaDataSet);


            reporte.setParametrosExtra(espdf, nombrePdf);
            reporte.cargarDatos(tablaDataSet, objeto, mensaje, imprimir, parametros);

            return reporte.bytes;
        }


        public static string convertirNumerosLetras(string numero, bool mayusculas)
        {
            string consulta = $"select * from numeltra ({globales.convertDouble(numero)})";
            List<Dictionary<string, object>> res = globales.web(consulta);


            return Convert.ToString(res[0]["numeltra"]);
        }



    }
}