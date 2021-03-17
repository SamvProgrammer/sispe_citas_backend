using Npgsql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace templatesPensiones.herramientas
{
    class baseDatos
    {

        private static string cadenaConexion = string.Empty;
        private static NpgsqlConnection conexion;

        //Cadena de conexión al servidor 192.168.100.102 
        private static string cadenaConexion2 = string.Empty;
        private static NpgsqlConnection conexion2;
        private static bool iniciaTransaccion = false;
        public static bool realizarConexion(string cadena)
        {
            cadenaConexion = cadena;
            bool conexionRealizada = false;
            try
            {
                conexion = new NpgsqlConnection(cadenaConexion);
                conexion.Open();
                conexionRealizada = true;
                conexion.Close();
            }
            catch (Exception e)
            {
                byte[] arreglo = System.Text.Encoding.Default.GetBytes(e.Message);
                string mensaje = System.Text.Encoding.ASCII.GetString(arreglo);
            }
            return conexionRealizada;
        }




        public static dynamic consulta(string query, bool tipoSelect = false, bool eliminando = false)
        {



            string host = templatesPensiones.Properties.Resources.servidor;
            string usuario = templatesPensiones.Properties.Resources.usuario;
            string password = templatesPensiones.Properties.Resources.password;
            string database = templatesPensiones.Properties.Resources.baseDatos;
            string port = templatesPensiones.Properties.Resources.puerto;
            string queryConexion = string.Format("Host={0};Username={1};Password={2};Database={3};port={4}", host, usuario, password, database, port);


            globales.datosConexion = queryConexion;

            baseDatos.realizarConexion(queryConexion);




            var consulta = new List<Dictionary<string, object>>();
            try
            {
                conexion.Open();

                NpgsqlCommand cmd = new NpgsqlCommand(query, conexion);
                if (!tipoSelect)
                {
                    System.Data.Common.DbDataReader datos = cmd.ExecuteReader();

                    while (datos.Read())
                    {
                        Dictionary<string, object> objeto = new Dictionary<string, object>();
                        for (int x = 0; x < datos.FieldCount; x++)
                        {
                            objeto.Add(datos.GetName(x), datos.GetValue(x));
                        }
                        consulta.Add(objeto);
                    }
                    // consulta.Clear();
                    //  consulta = new List<Dictionary<string, object>>();
                }
                else
                {



                    try
                    {
                        cmd.ExecuteNonQuery();
                        conexion.Close();
                        cmd.Dispose();
                        return true;
                    }

                    catch (Exception e)
                    {
                        if (eliminando) return false;
                        conexion.Close();
                        return false;
                    }
                }
                conexion.Close();

            }
            catch (Exception e)
            {
                if (eliminando) return false;
                conexion.Close();
            }


            return consulta;
        }


        public static dynamic web(string query, bool tipoSelect = false, bool eliminando = false)
        {



            //  string host = templatesPensiones.Properties.Resources.servidor;
            string host = "localhost";
            string usuario = "postgres";
            string password = "12345";
            string database = "sispe_web";
            string port = "5432";
            string queryConexion = string.Format("Host={0};Username={1};Password={2};Database={3};port={4}", host, usuario, password, database, port);


            globales.datosConexion = queryConexion;

            baseDatos.realizarConexion(queryConexion);




            var consulta = new List<Dictionary<string, object>>();
            try
            {
                conexion.Open();

                NpgsqlCommand cmd = new NpgsqlCommand(query, conexion);
                if (!tipoSelect)
                {
                    System.Data.Common.DbDataReader datos = cmd.ExecuteReader();

                    while (datos.Read())
                    {
                        Dictionary<string, object> objeto = new Dictionary<string, object>();
                        for (int x = 0; x < datos.FieldCount; x++)
                        {
                            objeto.Add(datos.GetName(x), datos.GetValue(x));
                        }
                        consulta.Add(objeto);
                    }
                    // consulta.Clear();
                    //  consulta = new List<Dictionary<string, object>>();
                }
                else
                {



                    try
                    {
                        cmd.ExecuteNonQuery();
                        conexion.Close();
                        cmd.Dispose();
                        return true;
                    }

                    catch (Exception e)
                    {
                        if (eliminando) return false;
                        conexion.Close();
                        return false;
                    }
                }
                conexion.Close();

            }
            catch (Exception e)
            {
                if (eliminando) return false;
                conexion.Close();
            }


            return consulta;
        }





    }

}
