using System;

public class Class1
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
                MessageBox.Show(mensaje, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            return conexionRealizada;
        }
        }
    public static dynamic consulta(string query, bool tipoSelect = false, bool eliminando = false)
    {

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
                    MessageBox.Show(e.Message, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    conexion.Close();
                    return false;
                }
            }
            conexion.Close();

        }
        catch (Exception e)
        {
            if (eliminando) return false;
            MessageBox.Show(e.Message, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            conexion.Close();
        }


        return consulta; 
    }




}
