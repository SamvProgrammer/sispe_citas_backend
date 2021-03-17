using System;

internal static string datosConexion;


public class globales
{
    public static dynamic consulta(string consulta, bool tipoSelect = false, bool eliminando = false)
    {
        return baseDatos.consulta(consulta, tipoSelect, eliminando);
    }

}
