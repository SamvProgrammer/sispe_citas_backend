using System;

static class principal
{
    string host = SISPE_MIGRACION.Properties.Resources.servidor;
    string usuario = SISPE_MIGRACION.Properties.Resources.usuario;
    string password = SISPE_MIGRACION.Properties.Resources.password;
    string database = SISPE_MIGRACION.Properties.Resources.baseDatos;
    string port = SISPE_MIGRACION.Properties.Resources.puerto;


    //host = "ec2-23-21-160-38.compute-1.amazonaws.com";
    //usuario = "hwvzppntjiviyu";
    //password = "8ec67b7ca03d1e00ba4ac06dc6cdf97e148da0ddc2b7bb69523dec23cdde5256";
    //database = "ddboilk04tmcso";
    //SSL Mode = Require; Trust Server Certificate = true
    string queryConexion = string.Format("Host={0};Username={1};Password={2};Database={3};port={4}", host, usuario, password, database, port);



    globales.datosConexion = queryConexion;
}
