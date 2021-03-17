using Google.Apis.Auth.OAuth2;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using System.Web;

namespace templatesPensiones.herramientas
{
    public class correoElectronico
    { 

        private string accessToken { get; set; }

        public correoElectronico() {
            this.accessToken = GoogleOAuth2.CreateGoogleMailService();
        }


        public void enviarCorreo(string base64,string correo) {
            string uri = $"https://www.googleapis.com/upload/gmail/v1/users/" + $"tecnologias%40pensionesoaxaca.com/messages/send?uploadType=multipart";
            WebClient cliente = new WebClient();
            cliente.Headers[HttpRequestHeader.Authorization] = $"{"Bearer".Trim()} {this.accessToken.Trim()}";
            cliente.Headers[HttpRequestHeader.ContentType] = "message/rfc822";

            string cuerpoMensajeEnviar = enviandoCorreo(correo, base64);


            try
            {
                var ver = cliente.UploadString(uri, cuerpoMensajeEnviar);
            }
            catch (Exception e) {
                
            }
        }


        private string enviandoCorreo(string correo, string codigo)
        {
            string formarMensaje = $"From: PENSIONES <tecnologias@pensionesoaxaca.com>\n";//el correo from jamás debe cambiar

            formarMensaje += $"to: {correo}\n";
            formarMensaje += $"Subject: CITA OFICINA DE PENSIONES\n";
            formarMensaje += "MIME-Version: 1.0\n";
            formarMensaje += "Content-Type: multipart/mixed;\n";
            formarMensaje += "        boundary=\"limite1\"\n\n";
            formarMensaje += "En esta sección se prepara el mensaje\n\n";
            formarMensaje += "--limite1\n";
            formarMensaje += "Content-Type: text/plain\n";
            formarMensaje += $"NOTA: CONVERVAR ESTE ARCHIVO AL MOMENTO DE PRESENTARSE EN LAS OFICINAS\n\n";
            formarMensaje += "--limite1\n";
            formarMensaje += "Content-Type: application/pdf;\n\tname=CITA.pdf;\n";
            formarMensaje += "Content-Transfer-Encoding: BASE64;\n\n";
            formarMensaje += codigo;
            
            return formarMensaje;
        }

    }


    public class GoogleOAuth2
    {
        
        internal static string CreateGoogleMailService()
        {
            // Se obtiene la cuenta de servicio en la consola de google
            // Leer más : https://developers.google.com/identity/protocols/OAuth2ServiceAccount
            const string SERVICE_ACCT_EMAIL = "controlcitas@tribal-bird-270219.iam.gserviceaccount.com";
            //Generar la llave .p12 en la consola de google en la cuenta de servicios.
            var certificate = new X509Certificate2(@"C:\inetpub\wwwroot\mariscal.p12", "notasecret", X509KeyStorageFlags.Exportable);

            var serviceAccountCredentialInitializer = new ServiceAccountCredential.Initializer(SERVICE_ACCT_EMAIL)
            {
                User = "tecnologias@pensionesoaxaca.com", // Muy importante, darle permiso a la cuenta de servicio el acceso a las aplicaciones de gmail
                Scopes = new[] {  "https://mail.google.com/",
              "https://www.googleapis.com/auth/gmail.modify",
              "https://www.googleapis.com/auth/gmail.readonly",
              "https://www.googleapis.com/auth/gmail.metadata"} //Aqui aunque se establece el scope igual en las cuentas de servico se debe realizar...
            }.FromCertificate(certificate);

            var credential = new ServiceAccountCredential(serviceAccountCredentialInitializer);
            if (!credential.RequestAccessTokenAsync(System.Threading.CancellationToken.None).Result)
                throw new InvalidOperationException("Access token failed.");

           return credential.Token.AccessToken;

        }
    }
}