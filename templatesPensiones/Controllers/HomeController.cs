using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Mvc;
using templatesPensiones.herramientas;
using templatesPensiones.Model;

namespace templatesPensiones.Controllers
{


    public class HomeController : Controller
    {

        public List<Dictionary<string, object>> lista;

        public ActionResult Index(string enviado = "")
        {

            ViewBag.enviado = enviado;
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }


        [HttpGet]
        public JsonResult getAreas()
        {
            string query = "SELECT id , nombre FROM oficialia.areas_pensiones where web=true;";
            List<Dictionary<string, object>> resultado = globales.consulta(query);
            return Json(resultado, JsonRequestBehavior.AllowGet);

        }


        [HttpGet]
        public JsonResult getTramite(string id_depto)
        {

            string query = $"SELECT nombre , id FROM oficialia.tramites where id_depto={id_depto}  and web =true;";
            List<Dictionary<string, object>> resultado = globales.consulta(query);


            return Json(resultado, JsonRequestBehavior.AllowGet);

        }


        [HttpGet]
        public JsonResult getCitas(string dia, string id_tramite, string id_depto)    //metodo para desglosar citas de un día 
        {

            string query = $"SELECT dia , hora FROM oficialia.control_citas where dia='{dia}' and id_tramite ={id_tramite} and id_depto ={id_depto} and status = false;";
            List<Dictionary<string, object>> resultado = globales.consulta(query);
            return Json(resultado, JsonRequestBehavior.AllowGet);

        }



        [HttpGet]
        public JsonResult getDiasDisponibles(int id_depto, int id_tramite)
        {
            string query = $" select dia  from oficialia.control_citas where status=FALSE and id_depto={id_depto} and id_tramite={id_tramite}  GROUP BY dia";
            List<Dictionary<string, object>> resultado = globales.consulta(query);

            foreach (Dictionary<string,object> item in resultado) {
                string query2 = $"SELECT  hora FROM oficialia.control_citas where dia='{string.Format("{0:yyyy-MM-dd}", DateTime.Parse(Convert.ToString(item["dia"])))}' and id_tramite ={id_tramite} and id_depto ={id_depto} and status = false;";
                List<Dictionary<string, object>> resultado2 = globales.consulta(query2);
                item.Add("horas",resultado2);
            }

            return Json(resultado, JsonRequestBehavior.AllowGet);
        }


        [HttpPost]
        public JsonResult SubmitCita(repositorio repo)
        {
            string query = $"update oficialia.control_citas set nombre='{repo.nombre}' , apellido='{repo.apellido}'  , correo ='{repo.correo}' , celular ='{repo.celular}' ,status=true ,tramite = '{repo.tramite}'  where hora ='{repo.hora}' and dia = '{string.Format("{0:yyyy-MM-dd}", repo.dia)}' and status=false returning *;";
            List<Dictionary<string, object>> resultado = globales.consulta(query);


            if (resultado.Count == 0)
                throw new Exception("Cita ocupada");

            Dictionary<string, object> diccionario = resultado[0];

            string fecha = Convert.ToString(diccionario["dia"]);
            string hora = Convert.ToString(diccionario["hora"]);
            string nombre = Convert.ToString(diccionario["nombre"]);
            string apellido = Convert.ToString(diccionario["apellido"]);
            string correo = Convert.ToString(diccionario["correo"]);
            string tramite = Convert.ToString(diccionario["tramite"]);



            object[][] parameters = new object[2][];

            object[] headers = { "nombre","apellido","tramite","fecha","hora" };
            object[] body = { nombre,apellido,tramite,string.Format("{0:d}",DateTime.Parse(fecha)),hora };
            parameters[0] = headers;
            parameters[1] = body;

            object[] cuerpo = new object[1];
            object[] tt1 = { "", "", "" };

            cuerpo[0] = tt1;



            byte[] pdfBytes =   globales.reportes("reporte1", "usuario", cuerpo, "",false, parameters);
            string base64pdf = Convert.ToBase64String(pdfBytes);

            correoElectronico objCorreo = new correoElectronico();

            objCorreo.enviarCorreo(base64pdf,correo);

            Dictionary<string, string> rsultado = new Dictionary<string, string>();
            rsultado.Add("resultado","enviado al correo");

            return Json(rsultado, JsonRequestBehavior.AllowGet);
        }






        //////// A PARTIR DE AQUÍ ES CÓDIGO DEL SISPE WEB


        [HttpGet]
        public HttpRequestMessage GeneraResumen(string emplea , string tiponomina , string descripcion, string cve)
        {

            string distinto = $"SELECT DISTINCT (rfc),a1.num FROM maestro a1 JOIN nominew a2 ON a1.num = a2.numjpp AND a1.jpp = a2.jpp WHERE a1.jpp = a2.jpp AND a1.jpp = '{emplea}' AND a2.tiponomina = '{tiponomina}' ORDER BY   a1.num ";
            List<Dictionary<string, object>> listaaa = globales.web(distinto);
            int cantidad = listaaa.Count;
            // Resumen de Nómina
            string queryresumen = $"SELECT	a2.clave,	a2.tipopago,	MAX (a3.descripcion) AS descri,	SUM (monto) AS monto FROM	maestro a1 LEFT JOIN nominew a2 ON a1.jpp = a2.jpp JOIN perded a3 ON a3.clave = a2.clave AND a1.num = a2.numjpp"
                + $" WHERE a1.jpp = '{emplea}' AND a2.tiponomina = '{tiponomina}' GROUP BY a2.clave, a2.tipopago ORDER BY a2.clave,a2.tipopago; ";


            List<Dictionary<string, object>> resumen = globales.web(queryresumen);
            object[] aux5 = new object[resumen.Count];
            int cont = 0;
            int veces = resumen.Count;
            double sumap = 0;
            double sumad = 0;

            foreach (var itemr in resumen)
            {

                string claveR = string.Empty;
                string deducR = string.Empty;
                string MontoR = string.Empty;
                string MooR = string.Empty;

                string LeyenR = string.Empty;
                string tipopago = string.Empty;
                string retro = "RETROACTIVO";

                object[] tt1 = { "", "", "", "", "", "", "" };

                claveR = Convert.ToString(itemr["clave"]);
                deducR = Convert.ToString(itemr["descri"]);
                MooR = Convert.ToString(itemr["monto"]);
                MontoR = string.Format("{0:c}", Convert.ToDouble(MooR));
                tipopago = Convert.ToString(itemr["tipopago"]);


                if (Convert.ToInt32(claveR) < 200)
                {
                    tt1[0] = claveR;
                    tt1[1] = deducR;
                    tt1[2] = MontoR;
                    if (tipopago == "R")
                        tt1[3] = retro;
                    aux5[cont] = tt1;
                    sumap = Convert.ToDouble(MooR) + sumap;
                }
                else
                {
                    tt1[3] = claveR;
                    tt1[4] = deducR;
                    tt1[5] = MontoR;
                    if (tipopago == "R")
                        tt1[7] = retro;
                    aux5[cont] = tt1;
                    sumad = Convert.ToDouble(MooR) + sumad;

                }
                cont++;

            }

            List<object> listaResumen = new List<object>();
            foreach (object item in aux5)
            {
                if (item == null)
                    break;
                listaResumen.Add(item);
            }


            aux5 = new object[listaResumen.Count];

            int y = 0;
            foreach (object item in listaResumen)
            {
                aux5[y] = item;
                y++;
            }

            string desc = $"RESUMEN CONTABLE DE LAS INCIDENCIAS ELECTRÓNICA PARA EL PAGO A ";

            double operacion = Convert.ToDouble(sumap) - Convert.ToDouble(sumad);
            string letra = $"LA PRESENTE NOMINA AMPARA LA CANTIDAD DE:{globales.convertirNumerosLetras(Convert.ToString(operacion), true)}";
            object[] parametrosR = { "descripcion", "sumap", "sumad", "liquido", "conteo", "letra", "cve" };
            object[] valorR = { desc, string.Format("{0:C}", Convert.ToDouble(sumap)), string.Format("{0:C}", Convert.ToDouble(sumad)), string.Format("{0:C}", Convert.ToDouble(sumap) - Convert.ToDouble(sumad)), Convert.ToString(cantidad), letra, cve };
            object[][] enviarParametrosR = new object[2][];

            enviarParametrosR[0] = parametrosR;
            enviarParametrosR[1] = valorR;
            //Restablece los objetos para evitar el break del reporteador
            //  globales.reportes("resumen", "resumen", aux5, "", false, enviarParametrosR);

            byte[] bytes2 = globales.reportes("resumen", "resumen", aux5, "", true, enviarParametrosR, true);

            string base642 = Convert.ToBase64String(bytes2);
            MemoryStream enmemoria2 = new MemoryStream(bytes2);
            byte[] bytesInStream2 = enmemoria2.ToArray();
            Response.Clear();
            Response.ContentType = "application/force-download";
            Response.AddHeader("content-disposition", "attachment;    filename=resumen.pdf");
            Response.BinaryWrite(bytesInStream2);
            Response.End();


            return null;
        }


        [HttpGet]
        public HttpResponseMessage GeneraNomina(string emplea, string tiponomina, string descripcion, string cve)
        {

            string query = $"SELECT	a1.categ,CONCAT (a1.jpp, a1.num) AS proyecto,a1.depto,a1.nombre,a1.curp,	a1.rfc,	a2.clave,	a2.descri,	a2.monto,	a2.pagon,	a2.pagot" +
             $" FROM maestro a1 JOIN nominew a2 ON a1.num = a2.numjpp AND a1.jpp = a2.jpp WHERE a1.jpp = a2.jpp AND a1.jpp = '{emplea}' AND a2.tiponomina = '{tiponomina}' ORDER BY   a1.jpp,	a1.num,	a2.clave";
            List<Dictionary<string, object>> resultado = globales.web(query);

            query = "select  clave,descripcion from perded order by clave";
            List<Dictionary<string, object>> perded = globales.web(query);

            resultado.ForEach(o =>
            {
                o["descripcion"] = perded.Where(p => Convert.ToString(o["clave"]) == Convert.ToString(p["clave"])).First()["descripcion"];
                //  o["descri"] += " (RETROACTIVO)";
            });

            object[] aux2 = new object[resultado.Count];
            int contadorPercepcion = 0;
            int contadorDeduccion = 0;



            string archivoPrimero = string.Empty;


            archivoPrimero = resultado[0]["proyecto"].ToString();


            foreach (var item in resultado)
            {
                string proyecto = string.Empty;
                string nombre = string.Empty;
                string curp = string.Empty;
                string rfc = string.Empty;
                string imss = string.Empty;
                string categ = string.Empty;
                string clave = string.Empty;
                string descri = string.Empty;
                double monto = 0;
                //  fecha = fec2.ToString();
                string archivo = string.Empty;
                string pagon = string.Empty;
                string pagot = string.Empty;
                string depto = string.Empty;


                categ = Convert.ToString(item["categ"]);
                proyecto = Convert.ToString(item["proyecto"]);
                depto = Convert.ToString(item["depto"]);
                nombre = Convert.ToString(item["nombre"]);
                rfc = Convert.ToString(item["rfc"]);
                clave = Convert.ToString(item["clave"]);
                descri = Convert.ToString(item["descri"]);
                monto = globales.convertDouble(Convert.ToString(item["monto"]));
                archivo = Convert.ToString(item["proyecto"]);
                pagon = Convert.ToString(item["pagon"]);
                pagot = Convert.ToString(item["pagot"]);




                object[] tt1 = { proyecto, rfc, nombre, categ, "", "", "", "", "", "", "", depto };

                if (archivoPrimero != archivo)
                {
                    archivoPrimero = archivo;
                    int tope = contadorDeduccion <= contadorPercepcion ? contadorPercepcion : contadorDeduccion;
                    contadorDeduccion = tope;
                    contadorPercepcion = tope;
                }

                if (Convert.ToInt32(clave) <= 201)
                {
                    if (aux2[contadorPercepcion] == null)
                    {
                        tt1[4] = clave;
                        tt1[5] = descri;
                        tt1[6] = monto;
                        aux2[contadorPercepcion] = tt1;
                    }
                    else
                    {
                        object[] tmp = (object[])aux2[contadorPercepcion];
                        tmp[4] = clave;
                        tmp[5] = descri;
                        tmp[6] = monto;
                    }
                    contadorPercepcion++;
                }
                else
                {

                    if (aux2[contadorDeduccion] == null)
                    {
                        tt1[7] = clave;
                        tt1[8] = descri;
                        tt1[10] = (string.IsNullOrWhiteSpace(pagon) || pagot == "0") ? "" : $"{pagon}/{pagot}";
                        tt1[9] = monto;
                        aux2[contadorDeduccion] = tt1;
                    }
                    else
                    {
                        object[] tmp = (object[])aux2[contadorDeduccion];
                        tmp[7] = clave;
                        tmp[8] = descri;
                        tmp[10] = (string.IsNullOrWhiteSpace(pagon) || pagot == "0") ? "" : $"{pagon}/{pagot}";
                        tmp[9] = monto;
                    }
                    contadorDeduccion++;
                }

            }



            int contador = 0;

            List<object> lista = new List<object>();
            foreach (object item in aux2)
            {
                if (item == null)
                    break;
                lista.Add(item);
            }


            aux2 = new object[lista.Count];

            int x = 0;
            foreach (object item in lista)
            {
                aux2[x] = item;
                x++;
            }



            descripcion = $"REPORTE DE INCIDENCIAS ELECTRÓNICO PARA EL ";
            string fecha = "01/02/2020";
            cve = "545645464";

            object[] parametros = { "descripcion", "fecha", "cve_presupuestal" };
            object[] valor = { descripcion, fecha, cve };
            object[][] enviarParametros = new object[2][];

            enviarParametros[0] = parametros;
            enviarParametros[1] = valor;

            byte[] bytes = globales.reportes("reporteGeneracionNominas", "nomina", aux2, "", true, enviarParametros, true);

            string base64 = Convert.ToBase64String(bytes);
            MemoryStream enmemoria = new MemoryStream(bytes);
            byte[] bytesInStream = enmemoria.ToArray();
            Response.Clear();
            Response.ContentType = "application/force-download";
            Response.AddHeader("content-disposition", "attachment;    filename=nomina.pdf");
            Response.BinaryWrite(bytesInStream);
            Response.End();

            return null;


        }

        [HttpGet]
        public HttpResponseMessage ProductosNominaQuincenal(String quincena)
        {
            string query = string.Empty;
            if (Convert.ToInt32(quincena) == 09 || Convert.ToInt32(quincena)==11 )
            {

                query = $"SELECT a1.rfc,	a1.nombre,	a1.cvecat,	a1.categ,    case a2.clave      when 1 then   (a2.monto*2)       when 70 then (a2.monto*2) end as sueldo, case a2.clave when 1 then (a2.monto*2) * 0.09     when 70 then (a2.monto*2) * 0.09   end as obrero,  case a2.clave" +
             " when 1 then (a2.monto*2) * 0.185 when 70 then (a2.monto*2) * 0.185  end as patronal FROM maestro a1 left JOIN nominew a2 ON a1.jpp = a2.jpp " +
               " AND a1.num = a2.numjpp where a2.clave in (1,70) GROUP BY  a1.rfc,	a1.nombre,	a1.cvecat,	a1.categ,   a2.clave,   a2.monto";
            }
            else
            {
                query = $"SELECT a1.rfc,	a1.nombre,	a1.cvecat,	a1.categ,    case a2.clave      when 1 then   a2.monto       when 70 then a2.monto end as sueldo, case a2.clave when 1 then a2.monto * 0.09     when 70 then a2.monto * 0.09   end as obrero,  case a2.clave" +
             " when 1 then a2.monto * 0.185 when 70 then a2.monto * 0.185  end as patronal FROM maestro a1 left JOIN nominew a2 ON a1.jpp = a2.jpp " +
               " AND a1.num = a2.numjpp where a2.clave in (1,70) GROUP BY  a1.rfc,	a1.nombre,	a1.cvecat,	a1.categ,   a2.clave,   a2.monto";
            }
            List<Dictionary<string, object>> lista = globales.web(query);

            string archivoTemp = Path.GetTempFileName();
            using (FileStream fs = new FileStream(archivoTemp, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                DotNetDBF.DBFWriter escribir = new DotNetDBF.DBFWriter();
                 archivoTemp.Replace("dbf", "dbt");

                DotNetDBF.DBFField c1 = new DotNetDBF.DBFField("Rfc", DotNetDBF.NativeDbType.Char,20 );
                DotNetDBF.DBFField c2 = new DotNetDBF.DBFField("Nombre", DotNetDBF.NativeDbType.Char, 40);
                DotNetDBF.DBFField c3 = new DotNetDBF.DBFField("Sicve", DotNetDBF.NativeDbType.Char, 17);
                DotNetDBF.DBFField c4 = new DotNetDBF.DBFField("Sicatg", DotNetDBF.NativeDbType.Char, 40);
                DotNetDBF.DBFField c5 = new DotNetDBF.DBFField("sueldo", DotNetDBF.NativeDbType.Numeric, 11, 2);
                DotNetDBF.DBFField c6 = new DotNetDBF.DBFField("obrero", DotNetDBF.NativeDbType.Numeric, 11, 2);
                DotNetDBF.DBFField c7 = new DotNetDBF.DBFField("patron", DotNetDBF.NativeDbType.Numeric, 11, 2);


                DotNetDBF.DBFField[] campos = new DotNetDBF.DBFField[] { c1, c2, c3, c4, c5, c6, c7};
                escribir.Fields = campos;


                foreach(var item in lista)
                {
                    string nombre = Convert.ToString(item["nombre"]);
                    string rfc = Convert.ToString(item["rfc"]);
                    string cvecat = Convert.ToString(item["cvecat"]);
                    string categ = Convert.ToString(item["categ"]);
                    double sueldo = Convert.ToDouble(item["sueldo"]);
                    double obrero = Convert.ToDouble(item["obrero"]);
                    double patron = Convert.ToDouble(item["patronal"]);

                    List<object> record = new List<object> {
                                rfc , nombre , cvecat , categ , sueldo , obrero , patron
                            };

                    escribir.AddRecord(record.ToArray());
                }
                escribir.Write(fs);

                escribir.Close();
                fs.Close();
                byte[] bytes = System.IO.File.ReadAllBytes(archivoTemp);
                string base64 = Convert.ToBase64String(bytes);

                MemoryStream enmemoria = new MemoryStream(bytes);
                byte[] bytesInStream = enmemoria.ToArray();
                Response.Clear();
                Response.ContentType = "application/force-download";
                Response.AddHeader("content-disposition", "attachment;    filename=arch.dbf");
                Response.BinaryWrite(bytesInStream);
                Response.End();

            }
            return null;
        }
    }


    
}