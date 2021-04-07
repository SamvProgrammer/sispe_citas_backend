using System;
using System.Collections.Generic;
using System.Data.OleDb;
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
        public object[] auxiliar;

        public bool bandera = false;

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

            foreach (Dictionary<string, object> item in resultado)
            {
                string query2 = $"SELECT  hora FROM oficialia.control_citas where dia='{string.Format("{0:yyyy-MM-dd}", DateTime.Parse(Convert.ToString(item["dia"])))}' and id_tramite ={id_tramite} and id_depto ={id_depto} and status = false;";
                List<Dictionary<string, object>> resultado2 = globales.consulta(query2);
                item.Add("horas", resultado2);
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

            object[] headers = { "nombre", "apellido", "tramite", "fecha", "hora" };
            object[] body = { nombre, apellido, tramite, string.Format("{0:d}", DateTime.Parse(fecha)), hora };
            parameters[0] = headers;
            parameters[1] = body;

            object[] cuerpo = new object[1];
            object[] tt1 = { "", "", "" };

            cuerpo[0] = tt1;



            byte[] pdfBytes = globales.reportes("reporte1", "usuario", cuerpo, "", false, parameters);
            string base64pdf = Convert.ToBase64String(pdfBytes);

            correoElectronico objCorreo = new correoElectronico();

            objCorreo.enviarCorreo(base64pdf, correo);

            Dictionary<string, string> rsultado = new Dictionary<string, string>();
            rsultado.Add("resultado", "enviado al correo");

            return Json(rsultado, JsonRequestBehavior.AllowGet);
        }






        //////// A PARTIR DE AQUÍ ES CÓDIGO DEL SISPE WEB


        [HttpGet]
        public HttpRequestMessage GeneraResumen(string emplea, string tiponomina, string descripcion, string cve)
        {

            string distinto = $"SELECT DISTINCT (rfc),a1.num FROM maestro a1 JOIN nominew a2 ON a1.num = a2.numjpp AND a1.jpp = a2.jpp WHERE a1.jpp = a2.jpp AND a1.jpp = '{emplea}' AND a2.tiponomina = '{tiponomina}' ORDER BY   a1.num ";
            List<Dictionary<string, object>> listaaa = globales.web(distinto);
            int cantidad = listaaa.Count;
            // Resumen de Nómina
            string queryresumen = $"SELECT	a2.clave,	a2.tipopago,	MAX (a3.descripcion) AS descri,	SUM (monto) AS monto FROM	maestro a1 LEFT JOIN nominew a2 ON a1.jpp = a2.jpp JOIN perded a3 ON a3.clave = a2.clave AND a1.num = a2.numjpp"
                + $" WHERE a1.jpp = '{emplea}' AND a2.tiponomina = '{tiponomina}' GROUP BY a2.clave, a2.tipopago ORDER BY a2.clave,a2.tipopago; ";


            List<Dictionary<string, object>> resumen = globales.web(queryresumen);

            if (resumen.Count <= 0) return null;
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

            string periodo = descripcion;
            string cve_p = cve;

            double operacion = Convert.ToDouble(sumap) - Convert.ToDouble(sumad);
            string letra = $"LA PRESENTE NOMINA AMPARA LA CANTIDAD DE:{globales.convertirNumerosLetras(Convert.ToString(operacion), true)}";
            object[] parametrosR = { "descripcion", "sumap", "sumad", "liquido", "conteo", "letra", "cve" };
            object[] valorR = { periodo, string.Format("{0:C}", Convert.ToDouble(sumap)), string.Format("{0:C}", Convert.ToDouble(sumad)), string.Format("{0:C}", Convert.ToDouble(sumap) - Convert.ToDouble(sumad)), Convert.ToString(cantidad), letra, cve_p };
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
            Response.AddHeader("content-disposition", "attachment;    filename=resumen_nomina.pdf");
            Response.BinaryWrite(bytesInStream2);
            Response.End();


            return null;
        }













        [HttpGet]
        public JsonResult GeneraNomina(string emplea, string tiponomina, string descripcion, string cve)
        {

            string query = $"SELECT	a1.categ,CONCAT (a1.jpp, a1.num) AS proyecto,a1.depto,a1.nombre,a1.curp,	a1.rfc,	a2.clave,	a2.descri,	a2.monto,	a2.pagon,	a2.pagot" +
             $" FROM maestro a1 JOIN nominew a2 ON a1.num = a2.numjpp AND a1.jpp = a2.jpp WHERE a1.jpp = a2.jpp AND a1.jpp = '{emplea}' AND a2.tiponomina = '{tiponomina}' ORDER BY   a1.jpp,	a1.num,	a2.clave";
            List<Dictionary<string, object>> resultado = globales.web(query);

            if (resultado.Count <= 0) return null;

            //query = "select  clave,descripcion from perded order by clave";
            //List<Dictionary<string, object>> perded = globales.web(query);

            //resultado.ForEach(o =>
            //{
            //    o["descripcion"] = perded.Where(p => Convert.ToString(o["clave"]) == Convert.ToString(p["clave"])).First()["descripcion"];
            //    //  o["descri"] += " (RETROACTIVO)";
            //});

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



            string periodo = descripcion;
            string clav_pe = cve;
            string fecha = "01/02/2020";
            cve = "545645464";

            object[] parametros = { "descripcion", "fecha", "cve_presupuestal" };
            object[] valor = { periodo, fecha, clav_pe };
            object[][] enviarParametros = new object[2][];

            enviarParametros[0] = parametros;
            enviarParametros[1] = valor;

            byte[] bytes = globales.reportes("reporteGeneracionNominas", "nomina", aux2, "", true, enviarParametros, true);

            string base64 = Convert.ToBase64String(bytes);
            MemoryStream enmemoria = new MemoryStream(bytes);
            byte[] bytesInStream = enmemoria.ToArray();
            Response.Clear();
            Response.ContentType = "application/force-download";
            Response.AddHeader("content-disposition", "attachment;    filename=desglose_nomina.pdf");
            Response.BinaryWrite(bytesInStream);
            Response.End();

            Dictionary<string, string> rsultado = new Dictionary<string, string>();
            rsultado.Add("resultado", "enviado al correo");

            return Json(rsultado, JsonRequestBehavior.AllowGet);

        }

        [HttpGet]
        public HttpResponseMessage ProductosNominaQuincenal(String quincena)
        {
            string query = string.Empty;
            if (Convert.ToInt32(quincena) == 09 || Convert.ToInt32(quincena) == 11)
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
            if (lista.Count <= 0) return null;

            string archivoTemp = Path.GetTempFileName();
            using (FileStream fs = new FileStream(archivoTemp, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                DotNetDBF.DBFWriter escribir = new DotNetDBF.DBFWriter();
                archivoTemp.Replace("dbf", "dbt");

                DotNetDBF.DBFField c1 = new DotNetDBF.DBFField("Rfc", DotNetDBF.NativeDbType.Char, 20);
                DotNetDBF.DBFField c2 = new DotNetDBF.DBFField("Nombre", DotNetDBF.NativeDbType.Char, 40);
                DotNetDBF.DBFField c3 = new DotNetDBF.DBFField("Sicve", DotNetDBF.NativeDbType.Char, 17);
                DotNetDBF.DBFField c4 = new DotNetDBF.DBFField("Sicatg", DotNetDBF.NativeDbType.Char, 40);
                DotNetDBF.DBFField c5 = new DotNetDBF.DBFField("sueldo", DotNetDBF.NativeDbType.Numeric, 11, 2);
                DotNetDBF.DBFField c6 = new DotNetDBF.DBFField("obrero", DotNetDBF.NativeDbType.Numeric, 11, 2);
                DotNetDBF.DBFField c7 = new DotNetDBF.DBFField("patron", DotNetDBF.NativeDbType.Numeric, 11, 2);


                DotNetDBF.DBFField[] campos = new DotNetDBF.DBFField[] { c1, c2, c3, c4, c5, c6, c7 };
                escribir.Fields = campos;


                foreach (var item in lista)
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
                Response.AddHeader("content-disposition", "attachment;    filename=p_quincenal.dbf");
                Response.BinaryWrite(bytesInStream);
                Response.End();

            }
            return null;
        }


        [HttpGet]
        public HttpResponseMessage GeneraAportas(string desde, string hasta)
        {
            string query = string.Empty;

            try
            {
                DateTime f_desde = Convert.ToDateTime(desde);
                DateTime f_hasta = Convert.ToDateTime(hasta);
                int mes = f_desde.Month;
                if (mes == 5 || mes == 12)
                {

                    query = $"CREATE TEMP table temporas as  select a2.nombre,a2.rfc,a2.cvecat , a2.proyecto,a2.jpp ,a1.numjpp from nominew a1 INNER JOIN maestro a2 on a1.jpp=a2.jpp and a1.numjpp = a2.num where a1.clave = 202; " +
                       " create temp table tmp as select a1.nombre , a1.rfc ,a1.cvecat ,a1.proyecto, (a2.monto * 2) as sueldo , (a2.monto *2 ) * 0.09 as aportacion , concat('N') as tipoaporta   from temporas a1 INNER JOIN nominew a2 on a1.jpp = a2.jpp and a1.numjpp = a2.numjpp where clave in (1, 70); " +
                   " select* from tmp";
                }
                else
                {
                    query = $"CREATE TEMP table temporas as  select a2.nombre,a2.rfc,a2.cvecat , a2.proyecto,a2.jpp ,a1.numjpp from nominew a1 INNER JOIN maestro a2 on a1.jpp=a2.jpp and a1.numjpp = a2.num where a1.clave = 202; " +
                         " create temp table tmp as select a1.nombre , a1.rfc ,a1.cvecat ,a1.proyecto, a2.monto as sueldo , a2.monto * 0.09 as aportacion , concat('N') as tipoaporta   from temporas a1 INNER JOIN nominew a2 on a1.jpp = a2.jpp and a1.numjpp = a2.numjpp where clave in (1, 70); " +
                     " select* from tmp";
                }


            }

            catch
            {
                return null;

            }


            List<Dictionary<string, object>> resultado = globales.web(query);
            if (resultado.Count <= 0) return null;

            string archivoTemp = Path.GetTempFileName();
            using (FileStream fs = new FileStream(archivoTemp, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                DotNetDBF.DBFWriter escribir = new DotNetDBF.DBFWriter();
                archivoTemp.Replace("dbf", "dbt");

                DotNetDBF.DBFField c1 = new DotNetDBF.DBFField("Nombre", DotNetDBF.NativeDbType.Char, 50);
                DotNetDBF.DBFField c2 = new DotNetDBF.DBFField("RFC", DotNetDBF.NativeDbType.Char, 20);
                DotNetDBF.DBFField c3 = new DotNetDBF.DBFField("Proyecto", DotNetDBF.NativeDbType.Char, 40);
                DotNetDBF.DBFField c4 = new DotNetDBF.DBFField("Categoria", DotNetDBF.NativeDbType.Char, 23);
                DotNetDBF.DBFField c5 = new DotNetDBF.DBFField("sueldo", DotNetDBF.NativeDbType.Numeric, 11, 2);
                DotNetDBF.DBFField c6 = new DotNetDBF.DBFField("Aportacion", DotNetDBF.NativeDbType.Numeric, 11, 2);
                DotNetDBF.DBFField c7 = new DotNetDBF.DBFField("Tipoaporta", DotNetDBF.NativeDbType.Char,2);
                DotNetDBF.DBFField c8 = new DotNetDBF.DBFField("desde", DotNetDBF.NativeDbType.Date);
                DotNetDBF.DBFField c9 = new DotNetDBF.DBFField("hasta", DotNetDBF.NativeDbType.Date);




                DotNetDBF.DBFField[] campos = new DotNetDBF.DBFField[] { c1, c2, c3, c4, c5, c6, c7, c8, c9 };
                escribir.Fields = campos;



                foreach (var item in resultado)
                {
                    string nombre = Convert.ToString(item["nombre"]);
                    string rfc = Convert.ToString(item["rfc"]);
                    string proyecto = Convert.ToString(item["proyecto"]);
                    string categoria = Convert.ToString(item["cvecat"]);
                    string sueldo = Convert.ToString(item["sueldo"]);
                    string aportacion = Convert.ToString(item["aportacion"]);
                    string tipoaporta = Convert.ToString(item["tipoaporta"]);




                    List<object> record = new List<object> {
                                nombre ,rfc, proyecto , categoria , sueldo , aportacion ,  tipoaporta, Convert.ToDateTime(desde) , Convert.ToDateTime(hasta)
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
                Response.AddHeader("content-disposition", "attachment;    filename=A490.dbf");
                Response.BinaryWrite(bytesInStream);
                Response.End();

            }
            return null;



        }




        [HttpGet]
        public HttpRequestMessage GenerarDeducciones(String desde, string hasta)
        {
            string query = string.Empty;

            DateTime f_desde = Convert.ToDateTime(desde);
            DateTime f_hasta = Convert.ToDateTime(hasta);
            int mes = f_desde.Month;

            try
            {
              
                    query = $"CREATE TEMP TABLE temporas AS SELECT	a2.nombre,	a2.rfc,	a2.proyecto,	a1.clave,	a1.folio,	a1.pagon,	a1.pagot,	a1.monto, a1.jpp,a1.numjpp FROM	nominew a1 INNER JOIN maestro a2 ON a1.jpp = a2.jpp AND a1.numjpp = a2.num;" +
             "CREATE TEMP TABLE tmp AS SELECT a1.nombre,a1.rfc,a1.proyecto,a1.clave,a1.folio,a1.pagon,a1.pagot,a1.monto, a1.jpp,a1.numjpp, concat('N') AS tipoaporta FROM temporas a1 INNER JOIN nominew a2 ON a1.jpp = a2.jpp" +
              " bAND a1.numjpp = a2.numjpp WHERE a2.clave IN (205, 206); SELECT* FROM tmp";
                
            }
            catch
            {
                return null;
            }



            List<Dictionary<string, object>> resultado = globales.web(query);
            if (resultado.Count <= 0) return null;


            string archivoTemp = Path.GetTempFileName();
            using (FileStream fs = new FileStream(archivoTemp, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                DotNetDBF.DBFWriter escribir = new DotNetDBF.DBFWriter();
                archivoTemp.Replace("dbf", "dbt");

                DotNetDBF.DBFField c1 = new DotNetDBF.DBFField("Nombre", DotNetDBF.NativeDbType.Char, 40);
                DotNetDBF.DBFField c2 = new DotNetDBF.DBFField("RFC", DotNetDBF.NativeDbType.Char, 20);
                DotNetDBF.DBFField c3 = new DotNetDBF.DBFField("Proyecto", DotNetDBF.NativeDbType.Char, 40);
                DotNetDBF.DBFField c4 = new DotNetDBF.DBFField("clave", DotNetDBF.NativeDbType.Numeric, 11);
                DotNetDBF.DBFField c5 = new DotNetDBF.DBFField("folio", DotNetDBF.NativeDbType.Numeric, 11, 2);
                DotNetDBF.DBFField c6 = new DotNetDBF.DBFField("numdesc", DotNetDBF.NativeDbType.Numeric, 11);
                DotNetDBF.DBFField c7 = new DotNetDBF.DBFField("totdesc", DotNetDBF.NativeDbType.Numeric, 11);
                DotNetDBF.DBFField c8 = new DotNetDBF.DBFField("importe", DotNetDBF.NativeDbType.Numeric, 11, 2);
                DotNetDBF.DBFField c9 = new DotNetDBF.DBFField("tipodesc", DotNetDBF.NativeDbType.Char, 1);
                DotNetDBF.DBFField c10 = new DotNetDBF.DBFField("desde", DotNetDBF.NativeDbType.Date);
                DotNetDBF.DBFField c11 = new DotNetDBF.DBFField("hasta", DotNetDBF.NativeDbType.Date);




                DotNetDBF.DBFField[] campos = new DotNetDBF.DBFField[] { c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11 };
                escribir.Fields = campos;



                foreach (var item in resultado)
                {
                    string nombre = Convert.ToString(item["nombre"]);
                    string rfc = Convert.ToString(item["rfc"]);
                    string proyecto = Convert.ToString(item["proyecto"]);
                    string clave = Convert.ToString(item["clave"]);
                    string folio = Convert.ToString(item["folio"]);
                    string numdesc = Convert.ToString(item["pagon"]);
                    string totdesc = Convert.ToString(item["pagon"]);
                    double montot = Convert.ToDouble(item["montot"]);


                    string tipoaporta = Convert.ToString(item["tipoaporta"]);




                    List<object> record = new List<object> {
                                nombre ,rfc, proyecto , clave , folio ,numdesc , totdesc,montot,'N',Convert.ToDateTime( desde) , Convert.ToDateTime(hasta)
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
                Response.AddHeader("content-disposition", "attachment;    filename=D490.dbf");
                Response.BinaryWrite(bytesInStream);
                Response.End();

            }
            return null;

        }



        [HttpGet]
       public HttpResponseMessage AvanzarSerie(string desde , string hasta)
        {

            string query = string.Empty;

            DateTime f_desde = Convert.ToDateTime(desde);
            DateTime f_hasta = Convert.ToDateTime(hasta);
            int mes = f_desde.Month;
            if (mes == 5 || mes == 12)
            {

                query = "UPDATE nominew SET pagon = ( 	CASE	WHEN (pagot - pagon) = 1 THEN		pagon + 1	ELSE		pagon + 2	END) ," +
                       " monto = (CASE WHEN(pagot - pagon) = 1 THEN monto  ELSE monto *2 END ) WHERE(pagon <= pagot) AND pagon <> 0  " +
                       " AND pagot <> 0 AND tiponomina='N'; DELETE FROM    nominew WHERE   pagon > pagot AND tiponomina='N'; update nominew set monto = monto * 2 where pagon is null and tiponomina = 'N' ";

            }
            else
            {
                query = "update nominew  set pagon = pagon  where(pagon <= pagot)and pagon<> 0 and pagot<> 0 and tiponomina='N'; DELETE FROM    nominew WHERE   pagon > pagot and tiponomina='N';";

            }

            globales.web(query);


            return null;
        }


        [HttpGet]
      public  HttpResponseMessage DividirPrestamos()
        {
            string query = "update nominew set monto= monto * 2 where pagon is null and tiponomina='N'";
            globales.web(query);
            return null;
        }




        [HttpGet]
        public HttpResponseMessage DbfRh(string archivo , string tiponomina)
        {

           string query = "drop table base;" +
$"CREATE TEMP TABLE t1 AS SELECT	JPP,	numjpp FROM	respaldos_nominew WHERE 	tiponomina = '{tiponomina}' AND archivo = '{archivo}' AND tipopago = 'N' GROUP BY	jpp,	numjpp; CREATE TEMP TABLE sacar AS SELECT	JPP,	numjpp, "+
"	SUM(monto) filter (where clave =1) as P1,SUM(monto) filter (where clave =2) as P2,SUM(monto) filter (where clave =3) as P3,SUM(monto) filter (where clave =4) as P4,SUM(monto) filter (where clave =5) as P5,SUM(monto) filter (where clave =6) as P6, " +
"SUM(monto) filter (where clave =7) as P7,SUM(monto) filter (where clave =8) as P8,SUM(monto) filter (where clave =9) as P9,  " +
"SUM(monto) filter (where clave =10) as P10,SUM(monto) filter (where clave =12) as P12,SUM(monto) filter (where clave =13) as P13,SUM(monto) filter (where clave =14) as P14,SUM(monto) filter (where clave =16) as P16, " +
"SUM(monto) filter (where clave =18) as P18,SUM(monto) filter (where clave =19) as P19,SUM(monto) filter (where clave =22) as P22,SUM(monto) filter (where clave =23) as P23,SUM(monto) filter (where clave =25) as P25, " +
"SUM(monto) filter (where clave =26) as P26,SUM(monto) filter (where clave =27) as P27,SUM(monto) filter (where clave =28) as P28,SUM(monto) filter (where clave =29) as P29,SUM(monto) filter (where clave =30) as P30, " +
"SUM(monto) filter (where clave =31) as P31,SUM(monto) filter (where clave =37) as P37,SUM(monto) filter (where clave =38) as P38,SUM(monto) filter (where clave =40) as P40,SUM(monto) filter (where clave =41) as P41, " +
"SUM(monto) filter (where clave =42) as P42,SUM(monto) filter (where clave =43) as P43,SUM(monto) filter (where clave =44) as P44,SUM(monto) filter (where clave =45) as P45,SUM(monto) filter (where clave =46) as P46, " +
"SUM(monto) filter (where clave =47) as P47,SUM(monto) filter (where clave =48) as P48,SUM(monto) filter (where clave =49) as P49,SUM(monto) filter (where clave =55) as P55,SUM(monto) filter (where clave =56) as P56, " +
"SUM(monto) filter (where clave =57) as P57,SUM(monto) filter (where clave =58) as P58,SUM(monto) filter (where clave =59) as P59,SUM(monto) filter (where clave =60) as P60,SUM(monto) filter (where clave =61) as P61, " +
"SUM(monto) filter (where clave =62) as P62,SUM(monto) filter (where clave =70) as P70,SUM(monto) filter (where clave =71) as P71,SUM(monto) filter (where clave =72) as P72,SUM(monto) filter (where clave =73) as P73, " +
"SUM(monto) filter (where clave =74) as P74,SUM(monto) filter (where clave =75) as P75,SUM(monto) filter (where clave =76) as P76,SUM(monto) filter (where clave =77) as P77,SUM(monto) filter (where clave =78) as P78, " +
"SUM(monto) filter (where clave =101) as P101,SUM(monto) filter (where clave =104) as P104,SUM(monto) filter (where clave =105) as P105,SUM(monto) filter (where clave =106) as P106,SUM(monto) filter (where clave =107) as P107, " +
"SUM(monto) filter (where clave =108) as P108,SUM(monto) filter (where clave =109) as P109,SUM(monto) filter (where clave =110) as P110,SUM(monto) filter (where clave =125) as P125,SUM(monto) filter (where clave =127) as P127,SUM(monto) filter (where clave =134) as P134, " +
"SUM(monto) filter (where clave =136) as P136,SUM(monto) filter (where clave =137) as P137,SUM(monto) filter (where clave =201) as D201,SUM(monto) filter (where clave =202) as D202,SUM(monto) filter (where clave =203) as D203, " +
"SUM(monto) filter (where clave =204) as D204,SUM(monto) filter (where clave =205) as D205,SUM(monto) filter (where clave =206) as D206,SUM(monto) filter (where clave =208) as D208,SUM(monto) filter (where clave =209) as D209, " +
"SUM(monto) filter (where clave =210) as D210,SUM(monto) filter (where clave =211) as D211,SUM(monto) filter (where clave =212) as D212,SUM(monto) filter (where clave =213) as D213,SUM(monto) filter (where clave =214) as D214,SUM(monto) filter (where clave =215) as D215, " +
"SUM(monto) filter (where clave =217) as D217,SUM(monto) filter (where clave =219) as D219,SUM(monto) filter (where clave =220) as D220,SUM(monto) filter (where clave =221) as D221,SUM(monto) filter (where clave =223) as D223, " +
"SUM(monto) filter (where clave =226) as D226,SUM(monto) filter (where clave =227) as D227,SUM(monto) filter (where clave =229) as D229,SUM(monto) filter (where clave =230) as D230,SUM(monto) filter (where clave =231) as D231,SUM(monto) filter (where clave =233) as D233, " +
"SUM(monto) filter (where clave =236) as D236,SUM(monto) filter (where clave =237) as D237,SUM(monto) filter (where clave =241) as D241,SUM(monto) filter (where clave =242) as D242,SUM(monto) filter (where clave =244) as D244,SUM(monto) filter (where clave =248) as D248, " +
"SUM(monto) filter (where clave =250) as D250,SUM(monto) filter (where clave =251) as D251,SUM(monto) filter (where clave =252) as D252,SUM(monto) filter (where clave =254) as D254,SUM(monto) filter (where clave =255) as D255,SUM(monto) filter (where clave =256) as D256, " +
"SUM(monto) filter (where clave =257) as D257,SUM(monto) filter (where clave =258) as D258,SUM(monto) filter (where clave =259) as D259,SUM(monto) filter (where clave =260) as D260,SUM(monto) filter (where clave =263) as D263,SUM(monto) filter (where clave =264) as D264, " +
"SUM(monto) filter (where clave =265) as D265,SUM(monto) filter (where clave =268) as D268,SUM(monto) filter (where clave =269) as D269,SUM(monto) filter (where clave =274) as D274,SUM(monto) filter (where clave =304) as D304,SUM(monto) filter (where clave =306) as D306, " +
"SUM(monto) filter (where clave =308) as D308,SUM(monto) filter (where clave =314) as D314,SUM(monto) filter (where clave =315) as D315,SUM(monto) filter (where clave =316) as D316,SUM(monto) filter (where clave =317) as D317,SUM(monto) filter (where clave =319) as D319, " +
"SUM(monto) filter (where clave =322) as D322,SUM(monto) filter (where clave =323) as D323,SUM(monto) filter (where clave =331) as D331 FROM	respaldos_nominew WHERE	tiponomina = 'N' AND archivo = '2107' AND tipopago = 'N' GROUP BY jpp,numjpp,clave; " +
"CREATE TEMP table sumas as  SELECT	jpp,	numjpp,	sum(p1) as p1,sum(p2) as p2,sum(p3) as p3,sum(p4) as p4,sum(p5) as p5,sum(p6) as p6,sum(p7) as p7,sum(p8) as p8,sum(p9) as p9,sum(p10) as p10,sum(p12) as p12,sum(p13) as p13,sum(p14) as p14,sum(p16) as p16,sum(p18) as p18,sum(p19) as p19, " +
"sum(p22) as p22,sum(p23) as p23,sum(p25) as p25,sum(p26) as p26,sum(p27) as p27,sum(p28) as p28,sum(p29) as p29,sum(p30) as p30,sum(p31) as p31,sum(p37) as p37,sum(p38) as p38,sum(p40) as p40,sum(p41) as p41,sum(p42) as p42,sum(p43) as p43,sum(p44) as p44,sum(p45) as p45,sum(p46) as p46, " +
"sum(p47) as p47,sum(p48) as p48,sum(p49) as p49,sum(p55) as p55,sum(p56) as p56,sum(p57) as p57,sum(p58) as p58,sum(p59) as p59,sum(p60) as p60,sum(p61) as p61,sum(p62) as p62,sum(p70) as p70,sum(p71) as p71,sum(p72) as p72,sum(p73) as p73,sum(p74) as p74,sum(p75) as p75,sum(p76) as p76, " +
"sum(p77) as p77,sum(p78) as p78,sum(p101) as p101,sum(p104) as p104,sum(p105) as p105,sum(p106) as p106,sum(p107) as p107,sum(p108) as p108,sum(p109) as p109,sum(p110) as p110,sum(p125) as p125,sum(p127) as p127, " +
"sum(p134) as p134,sum(p136) as p136,sum(p137) as p137,sum(d201) as d201,sum(d202) as d202,sum(d203) as d203,sum(d204) as d204,sum(d205) as d205,sum(d206) as d206,sum(d208) as d208,sum(d209) as d209,sum(d210) as d210, " +
"sum(d211) as d211,sum(d212) as d212,sum(d213) as d213,sum(d214) as d214,sum(d215) as d215,sum(d217) as d217,sum(d219) as d219,sum(d220) as d220,sum(d221) as d221,sum(d223) as d223,sum(d226) as d226,sum(d227) as d227, " +
"sum(d229) as d229,sum(d230) as d230,sum(d231) as d231,sum(d233) as d233,sum(d236) as d236,sum(d237) as d237,sum(d241) as d241,sum(d242) as d242,sum(d244) as d244,sum(d248) as d248,sum(d250) as d250,sum(d251) as d251,sum(d252) as d252,sum(d254) as d254, " +
"sum(d255) as d255,sum(d256) as d256,sum(d257) as d257,sum(d258) as d258,sum(d259) as d259,sum(d260) as d260,sum(d263) as d263,sum(d264) as d264,sum(d265) as d265,sum(d268) as d268,sum(d269) as d269,sum(d274) as d274, " +
"sum(d304) as d304,sum(d306) as d306,sum(d308) as d308,sum(d314) as d314,sum(d315) as d315,sum(d316) as d316,sum(d317) as d317,sum(d319) as d319,sum(d322) as d322,sum(d323) as d323,sum(d331) as d331  " +
"FROM sacar GROUP BY	jpp,	numjpp ORDER BY	numjpp; " +
"CREATE  TABLE base AS SELECT concat (0) as cqnacve,	concat (1) AS cqnaind,	concat (0) AS pbpnup,	concat (0) AS pbpnue,	substr(mma.rfc, 1, 10) AS sirfc,	mma.rfc AS rfchomo,	mma.nombre AS sinom, " +
"	mma.proyecto AS sidep,	mma.cvecat AS sicatg,mma.sexo AS sexo,CASE WHEN (mma.jpp = 'DIB') THEN 	1 WHEN (mma.jpp = 'COC') THEN 	6 WHEN (mma.jpp = 'MMS') THEN 	4 WHEN (mma.JPP = 'NOC') THEN 	3 " +
"END AS tip_emp, mma.indisindi, concat (0) AS numfolio, concat (3) AS ubica, concat (470) AS cve, concat (0) AS tipopago, mma.fchinggob AS pbpfig, mma.fpuesto AS pbpipto, mma.fchingnom AS pbpfnom, mma.status AS pbpstatus, " +
" mma.licedes AS licdes, mma.licehas AS lichas, mma.mot AS cmotcve, mma.hijos AS pbphijos, mma.gufdes AS guades, mma.gufhas AS guahas, mma.curp AS curp, mma.depto AS areads, mma.nombanco,mma.numcuenta, " +
" concat (0) AS clave_inte, mma.nivest AS cestniv, mma.gdoest AS cestgdo, mma.numquin AS qnios, concat (0) AS nfalta, concat (0) AS nretar,mma.imss AS pbpimss,mma.bgimss,mma.uimss AS cuopimss,concat (0) AS cuoprcv,mma.infdes AS incdes, " +
" mma.infhas AS inchas,concat (0) AS cuopinf,mma.fonpen AS cuopfpen,mma.bgimss AS bgispt,mma.profesion,mma.statpago AS statpago,mma.aguifdes,mma.aguifhas,mma.dafdes,mma.dafhas,sumas.p1,sumas.p2, " +
"sumas.p3,sumas.p4,sumas.p5,sumas.p6,sumas.p7,sumas.p8,sumas.p9,sumas.p10,sumas.p12,sumas.p13,sumas.p14,sumas.p16,sumas.p18,sumas.p19,sumas.p22,sumas.p23,sumas.p25,sumas.p26,sumas.p27,sumas.p28,sumas.p29, " +
"sumas.p30,sumas.p31,sumas.p37,sumas.p38,sumas.p40,sumas.p41,sumas.p42,sumas.p43,sumas.p44,sumas.p45,sumas.p46,sumas.p47,sumas.p48,sumas.p49,sumas.p55,sumas.p56,sumas.p57,sumas.p58,sumas.p59,sumas.p60, " +
"sumas.p61,sumas.p62,sumas.p70,sumas.p71,sumas.p72,sumas.p73,sumas.p74,sumas.p75,sumas.p76,sumas.p77,sumas.p78,sumas.p101,sumas.p104,sumas.p105,sumas.p106,sumas.p107,sumas.p108,sumas.p109,sumas.p110,sumas.p125, " +
"sumas.p127,sumas.p134,sumas.p136,sumas.p137,sumas.d201,sumas.d202,sumas.d203,sumas.d204,sumas.d205,sumas.d206,sumas.d208,sumas.d209,sumas.d210,sumas.d211,sumas.d212,sumas.d213,sumas.d214,sumas.d215,sumas.d217, " +
"sumas.d219,sumas.d220,sumas.d221,sumas.d223,sumas.d226,sumas.d227,sumas.d229,sumas.d230,sumas.d231,sumas.d233,sumas.d236,sumas.d237,sumas.d241,sumas.d242,sumas.d244,sumas.d248,sumas.d250,sumas.d251,sumas.d252, " +
"sumas.d254,sumas.d255,sumas.d256,sumas.d257,sumas.d258,sumas.d259,sumas.d260,sumas.d263,sumas.d264,sumas.d265,sumas.d268,sumas.d269,sumas.d274,sumas.d304,sumas.d306,sumas.d308,sumas.d314, " +
"sumas.d315,sumas.d316,sumas.d317,sumas.d319,sumas.d322,sumas.d323,sumas.d331 FROM	maestro mma INNER JOIN t1 ON mma.jpp = t1.jpp AND mma.num = t1.numjpp INNER JOIN sumas on sumas.jpp = mma.jpp  " +
"and sumas.numjpp = mma.num ORDER BY 	mma.jpp,	mma.num; ";

            globales.web(query);

            return null;
        }











         [HttpGet]
        public HttpResponseMessage Listado()
        {
            string query = "SELECT	rfc,	nombre,proyecto,pagon,	pagot,folio,monto , a2.clave FROM maestro a1 INNER JOIN nominew a2 ON a1.num = a2.numjpp AND a1.num = a2.numjpp where a2.clave in (221 , 226 , 227) ;";
            List<Dictionary<string, object>> resultado = globales.web(query);
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                return null;
            }

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "RFC";
            xlWorkSheet.Cells[1, 2] = "NOMBRE";
            xlWorkSheet.Cells[1, 3] = "PROYECTO";
            xlWorkSheet.Cells[1, 4] = "PAGON";
            xlWorkSheet.Cells[1, 5] = "PAGOT";
            xlWorkSheet.Cells[1, 6] = "FOLIO";
            xlWorkSheet.Cells[1, 7] = "MONTO";
            xlWorkSheet.Cells[1, 8] = "CLAVE";


            int c = 2;
            foreach(var item in resultado)
            {

                xlWorkSheet.Cells[c, 1] = Convert.ToString(item["rfc"]);
                xlWorkSheet.Cells[c, 2] = Convert.ToString(item["nombre"]);
                xlWorkSheet.Cells[c, 3] = Convert.ToString(item["proyecto"]);
                xlWorkSheet.Cells[c, 4] = Convert.ToString(item["pagon"]);
                xlWorkSheet.Cells[c, 5] = Convert.ToString(item["pagot"]);
                xlWorkSheet.Cells[c, 6] = Convert.ToString(item["folio"]);
                xlWorkSheet.Cells[c, 7] = Convert.ToString(item["monto"]);
                xlWorkSheet.Cells[c, 8] = Convert.ToString(item["clave"]);
                c++;

            }

            string archivoTemp = Path.GetTempFileName();
            xlWorkBook.SaveAs(archivoTemp, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            byte[] bytes = System.IO.File.ReadAllBytes(archivoTemp);
            string base64 = Convert.ToBase64String(bytes);

            MemoryStream enmemoria = new MemoryStream(bytes);
            byte[] bytesInStream = enmemoria.ToArray();
            Response.Clear();
            Response.ContentType = "application/force-download";
            Response.AddHeader("content-disposition", "attachment;    filename=listado.xls");
            Response.BinaryWrite(bytesInStream);
            Response.End();



            return null;
        }


        [HttpGet]
      public  HttpResponseMessage GenerarSobre(string anio, string quincena , string tiponomina)
        {
            anio = anio.Substring(2, 2);
            string archivo = anio + quincena;
            string periodo = quincena;
            switch (periodo)
            {
                case "01": periodo = $"PAGO DEL 01.ENE.{anio} AL 15.ENE.{anio}"; break;
                case "02": periodo = $"PAGO DEL 16.ENE.{anio} AL 30.ENE.{anio}"; break;
                case "03": periodo = $"PAGO DEL 01.FEB.{anio} AL 15.FEB.{anio}"; break;
                case "04": periodo = $"PAGO DEL 16.ENE.{anio} AL 28.FEB.{anio}"; break;
                case "05": periodo = $"PAGO DEL 01.MAR.{anio} AL 15.MAR.{anio}"; break;
                case "06": periodo = $"PAGO DEL 16.MAR.{anio} AL 31.MAR.{anio}"; break;
                case "07": periodo = $"PAGO DEL 01.ABR.{anio} AL 15.ABR.{anio}"; break;
                case "08": periodo = $"PAGO DEL 16.ABR.{anio} AL 30.ABR.{anio}"; break;
                case "09": periodo = $"PAGO DEL 01.MAY.{anio} AL 30.MAY.{anio}"; break;
                case "10": periodo = $"PAGO DEL 01.JUN.{anio} AL 15.JUN.{anio}"; break;
                case "11": periodo = $"PAGO DEL 16.JUN.{anio} AL 30.JUN.{anio}"; break;
                case "12": periodo = $"PAGO DEL 01.JUL.{anio} AL 15.JUL.{anio}"; break;
                case "13": periodo = $"PAGO DEL 16.JUL.{anio} AL 31.JUL.{anio}"; break;
                case "14": periodo = $"PAGO DEL 01.AGO.{anio} AL 15.AGO.{anio}"; break;
                case "15": periodo = $"PAGO DEL 16.AGO.{anio} AL 31.AGO.{anio}"; break;
                case "16": periodo = $"PAGO DEL 01.SEP.{anio} AL 15.SEP.{anio}"; break;
                case "17": periodo = $"PAGO DEL 16.SEP.{anio} AL 30.SEP.{anio}"; break;
                case "18": periodo = $"PAGO DEL 01.OCT.{anio} AL 15.OCT.{anio}"; break;
                case "19": periodo = $"PAGO DEL 16.OCT.{anio} AL 31.OCT.{anio}"; break;
                case "20": periodo = $"PAGO DEL 01.NOV.{anio} AL 15.NOV.{anio}"; break;
                case "21": periodo = $"PAGO DEL 16.NOV.{anio} AL 30.NOV.{anio}"; break;
                case "22": periodo = $"PAGO DEL 01.DIC.{anio} AL 31.DIC.{anio}"; break;
            }



            string query = "CREATE TEMP TABLE t1 AS SELECT a1.jpp,a1.num,a1.nombre,a1.rfc FROM maestro a1;" +
        "CREATE TEMP TABLE t2 AS SELECT a2.numjpp,a2.jpp,a2.clave,a2.descri,a2.monto,	a2.archivo,	a2.pagon,a2.pagot FROM" +
           $" respaldos_nominew a2 WHERE a2.archivo = '{archivo}' AND a2.tiponomina = 'N';" +
      "  SELECT concat (t1.num, ' ', t1.nombre) AS nombre, t1.rfc, t2.clave, t2.descri, t2.monto, t2.archivo, t2.pagon, t2.pagot, CASE WHEN t1.jpp = 'MMS' THEN  'MANDOS MEDIOS Y SUP.'" +
         "WHEN t1.jpp = 'DIB' THEN 'ADMTVO. DIRECC. BASE' WHEN t1.jpp = 'DIC' THEN    'ADMTVO. DIRECC. CONF' WHEN t1.jpp = 'NOC' THEN     'ADMTVO. NOMB. CONF' WHEN t1.jpp = 'COC' THEN " +
        "'ADMTVO. CONTR. CONF' END AS categoria FROM     t1 INNER JOIN t2 ON t1.num = t2.numjpp ORDER BY t1.nombre , t2.clave ";
                List<Dictionary<string, object>> listado = globales.web(query);
                if (listado.Count <= 0)
                {
                    return null;
                }

                if (bandera == false)
                {
                    this.auxiliar = new object [listado.Count];
                    bandera = true;
                }
                else
                {

                }
                int conta = 0;
         
                foreach (var i in listado)
                {
                   string nombre = string.Empty;
                   string rfc = string.Empty;
                   string clave = string.Empty;
                   string descri = string.Empty;
                   string monto = string.Empty;
                    string categoria = string.Empty;

                    string pagon = string.Empty;
                    string pagot = string.Empty;
                    try
                    {

                        nombre = Convert.ToString(i["nombre"]);
                        rfc = Convert.ToString(i["rfc"]);
                        categoria = Convert.ToString(i["categoria"]);
                        clave = Convert.ToString(i["clave"]);
                        descri = Convert.ToString(i["descri"]);
                        monto = string.Format("{0:C}", Convert.ToDouble(i["monto"])).Replace("$", "");
                        pagon = Convert.ToString(i["pagon"]);
                        pagot = Convert.ToString(i["pagot"]);

                    }
                    catch
                    {

                    }
                    object[] tt1 = { nombre, categoria, rfc, periodo, "", "", "", "", "", "", "", "", "" , "" };
                    if (Convert.ToInt32(clave) < 200)
                    {
                        if (this.auxiliar[conta] == null)
                        {
                            tt1[4] = clave + " "+ descri;
                            tt1[5] = descri;
                            tt1[8] = monto;
                            this.auxiliar[conta] = tt1;
                        }
                        else
                        {
                            object[] tmp = (object[])this.auxiliar[conta];
                            tmp[4] = clave +" "+ descri;
                            tmp[5] = descri;
                            tmp[8] = monto;
                        }

                    }
                    else
                    {

                        if (this.auxiliar[conta] == null)
                        {
                            tt1[6] = clave+ " "+descri;
                            tt1[7] = descri;
                            tt1[13] = (string.IsNullOrWhiteSpace(pagon) || pagot == "0") ? "" : $"{pagon}/{pagot}";
                            tt1[9] = monto;
                            this.auxiliar[conta] = tt1;
                        }
                        else
                        {
                            object[] tmp = (object[])this.auxiliar[conta];
                            tmp[6] = clave +" "+ descri;
                            tmp[7] = descri;
                            tmp[13] = (string.IsNullOrWhiteSpace(pagon) || pagon == "0") ? "" : $"{pagon}/{pagot}";
                            tmp[9] = monto;

                        }
                    
                    }

                    conta++;
                }

             

            


           



            object[] parametros = { "periodo" };
            object[] valor = { periodo }; ;

            object[][] enviarParametros = new object[2][];   //joe

            enviarParametros[0] = parametros;
            enviarParametros[1] = valor;


            byte[] bytes = globales.reportes("sobres", "sobres", this.auxiliar, "", true, enviarParametros, true);

            string base64 = Convert.ToBase64String(bytes);
            MemoryStream enmemoria = new MemoryStream(bytes);
            byte[] bytesInStream = enmemoria.ToArray();
            Response.Clear();
            Response.ContentType = "application/force-download";
            Response.AddHeader("content-disposition", "attachment;    filename=sobres.pdf");
            Response.BinaryWrite(bytesInStream);
            Response.End();

            return null;
        }






        [HttpGet]
        public HttpRequestMessage GeneraResumenHistorial(string emplea, string tiponomina, string descripcion, string cve , string archivo)
        {

            string distinto = $"SELECT DISTINCT (rfc),a1.num FROM maestro a1 JOIN nominew a2 ON a1.num = a2.numjpp AND a1.jpp = a2.jpp WHERE a1.jpp = a2.jpp AND a1.jpp = '{emplea}' AND a2.tiponomina = '{tiponomina}' ORDER BY   a1.num ";
            List<Dictionary<string, object>> listaaa = globales.web(distinto);
            int cantidad = listaaa.Count;
            // Resumen de Nómina
            string queryresumen = $"SELECT	a2.clave,	a2.tipopago,	MAX (a3.descripcion) AS descri,	SUM (monto) AS monto FROM	maestro a1 LEFT JOIN respaldos_nominew a2 ON a1.jpp = a2.jpp JOIN perded a3 ON a3.clave = a2.clave AND a1.num = a2.numjpp"
                + $" WHERE a1.jpp = '{emplea}' AND a2.tiponomina = '{tiponomina}' AND a2.archivo='{archivo}' GROUP BY a2.clave, a2.tipopago ORDER BY a2.clave,a2.tipopago; ";


            List<Dictionary<string, object>> resumen = globales.web(queryresumen);

            if (resumen.Count <= 0) return null;
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

            string periodo = descripcion;
            string cve_p = cve;

            double operacion = Convert.ToDouble(sumap) - Convert.ToDouble(sumad);
            string letra = $"LA PRESENTE NOMINA AMPARA LA CANTIDAD DE:{globales.convertirNumerosLetras(Convert.ToString(operacion), true)}";
            object[] parametrosR = { "descripcion", "sumap", "sumad", "liquido", "conteo", "letra", "cve" };
            object[] valorR = { periodo, string.Format("{0:C}", Convert.ToDouble(sumap)), string.Format("{0:C}", Convert.ToDouble(sumad)), string.Format("{0:C}", Convert.ToDouble(sumap) - Convert.ToDouble(sumad)), Convert.ToString(cantidad), letra, cve_p };
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
            Response.AddHeader("content-disposition", "attachment;    filename=resumen_nomina_historiu.pdf");
            Response.BinaryWrite(bytesInStream2);
            Response.End();


            return null;
        }




        [HttpGet]
        public JsonResult GeneraNominaHistorial(string emplea, string tiponomina, string descripcion, string cve , string archivo)
        {

            string query = $"SELECT	a1.categ,CONCAT (a1.jpp, a1.num) AS proyecto,a1.depto,a1.nombre,a1.curp,	a1.rfc,	a2.clave,	a2.descri,	a2.monto,	a2.pagon,	a2.pagot" +
             $" FROM maestro a1 JOIN respaldos_nominew a2 ON a1.num = a2.numjpp AND a1.jpp = a2.jpp WHERE a1.jpp = a2.jpp AND a1.jpp = '{emplea}' AND a2.tiponomina = '{tiponomina}' AND a2.archivo='{archivo}' ORDER BY   a1.jpp,	a1.num,	a2.clave";
            List<Dictionary<string, object>> resultado = globales.web(query);

            if (resultado.Count <= 0) return null;

       

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



            string periodo = descripcion;
            string clav_pe = cve;
            string fecha = "01/02/2020";
            cve = "545645464";

            object[] parametros = { "descripcion", "fecha", "cve_presupuestal" };
            object[] valor = { periodo, fecha, clav_pe };
            object[][] enviarParametros = new object[2][];

            enviarParametros[0] = parametros;
            enviarParametros[1] = valor;

            byte[] bytes = globales.reportes("reporteGeneracionNominas", "nomina", aux2, "", true, enviarParametros, true);

            string base64 = Convert.ToBase64String(bytes);
            MemoryStream enmemoria = new MemoryStream(bytes);
            byte[] bytesInStream = enmemoria.ToArray();
            Response.Clear();
            Response.ContentType = "application/force-download";
            Response.AddHeader("content-disposition", "attachment;    filename=desglose_nomina_historial.pdf");
            Response.BinaryWrite(bytesInStream);
            Response.End();

            Dictionary<string, string> rsultado = new Dictionary<string, string>();
            rsultado.Add("resultado", "enviado al correo");

            return Json(rsultado, JsonRequestBehavior.AllowGet);

        }


    }
}