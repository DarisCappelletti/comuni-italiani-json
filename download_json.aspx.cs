using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Formatters.Binary;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace PortFolio.comuni_italiani_json
{
    public partial class download_json : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            RicercaComuniIstat();
        }

        private void RicercaComuniIstat()
        {
            // Verifico se sono stati richiesti i dettagli
            string richiestaDettagliIndicePa = Request.QueryString["dettagliIndicePa"];

            // Inizializzo la lista
            var list = new List<ComuneValore>();

            // Imposto lo stream
            Stream streamXls = null;
            
            using (var wc = new WebClient())
            {
                // Estrapolo i dati dall'url
                streamXls = wc.OpenRead("https://www.istat.it/storage/codici-unita-amministrative/Elenco-comuni-italiani.xls");
            }

            using (var mss = new MemoryStream())
            {
                streamXls.CopyTo(mss);
                mss.Position = 0;

                //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(mss);

                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                //IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                var i = 0;

                //5. Data Reader methods
                while (excelReader.Read())
                {
                    if (i != 0)
                    {
                        // Imposto l'oggetto
                        var obj = new ComuneValore
                        {
                            CodRegione = excelReader.GetValue(0),
                            CodUnitaTerritoriale = excelReader.GetValue(1),
                            CodProvinciaStorico = excelReader.GetValue(2),
                            CodProgressivoComune = excelReader.GetValue(3),
                            CodComuneAlfanumerico = excelReader.GetValue(4),
                            DenominazioneUniversale = excelReader.GetString(5),
                            DenominazioneItaliana = excelReader.GetString(6),
                            DenominazioneAltraLingua = excelReader.GetValue(7),
                            CodiceRipartizioneGeografica = excelReader.GetValue(8),
                            RipartizioneGeografica = excelReader.GetValue(9),
                            DenominazioneRegione = excelReader.GetValue(10),
                            DenominazioneUnitaTerritorialeSovracomunale = excelReader.GetValue(11),
                            TipologiaUnitaTerritorialeSovracomunale = excelReader.GetValue(12),
                            Capoluogo_metropolitana_liberoConsorzio = excelReader.GetValue(13),
                            SiglaAutomobilistica = excelReader.GetValue(14),
                            CodComuneNumerico = excelReader.GetValue(15),
                            CodComuneNumerico_110Provincie_2010_2016 = excelReader.GetValue(16),
                            CodComuneNumerico_107Provincie_2006_2009 = excelReader.GetValue(17),
                            CodComuneNumerico_103Provincie_1995_2005 = excelReader.GetValue(18),
                            CodCatastale = excelReader.GetValue(19),
                            CodNuts1_2010 = excelReader.GetValue(20),
                            CodNuts2_2010 = excelReader.GetValue(21),
                            CodNuts3_2010 = excelReader.GetValue(22),
                            CodNuts1_2021 = excelReader.GetValue(23),
                            CodNuts2_2021 = excelReader.GetValue(24),
                            CodNuts3_2021 = excelReader.GetValue(25)
                        };

                        // Scarico i dettagli da indicePA
                        var dettagliEnte = richiestaDettagliIndicePa == "true" ? RicercaEnte(obj.DenominazioneItaliana) : null;

                        obj.DettagliEnte = dettagliEnte;

                        list.Add(obj);
                    }

                    i++;
                }

                excelReader.Close();

                // Serializzo la lista
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                serializer.MaxJsonLength = Int32.MaxValue;
                var json = serializer.Serialize(list);

                BinaryFormatter bf = new BinaryFormatter();
                MemoryStream ms = new MemoryStream();
                bf.Serialize(ms, json);

                // Preparo il file
                string FileName = "ComuniItaliani_" + DateTime.Now + ".json";
                HttpResponse response = HttpContext.Current.Response;
                response.Clear();
                response.Charset = "";
                response.ContentType = "application/json";
                response.ContentEncoding = Encoding.UTF8;
                response.AddHeader("Content-Disposition", "attachment;filename=" + FileName);
                response.Write(json);
                response.End();
            }
        }

        private RootDettagliEnte RicercaEnte(string nomeEnte)
        {
            // Chiamata per ricercare l'ente
            string api = "https://www.indicepa.gov.it:443/public-ws/WS16_DES_AMM.php";

            NameValueCollection outgoingQueryString = HttpUtility.ParseQueryString(String.Empty);
            outgoingQueryString.Add("AUTH_ID", ConfigurationManager.AppSettings["indicePa.authID"]);
            outgoingQueryString.Add("DESCR", nomeEnte);

            var risultato = chiamataMultiPart(api, outgoingQueryString);

            var ente = risultato == null ? null : new JavaScriptSerializer().Deserialize<RootEnte>(risultato);

            // Se trovo un risultato lo associo ai dati
            if (ente != null && 
                ente.result.num_items == 1 && 
                ente.data.Any(x => x.des_amm.ToLower().Contains("comune")))
            {
                // Esiste soltanto un ente ed è il comune
                var dettagliEnte = RicercaDettagliEnte(ente.data.SingleOrDefault().cod_amm);

                return dettagliEnte;
            }
            else if(ente != null && ente.result.num_items > 1)
            {
                // Esistono più enti con lo stesso nome, prendo soltanto il comune
                var comune = ente.data.FirstOrDefault(x => x.des_amm.ToLower().Contains("comune")).cod_amm;
                var dettagliEnte = RicercaDettagliEnte(comune);

                return dettagliEnte;
            }
            else
            {
                // L'ente non è stato trovato
                return null;
            }
        }

        private RootDettagliEnte RicercaDettagliEnte(string cod_amm)
        {
            string api = "https://www.indicepa.gov.it:443/public-ws/WS05_AMM.php";

            NameValueCollection outgoingQueryString = HttpUtility.ParseQueryString(String.Empty);
            outgoingQueryString.Add("AUTH_ID", ConfigurationManager.AppSettings["indicePa.authID"]);
            outgoingQueryString.Add("COD_AMM", cod_amm);

            var risultato = chiamataMultiPart(api, outgoingQueryString);

            var dettagliEnte = new JavaScriptSerializer().Deserialize<RootDettagliEnte>(risultato);

            return dettagliEnte;
        }

        private string chiamataMultiPart(string api, NameValueCollection outgoingQueryString)
        {
            string boundary = "---------------------------" + DateTime.Now.Ticks.ToString("x");
            byte[] boundarybytes = System.Text.Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");

            var request = (HttpWebRequest)WebRequest.Create(api);
            request.ContentType = "multipart/form-data; boundary=" + boundary;
            request.Headers.Add("accept-language", "it-IT");
            request.Method = "POST";
            request.KeepAlive = true;
            request.Credentials = CredentialCache.DefaultCredentials;

            Stream rs = request.GetRequestStream();
            string formdataTemplate = "Content-Disposition: form-data; name=\"{0}\"\r\n\r\n{1}";

            foreach (string key in outgoingQueryString.Keys)
            {
                rs.Write(boundarybytes, 0, boundarybytes.Length);
                string formitem = string.Format(formdataTemplate, key, outgoingQueryString[key]);
                byte[] formitembytes = Encoding.UTF8.GetBytes(formitem);
                rs.Write(formitembytes, 0, formitembytes.Length);
            }

            rs.Write(boundarybytes, 0, boundarybytes.Length);

            WebResponse wresp = null;
            try
            {
                wresp = request.GetResponse();
                Stream stream2 = wresp.GetResponseStream();
                StreamReader reader3 = new StreamReader(stream2);

                string streamJson3 = reader3.ReadToEnd();

                return streamJson3;
            }
            catch (Exception ex)
            {
                if (wresp != null)
                {
                    wresp.Close();
                    wresp = null;
                }

                return null;
            }
        }

        public class ComuneValore
        {
            public dynamic CodRegione { get; set; }
            public dynamic CodUnitaTerritoriale { get; set; }
            public dynamic CodProvinciaStorico { get; set; }
            public dynamic CodProgressivoComune { get; set; }
            public dynamic CodComuneAlfanumerico { get; set; }
            public dynamic DenominazioneUniversale { get; set; }
            public dynamic DenominazioneItaliana { get; set; }
            public dynamic DenominazioneAltraLingua { get; set; }
            public dynamic CodiceRipartizioneGeografica { get; set; }
            public dynamic RipartizioneGeografica { get; set; }
            public dynamic DenominazioneRegione { get; set; }
            public dynamic DenominazioneUnitaTerritorialeSovracomunale { get; set; }
            public dynamic TipologiaUnitaTerritorialeSovracomunale { get; set; }
            public dynamic Capoluogo_metropolitana_liberoConsorzio { get; set; }
            public dynamic SiglaAutomobilistica { get; set; }
            public dynamic CodComuneNumerico { get; set; }
            public dynamic CodComuneNumerico_110Provincie_2010_2016 { get; set; }
            public dynamic CodComuneNumerico_107Provincie_2006_2009 { get; set; }
            public dynamic CodComuneNumerico_103Provincie_1995_2005 { get; set; }
            public dynamic CodCatastale { get; set; }
            public dynamic CodNuts1_2010 { get; set; }
            public dynamic CodNuts2_2010 { get; set; }
            public dynamic CodNuts3_2010 { get; set; }
            public dynamic CodNuts1_2021 { get; set; }
            public dynamic CodNuts2_2021 { get; set; }
            public dynamic CodNuts3_2021 { get; set; }
            public RootDettagliEnte DettagliEnte { get; set; }
        }


        // Nuova chiamata GET ENTE
        public class Datum
        {
            public string cod_amm { get; set; }
            public string acronimo { get; set; }
            public string des_amm { get; set; }
        }

        public class Result
        {
            public int cod_err { get; set; }
            public string desc_err { get; set; }
            public int num_items { get; set; }
        }

        public class RootEnte
        {
            public Result result { get; set; }
            public List<Datum> data { get; set; }
        }

        // Nuova POST Dettagli Ente

        public class DataDettagliEnte
        {
            public string cod_amm { get; set; }
            public string acronimo { get; set; }
            public string des_amm { get; set; }
            public string regione { get; set; }
            public string provincia { get; set; }
            public string comune { get; set; }
            public string cap { get; set; }
            public string indirizzo { get; set; }
            public string titolo_resp { get; set; }
            public string nome_resp { get; set; }
            public string cogn_resp { get; set; }
            public string sito_istituzionale { get; set; }
            public object liv_access { get; set; }
            public string mail1 { get; set; }
            public string mail2 { get; set; }
            public string mail3 { get; set; }
            public string mail4 { get; set; }
            public string mail5 { get; set; }
            public string tipologia { get; set; }
            public string categoria { get; set; }
            public string data_accreditamento { get; set; }
            public string cf { get; set; }
        }

        public class ResultDettagliEnte
        {
            public int cod_err { get; set; }
            public string desc_err { get; set; }
            public int num_items { get; set; }
        }

        public class RootDettagliEnte
        {
            public ResultDettagliEnte result { get; set; }
            public DataDettagliEnte data { get; set; }
        }

    }
}