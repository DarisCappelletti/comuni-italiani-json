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
using PortFolio.comuni_italiani_json.Helpers;

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
                        var obj = creoOggettoComune(excelReader);
                        
                        // Scarico i dettagli da indicePA
                        var dettagliEnte = 
                            richiestaDettagliIndicePa == "true" 
                            ? IstatHelper.RicercaDettagliEnte(obj.DenominazioneItaliana) 
                            : null;

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

        public static string chiamataMultiPart(string api, NameValueCollection outgoingQueryString)
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

        public static ComuneValore creoOggettoComune(IExcelDataReader excelReader)
        {
            var obj = new ComuneValore
            {
                CodRegione = excelReader.GetString(0),
                CodUnitaTerritoriale = excelReader.GetString(1),
                CodProvinciaStorico = excelReader.GetString(2),
                CodProgressivoComune = excelReader.GetString(3),
                CodComuneAlfanumerico = excelReader.GetString(4),
                DenominazioneUniversale = excelReader.GetString(5),
                DenominazioneItaliana = excelReader.GetString(6),
                DenominazioneAltraLingua = excelReader.GetString(7),
                CodiceRipartizioneGeografica = excelReader.GetDouble(8),
                RipartizioneGeografica = excelReader.GetString(9),
                DenominazioneRegione = excelReader.GetString(10),
                DenominazioneUnitaTerritorialeSovracomunale = excelReader.GetString(11),
                TipologiaUnitaTerritorialeSovracomunale = excelReader.GetDouble(12),
                Capoluogo_metropolitana_liberoConsorzio = excelReader.GetDouble(13),
                SiglaAutomobilistica = excelReader.GetString(14),
                CodComuneNumerico = excelReader.GetDouble(15),
                CodComuneNumerico_110Provincie_2010_2016 = excelReader.GetDouble(16),
                CodComuneNumerico_107Provincie_2006_2009 = excelReader.GetDouble(17),
                CodComuneNumerico_103Provincie_1995_2005 = excelReader.GetDouble(18),
                CodCatastale = excelReader.GetString(19),
                CodNuts1_2010 = excelReader.GetString(20),
                CodNuts2_2010 = excelReader.GetString(21),
                CodNuts3_2010 = excelReader.GetString(22),
                CodNuts1_2021 = excelReader.GetString(23),
                CodNuts2_2021 = excelReader.GetString(24),
                CodNuts3_2021 = excelReader.GetString(25)
            };

            return obj;
        }

        public class ComuneValore
        {
            public string CodRegione { get; set; }
            public string CodUnitaTerritoriale { get; set; }
            public string CodProvinciaStorico { get; set; }
            public string CodProgressivoComune { get; set; }
            public string CodComuneAlfanumerico { get; set; }
            public string DenominazioneUniversale { get; set; }
            public string DenominazioneItaliana { get; set; }
            public string DenominazioneAltraLingua { get; set; }
            public double CodiceRipartizioneGeografica { get; set; }
            public string RipartizioneGeografica { get; set; }
            public string DenominazioneRegione { get; set; }
            public string DenominazioneUnitaTerritorialeSovracomunale { get; set; }
            public double TipologiaUnitaTerritorialeSovracomunale { get; set; }
            public double Capoluogo_metropolitana_liberoConsorzio { get; set; }
            public string SiglaAutomobilistica { get; set; }
            public double CodComuneNumerico { get; set; }
            public double CodComuneNumerico_110Provincie_2010_2016 { get; set; }
            public double CodComuneNumerico_107Provincie_2006_2009 { get; set; }
            public double CodComuneNumerico_103Provincie_1995_2005 { get; set; }
            public string CodCatastale { get; set; }
            public string CodNuts1_2010 { get; set; }
            public string CodNuts2_2010 { get; set; }
            public string CodNuts3_2010 { get; set; }
            public string CodNuts1_2021 { get; set; }
            public string CodNuts2_2021 { get; set; }
            public string CodNuts3_2021 { get; set; }
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