using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Web;
using System.Web.Http;
using System.Web.Script.Serialization;
using static PortFolio.comuni_italiani_json.download_json;
using static PortFolio.comuni_italiani_json.Models.WikiData;

namespace PortFolio.comuni_italiani_json.api
{
    public class ComuniItalianiController : ApiController
    {
        // GET api/<controller>
        public List<Comune> Get(string comune = null, bool datiIndicePA = false, bool datiWikiData = false)
        {
            // Ritorno la lista dei comuni italiani
            return RicercaComuniIstat(comune, datiIndicePA, datiWikiData);
        }

        private List<Comune> RicercaComuniIstat(
            string comune = null,
            bool datiIndicePA = false,
            bool datiWikiData = false)
        {
            // Inizializzo la lista
            var list = new List<Comune>();

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
                    // Verifico che il comune sia uguale o se contiene la parola passata
                    if (
                            i != 0 &&
                            (
                                comune == null ||
                                comune == "" ||
                                excelReader.GetString(6).ToLower() == comune.ToLower() ||
                                excelReader.GetString(6).ToLower().Contains(comune)
                            )
                        )
                    {
                        var com = new Comune();

                        // Imposto l'oggetto
                        com.DettagliIstat = creoDettagliIstat(excelReader);

                        com.Nome = com.DettagliIstat.DenominazioneItaliana;
                        com.Regione = com.DettagliIstat.DenominazioneRegione;
                        com.Provincia = com.DettagliIstat.SiglaAutomobilistica;

                        if (datiIndicePA)
                        {
                            // Scarico i dettagli da indicePA
                            var dettagliEnte = RicercaEnteIndicePA(com.DettagliIstat.DenominazioneItaliana);

                            com.DettagliIndicePA = dettagliEnte;
                        }
                        if (datiWikiData)
                        {
                            // Scarico i dettagli da WikiData
                            var dettagliWikiData = RicercaEnteWikiData(com.DettagliIstat.DenominazioneItaliana);

                            com.DettagliWikiData = dettagliWikiData;
                        }

                        // Aggiunto il comune alla lista
                        list.Add(com);
                    }

                    i++;
                }

                excelReader.Close();

                return list;
            }
        }

        // Ricerca del comune su IndicePA
        private DettagliIndicePA RicercaEnteIndicePA(string nomeEnte)
        {
            // If ad hoc per il comune di "Front Canavese"
            if (nomeEnte == "Front") { nomeEnte = "Front Canavese"; }

            // Sostituisco i caratteri speciali per la chiamata GET
            var nomeEntePerChiamataGET = ConvertNomeComuneChiamataIndicePA(nomeEnte);

            // Sostituisco i caratteri speciali per la ricerca del comune
            nomeEnte = ConvertNomeComuneRicercaIndicePA(nomeEnte);

            // Chiamata per ricercare l'ente
            string api = "https://www.indicepa.gov.it:443/public-ws/WS16_DES_AMM.php";

            NameValueCollection outgoingQueryString = HttpUtility.ParseQueryString(String.Empty);
            outgoingQueryString.Add("AUTH_ID", ConfigurationManager.AppSettings["indicePa.authID"]);
            outgoingQueryString.Add("DESCR", nomeEntePerChiamataGET);

            var risultato = chiamataMultiPart(api, outgoingQueryString);

            var ente =
                risultato == null
                ? null
                : new JavaScriptSerializer().Deserialize<RootEnte>(risultato);

            // Se trovo un risultato lo associo ai dati
            if (ente != null &&
                ente.result.num_items == 1 &&
                ente.data.Any(x => x.des_amm.ToLower().Contains("comune")))
            {
                // Esiste soltanto un ente ed è il comune
                var dettagliEnte = RicercaDettagliEnteIndicePA(ente.data.SingleOrDefault().cod_amm);

                return dettagliEnte;
            }
            else if (ente != null && ente.result.num_items > 1)
            {
                // Esistono più enti con lo stesso nome, prendo soltanto il comune
                string nomeComune = "comune di " + nomeEnte.ToLower();
                var comune = ente.data.FirstOrDefault(x =>
                        x.des_amm.ToLower() == nomeComune);

                if (comune != null)
                {
                    var dettagliEnte = RicercaDettagliEnteIndicePA(comune.cod_amm);
                    return dettagliEnte;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                // L'ente non è stato trovato
                return null;
            }
        }

        // Ricerca dei dettagli del comune su IndicePA
        private DettagliIndicePA RicercaDettagliEnteIndicePA(string cod_amm)
        {
            string api = "https://www.indicepa.gov.it:443/public-ws/WS05_AMM.php";

            NameValueCollection outgoingQueryString = HttpUtility.ParseQueryString(String.Empty);
            outgoingQueryString.Add("AUTH_ID", ConfigurationManager.AppSettings["indicePa.authID"]);
            outgoingQueryString.Add("COD_AMM", cod_amm);

            var risultato = download_json.chiamataMultiPart(api, outgoingQueryString);

            var dettagliEnte = new JavaScriptSerializer().Deserialize<RootIndicePA>(risultato);

            return dettagliEnte.data;
        }

        // Ricerca del comune su WikiData
        public DettagliWikiData RicercaEnteWikiData(string nomeEnte)
        {
            // Chiamata per ricercare l'ente
            string api = $"https://www.wikidata.org/w/api.php?action=wbsearchentities&search={nomeEnte}&format=json&errorformat=plaintext&language=it&uselang=it&type=item";

            var request = (HttpWebRequest)WebRequest.Create(api);
            request.ContentType = "application/json";
            request.Method = "GET";

            WebResponse response = request.GetResponse() as HttpWebResponse;
            var stream = response.GetResponseStream();
            StreamReader reader = new StreamReader(stream);
            // Leggo la risposta.
            string streamJson = reader.ReadToEnd();

            var data = new JavaScriptSerializer().Deserialize<WikiDataRoot>(streamJson);

            // Ricerco l'ente che corrisponde al nome passato
            var ente = data.search == null ? null : data.search.FirstOrDefault(x =>
                x.description != null &&
                x.label.ToLower() == nomeEnte.ToLower() &&
                x.description.ToLower().Contains("comune italiano"));

            if (ente == null)
            {
                return null;
            }
            else
            {
                return RicercaCoordinateEnteWikiData(ente.id);
            }
        }

        // Ricerca dei dettagli del comune su WikiData
        private DettagliWikiData RicercaCoordinateEnteWikiData(string id)
        {
            string api = $"https://www.wikidata.org/w/api.php?action=wbgetentities&ids={id}&format=json&languages=it&props=claims";

            var request = (HttpWebRequest)WebRequest.Create(api);
            request.ContentType = "application/json";
            request.Method = "GET";

            WebResponse response = request.GetResponse() as HttpWebResponse;
            var stream = response.GetResponseStream();
            StreamReader reader = new StreamReader(stream);
            // Leggo la risposta.
            string streamJson = reader.ReadToEnd();

            var data = new JavaScriptSerializer().Deserialize<Root>(streamJson);
            var dynamicObj = JsonConvert.DeserializeObject<Root>(streamJson);

            // A causa della deserializzazione con javascriptserializer è necessario
            // farlo su un oggetto anonimo (non permette di impostare un nome fisso ad una proprietà della classe
            var stuff = new JavaScriptSerializer().Deserialize<object>(streamJson);
            dynamic hitchhiker = JObject.Parse(streamJson);
            var entities = hitchhiker.entities;

            // Estrapolo l'elemento con l'id
            var comuneWiki = entities[id];

            // Prendo le coordinate - Codice P625 su WikiData
            var coord = comuneWiki.claims.P625 != null ? comuneWiki.claims.P625[0].mainsnak.datavalue.value : null;
            var latitudine = coord.latitude.ToString();
            var longitudine = coord.longitude.ToString();

            // Prendo il CAP - Codice P281 su WikiData
            var cap = comuneWiki.claims.P281 != null ? comuneWiki.claims.P281[0].mainsnak.datavalue.value : null;

            return new DettagliWikiData()
            {
                CAP = cap,
                latitude = latitudine,
                longitude = longitudine
            };
        }

        // Converte i caratteri speciali per le chiamate alle api di IncidePA
        private string ConvertNomeComuneChiamataIndicePA(string comune)
        {
            // Sostituisco il singolo apice con 2 apici
            if (comune.Contains("'")) { comune = comune.Replace("'", "''"); }
            // Sostituisco il tratto con uno spazio vuoto
            if (comune.Contains("-") && !ComuniConTratto.Any(x => x == comune)) { comune = comune.Replace("-", " "); }

            // Effettuo il replace dei caratteri speciali
            if (comune.Contains("à")) {
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach (var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("à") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "a''" : comuneSplit;
                    str = str.Contains("à") ? str.Replace("à", "a") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            if (comune.Contains("è")) 
            {
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach(var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("è") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "e''" : comuneSplit;
                    str = str.Contains("è") ? str.Replace("è", "e") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            if (comune.Contains("é")) { 
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach (var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("é") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "e''" : comuneSplit;
                    str = str.Contains("é") ? str.Replace("é", "e") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            if (comune.Contains("ê")) { 
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach (var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("ê") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "e''" : comuneSplit;
                    str = str.Contains("ê") ? str.Replace("ê", "e") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            if (comune.Contains("ì")) { 
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach (var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("ì") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "i''" : comuneSplit;
                    str = str.Contains("ì") ? str.Replace("ì", "i") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            if (comune.Contains("ò")) { 
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach (var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("ò") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "o''" : comuneSplit;
                    str = str.Contains("ò") ? str.Replace("ò", "o") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            if (comune.Contains("ù")) {
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach (var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("ù") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "u''" : comuneSplit;
                    str = str.Contains("ù") ? str.Replace("ù", "u") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            return comune;
        }

        private string ConvertNomeComuneRicercaIndicePA(string comune)
        {
            // Effettuo il replace dei caratteri speciali
            if (comune.Contains("à")) {
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach (var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("à") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "a'" : comuneSplit;
                    str = str.Contains("à") ? str.Replace("à", "a") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            if (comune.Contains("è")) 
            {
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach(var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("è") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "e'" : comuneSplit;
                    str = str.Contains("è") ? str.Replace("è", "e") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            if (comune.Contains("é")) { 
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach (var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("é") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "e'" : comuneSplit;
                    str = str.Contains("é") ? str.Replace("é", "e") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            if (comune.Contains("ê")) { 
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach (var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("ê") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "e'" : comuneSplit;
                    str = str.Contains("ê") ? str.Replace("ê", "e") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            if (comune.Contains("ì")) { 
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach (var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("ì") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "i'" : comuneSplit;
                    str = str.Contains("ì") ? str.Replace("ì", "i") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            if (comune.Contains("ò")) { 
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach (var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("ò") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "o'" : comuneSplit;
                    str = str.Contains("ò") ? str.Replace("ò", "o") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            if (comune.Contains("ù")) {
                var split = comune.Split(' ');
                var stringaFinale = "";
                foreach (var comuneSplit in split)
                {
                    var str = comuneSplit.EndsWith("ù") ? comuneSplit.Remove(comuneSplit.Length - 1, 1) + "u'" : comuneSplit;
                    str = str.Contains("ù") ? str.Replace("ù", "u") : str;
                    stringaFinale += str + " ";
                }
                comune = stringaFinale.Trim();
            }
            return comune;
        }

        public static DettagliIstat creoDettagliIstat(IExcelDataReader excelReader)
        {
            var obj = new DettagliIstat
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

        public class Comune
        {
            public string Nome { get; set; }
            public string Regione { get; set; }
            public string Provincia { get; set; }
            public DettagliIstat DettagliIstat { get; set; }
            public DettagliIndicePA DettagliIndicePA { get; set; }
            public DettagliWikiData DettagliWikiData { get; set; }
        }

        public class WikiDataDescription
        {
            public string value { get; set; }
            public string language { get; set; }
        }

        public class WikiDataDisplay
        {
            public WikiDataLabel label { get; set; }
            public WikiDataDescription description { get; set; }
        }

        public class WikiDataLabel
        {
            public string value { get; set; }
            public string language { get; set; }
        }

        public class Match
        {
            public string type { get; set; }
            public string language { get; set; }
            public string text { get; set; }
        }

        public class WikiDataRoot
        {
            public WikiDataSearchinfo searchinfo { get; set; }
            public List<WikiDataSearch> search { get; set; }
            public int SearchContinue { get; set; }
            public int success { get; set; }
        }

        public class WikiDataSearch
        {
            public string id { get; set; }
            public string title { get; set; }
            public int pageid { get; set; }
            public WikiDataDisplay display { get; set; }
            public string repository { get; set; }
            public string url { get; set; }
            public string concepturi { get; set; }
            public string label { get; set; }
            public string description { get; set; }
            public Match match { get; set; }
        }

        public class WikiDataSearchinfo
        {
            public string search { get; set; }
        }

        public static readonly List<string> ComuniConTratto = new List<string>(
          new string[] {
            "Gattico-Veruno",
            "Antey-Saint-André",
            "Challand-Saint-Anselme",
            "Gressoney-La-Trinité",
            "Gressoney-Saint-Jean",
            "Pont-Saint-Martin",
            "Pré-Saint-Didier",
            "Rhêmes-Notre-Dame",
            "Rhêmes-Saint-Georges",
            "Saint-Christophe",
            "Saint-Denis",
            "Saint-Marcel",
            "Saint-Nicolas",
            "Saint-Oyen",
            "Saint-Pierre",
            "Saint-Rhémy-en-Bosses",
            "Saint-Vincent",
            "Laveno-Mombello",
            "Travedona-Monate",
            "Uggiate-Trevano",
            "Casale Cremasco-Vidolasco",
            "Castelbello-Ciardes",
            "Monguelfo-Tesido",
            "Senale-San Felice",
            "Castello-Molina di Fiemme",
            "Lona-Lases",
            "Nago-Torbole",
            "Ruffrè-Mendola",
            "Pieve di Bono-Prezzo",
            "Cavallino-Treporti",
            "Chiopris-Viscone",
            "Montescudo-Monte Colombo",
            "Presicce-Acquarica",
            "Corigliano-Rossano"
          }
        );
    }
}