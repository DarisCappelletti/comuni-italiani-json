using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PortFolio.comuni_italiani_json.Helpers;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web.Http;
using System.Web.Script.Serialization;
using static PortFolio.comuni_italiani_json.Models.WikiData;

namespace PortFolio.comuni_italiani_json.api
{
    public class ComuniItalianiCoordinateController : ApiController
    {
        // GET api/<controller>
        public List<DettagliWikiData> Get(string comune = "")
        {
            return RicercaComuniIstat(comune);
        }

        private List<DettagliWikiData> RicercaComuniIstat(string comune)
        {
            // Inizializzo la lista
            var list = new List<DettagliWikiData>();

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
                    if (i != 0 && (comune == "" || excelReader.GetString(6).ToLower().Contains(comune)))
                    {
                        // Scarico le coordinate da indicePA
                        var coordinate = RicercaEnteWikiData(excelReader.GetString(6));

                        if(coordinate != null)
                        {
                            coordinate.NomeComune = excelReader.GetString(6);
                            coordinate.CodAmm = IstatHelper.RicercaEnte(excelReader.GetString(6));

                            list.Add(coordinate);
                        }
                    }

                    i++;
                }

                excelReader.Close();

                return list;
            }
        }

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

            // Ricerco l'ente che corrisponde al nome passato oppure lo contiene
            var ente = data.search.FirstOrDefault(x => 
                x.description != null && 
                x.label == nomeEnte &&
                x.description.ToLower() == "comune italiano");

            if(ente == null)
            {
                return null;
            }
            else
            {
                return RicercaCoordinateEnteWikiData(ente.id);
            }
        }

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
            var coord = comuneWiki.claims.P625[0].mainsnak.datavalue.value;
            var latitudine = coord.latitude.ToString();
            var longitudine = coord.longitude.ToString();

            // Prendo il CAP - Codice P281 su WikiData
            var cap = comuneWiki.claims.P281[0].mainsnak.datavalue.value;

            return new DettagliWikiData()
            {
                CAP = cap,
                latitude = latitudine,
                longitude = longitudine
            };
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
    }
}