using ExcelDataReader;
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

namespace PortFolio.comuni_italiani_json
{
    public class ComuniItalianiController : ApiController
    {
        // GET api/<controller>
        public List<ComuneValore> Get(string comune)
        {
            return RicercaComuniIstat(comune);
        }

        private List<ComuneValore> RicercaComuniIstat(string comune)
        {
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
                    if (i != 0 && excelReader.GetString(6).ToLower().Contains(comune))
                    {
                        // Imposto l'oggetto
                        var obj = download_json.creoOggettoComune(excelReader);

                        // Scarico i dettagli da indicePA
                        var dettagliEnte = RicercaEnte(obj.DenominazioneItaliana);

                        obj.DettagliEnte = dettagliEnte;

                        list.Add(obj);
                    }

                    i++;
                }

                excelReader.Close();

                return list;
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
            else if (ente != null && ente.result.num_items > 1)
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

            var risultato = download_json.chiamataMultiPart(api, outgoingQueryString);

            var dettagliEnte = new JavaScriptSerializer().Deserialize<RootDettagliEnte>(risultato);

            return dettagliEnte;
        }
    }
}