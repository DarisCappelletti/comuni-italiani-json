using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Script.Serialization;
using static PortFolio.comuni_italiani_json.download_json;

namespace PortFolio.comuni_italiani_json.Helpers
{
    public static class IstatHelper
    {
        public static string RicercaEnte(string nomeEnte)
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
                return ente.data.SingleOrDefault().cod_amm;
            }
            else if (ente != null && ente.result.num_items > 1)
            {
                // Esistono più enti con lo stesso nome, prendo soltanto il comune
                var comune = ente.data.FirstOrDefault(x => x.des_amm.ToLower().Contains("comune")).cod_amm;

                return comune;
            }
            else
            {
                // L'ente non è stato trovato
                return null;
            }
        }

        public static RootDettagliEnte RicercaDettagliEnte(string nomeEnte)
        {
            var cod_amm = RicercaEnte(nomeEnte);

            if (cod_amm != null)
            {
                string api = "https://www.indicepa.gov.it:443/public-ws/WS05_AMM.php";

                NameValueCollection outgoingQueryString = HttpUtility.ParseQueryString(String.Empty);
                outgoingQueryString.Add("AUTH_ID", ConfigurationManager.AppSettings["indicePa.authID"]);
                outgoingQueryString.Add("COD_AMM", cod_amm);

                var risultato = chiamataMultiPart(api, outgoingQueryString);

                var dettagliEnte = new JavaScriptSerializer().Deserialize<RootDettagliEnte>(risultato);

                return dettagliEnte;
            }
            else
            {
                return null;
            }
        }
    }
}