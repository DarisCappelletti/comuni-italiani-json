using ExcelDataReader;
using System;
using System.Collections.Generic;
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
    public partial class Index : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            RicercaComuni();
        }

        private void RicercaComuni()
        {
            var list = new List<ComuneValore>();

            Stream streamXls = null;
            
            using (var wc = new WebClient())
            {
                streamXls = wc.OpenRead("https://www.istat.it/storage/codici-unita-amministrative/Elenco-comuni-italiani.xls");
            }

            using (var mss = new MemoryStream())
            {
                streamXls.CopyTo(mss);
                mss.Position = 0;

                //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(mss);
                //...
                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                //IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                var i = 0;

                //5. Data Reader methods
                while (excelReader.Read())
                {
                    if (i != 0)
                    {
                        var obj = new ComuneValore
                        {
                            CodRegione = excelReader.GetValue(0),
                            CodUnitaTerritoriale = excelReader.GetValue(1),
                            CodProvinciaStorico = excelReader.GetValue(2),
                            CodProgressivoComune = excelReader.GetValue(3),
                            CodComuneAlfanumerico = excelReader.GetValue(4),
                            DenominazioneUniversale = excelReader.GetValue(5),
                            DenominazioneItaliana = excelReader.GetValue(6),
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
                        list.Add(obj);
                    }

                    i++;
                }

                //6. Free resources (IExcelDataReader is IDisposable)
                excelReader.Close();
                var json = JsonSerializer.Serialize(list);
                BinaryFormatter bf = new BinaryFormatter();
                MemoryStream ms = new MemoryStream();
                bf.Serialize(ms, json);

                var byteArray = ms.ToArray();
                // processing the stream.

                string FileName = "ComuniItaliani_" + DateTime.Now + ".json";
                HttpResponse response = HttpContext.Current.Response;
                response.Clear();
                response.Charset = "";
                response.ContentType = "application/json";
                response.ContentEncoding = Encoding.UTF8;
                response.AddHeader("Content-Disposition", "attachment;filename=" + FileName);
                //response.BinaryWrite(byteArray);
                response.Write(json);
                response.End();
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
        }

        public class CODICIAl03032022
        {
            public string CodiceRegione { get; set; }
            public string CodiceDellUnitàTerritorialeSovracomunaleValidaAFiniStatistici { get; set; }
            public string CodiceProvinciaStorico1 { get; set; }
            public string ProgressivoDelComune2 { get; set; }
            public string CodiceComuneFormatoAlfanumerico { get; set; }
            public string DenominazioneItalianaEStraniera { get; set; }
            public string DenominazioneInItaliano { get; set; }
            public int CodiceRipartizioneGeografica { get; set; }
            public string RipartizioneGeografica { get; set; }
            public string DenominazioneRegione { get; set; }
            public string DenominazioneDellUnitàTerritorialeSovracomunaleValidaAFiniStatistici { get; set; }
            public int TipologiaDiUnitàTerritorialeSovracomunale { get; set; }
            public int FlagComuneCapoluogoDiProvinciaCittàMetropolitanaLiberoConsorzio { get; set; }
            public string SiglaAutomobilistica { get; set; }
            public int CodiceComuneFormatoNumerico { get; set; }
            public int CodiceComuneNumericoCon110ProvinceDal2010Al2016 { get; set; }
            public int CodiceComuneNumericoCon107ProvinceDal2006Al2009 { get; set; }
            public int CodiceComuneNumericoCon103ProvinceDal1995Al2005 { get; set; }
            public string CodiceCatastaleDelComune { get; set; }
            public string CodiceNUTS12010 { get; set; }
            public string CodiceNUTS220103 { get; set; }
            public string CodiceNUTS32010 { get; set; }
            public string CodiceNUTS12021 { get; set; }
            public string CodiceNUTS220213 { get; set; }
            public string CodiceNUTS32021 { get; set; }
        }
    }
}