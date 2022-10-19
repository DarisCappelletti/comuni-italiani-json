using System;
using System.Configuration;
using System.Web.UI;

namespace PortFolio.comuni_italiani_json_api
{
    public partial class MappaComuni : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            litSiteUrl.Text = $"'{ConfigurationManager.AppSettings["SiteUrl"]}'";
        }
    }
}