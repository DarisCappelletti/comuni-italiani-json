﻿<%@ Page 
    Title="Mappa comuni"
    Language="C#" 
    MasterPageFile="~/Site.Master"
    AutoEventWireup="true" 
    CodeBehind="MappaComuni.aspx.cs" 
    Inherits="PortFolio.comuni_italiani_json_api.MappaComuni" 
%>

<asp:Content ID="HeadContet" ContentPlaceHolderID="Head" runat="server">
    <style>
        #map { height: 480px; }
    </style>
    <link href="Assets/Vendors/leaflet-1.8.0/leaflet.css" rel="stylesheet" />
    <link href="Assets/Vendors/Leaflet.markercluster-1.5.3/dist/MarkerCluster.css" rel="stylesheet" />
    <link href="Assets/Vendors/Leaflet.markercluster-1.5.3/dist/MarkerCluster.Default.css" rel="stylesheet" />
    <script src="Assets/Vendors/leaflet-1.8.0/leaflet.js"></script>
    <script src="Assets/Vendors/Leaflet.markercluster-1.5.3/dist/leaflet.markercluster.js"></script>
</asp:Content>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server" class="container-fluid">
    <h4>Mappa dei comuni italiani</h4>
    <p>Ho implementato questa mappa utilizzanto la libreria leaflet.js in modo da poter consultare tutti i comuni italiani.<br />
       I punti della mappa vengono generati tramite un file json precedentemente scaricato tramite un mio altro progetto ovvero 
        <a href="https://github.com/DarisCappelletti/comuni-italiani-json-api">"Comuni italiani json"</a>.<br />
        Cliccando su un punto verranno mostrare le informazioni collegate a quel determinato comune.
        Gli stemmi sono stati scaricati utilizzando il mio progetto <a href="https://github.com/DarisCappelletti/Stemmi-comuni-italiani">"Stemmi comuni italiani"</a>.
    </p>

    <div id="map"></div>

<script>
    var siteUrl = <asp:Literal ID="litSiteUrl" runat="server" />
    var map
    $(document).ready(function () {
        // Inizializzo la mappa
        map = L.map('map').setView([43.5, 13.5], 8);
        L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', {
            maxZoom: 19,
            attribution: '© OpenStreetMap'
        }).addTo(map);

        // Group we will append our markers to
        var element = document.getElementById('map')
        if (element !== null) {
            // Set up cluster group
            markers = new L.MarkerClusterGroup()
        } else {
            // Otherwise set up normal groupx`
            markers = new L.LayerGroup()
        }

        markers.on("click", generatePopup);

        // Name of lat, long columns in Google spreadsheet
        var latitudine = 'Latitudine'
        var longitudine = 'Longitudine'

        // Marker options
        var global_markers_data

        // chiamata alla funzione di inizializzazione mappa con i dati dal json...
        initMap()
    });



    // Imposta i dati del json sulla mappa
    function loadMarkersToMap(markers_data) {
        // If we haven't captured the Tabletop data yet
        // We'll set the Tabletop data to global_markers_data
        if (global_markers_data !== undefined) {
            markers_data = global_markers_data;
            // If we have, we'll load global_markers_data instead of loading Tabletop again
        } else {
            global_markers_data = markers_data;
        }

        for (var num = 0; num < markers_data.length; num++) {
            // Capture current iteration through JSON file
            current = markers_data[num];
            let lat = current.latitude.replace(',', '.')
            let long = current.longitude.replace(',', '.')
            console.log(lat)
            console.log(long)
            //controllo se la proprieta'  latitudine non e' vuota...
            if ((lat !== null) && (long !== null)) {
                // Add lat, long to marker
                var marker_location = new L.LatLng(lat, long);

                // Options for our circle marker
                var layer_marker = L.marker(marker_location, {
                    opacity: 1
                });

                // Generate popup
                layer_marker.bindPopup(generatePopup(current), { minWidth: 250, maxWidth: 350 });
                // Add the ID

                feature = layer_marker.feature = layer_marker.feature || {}; // Initialize feature

                feature.type = feature.type || "Feature"; // Initialize feature.type

                // Preparo i parametri per i dettagli
                var props = feature.properties = feature.properties || {}; // Initialize feature.properties
                //props.Id = current.Id;
                //props.IdTpe = current.IdTpe;
                //props.Nome = current.Nome;
                //props.TipoEnte = current.TipoEnte;
                //props.Ente = current.Ente;
                //props.Uo = current.Uo;
                //props.Stemma = current.Stemma;
                //props.Categoria = current.Categoria;
                //props.Descrizione = current.Descrizione;

                // Add to feature group
                markers.addLayer(layer_marker);

            }
        }

        // Add feature group to map
        map.addLayer(markers);

    }

    async function initMap(filtroTesto, filtroTipoEnte) {
        // resetto i markers
        markers.clearLayers();
        global_markers_data = undefined;

        // imposto i parametri per la chiamata
        filtroTesto = filtroTesto == undefined ? "" : filtroTesto
        await fetch(siteUrl + 'api/ComuniItalianiCoordinate?dataFromJson=true'
            , {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json'
                }
            })
            .then(async (response) => {
                let data = await response.json();
                loadMarkersToMap(data);
                if (data.length == 1) {
                    var latitudine = data[0].Latitudine == undefined ? "" : data[0].Latitudine.toString()
                    var longitudine = data[0].Longitudine == undefined ? "" : data[0].Longitudine.toString()
                    simulaClickEvento(latitudine, longitudine);
                    generateDettagli(data[0])
                }
                else {
                    /*$('.dettagli-ente').empty()*/
                }
            })
    }

    function generatePopup(content) {
        // Generate header
        var nome_comune = content['NomeComune'];

        /*var popup_header = '<h4><img class="stemma" src="https://procedimenti.regione.marche.it/images/stemmi/';
        popup_header += cf_comune + '.png" alt="stemma ' + nome_comune + '">';
    popup_header += nome_comune + '</h4><hr>'; */

        // Imposto lo stemma del comune
        var popup_header = '<div class="row"><div class="col-md-3">';
        popup_header += '<img class="stemma" src="' + siteUrl +'/comuni-italiani-json-api/Assets/Images/Stemmi/' + nome_comune;
        popup_header += '-Stemma.png" alt="stemma ' + nome_comune + '" style="width: 64px;"></div>';
        popup_header += '<div class="col-md-9"><h4>' + nome_comune + '</h4></div></div><hr>';


        // Generate content
        var popup_content = '<div class="row">';
        //if (totale_procedimenti === 0) //il comune ha aderito ma non ha inserito procedimenti, quindi visualizzo solo il pulsante con link generico
        //{
        //    popup_content += '<div class="col-md-5"><a href="/TipologieProcedimento/Index2"';
        //    popup_content += '" class="btn btn-primary btn-sm" target="_blank" alt="visualizza tutti i procedimenti">Visualizza</a></div>';
        //}
        popup_content += '<div class="col-md-5"><strong>Sito web:</strong></div>';
        popup_content += '<div class="col-md-2"></div>';
        popup_content += '<div class="col-md-5"><a href="/TipologieProcedimento/Index2?enteCf=';
        popup_content += '" class="btn btn-primary btn-sm" target="_blank" alt="visualizza tutti i procedimenti del ';
        popup_content += nome_comune + '">Link</a></div>';
        popup_content += '</div>';

        /* var contatti = '<div class="row"><div class="col-md-12"><strong>Contatti</strong></div><div class="col-md-12">sportellolavorocingoli@regione.marche.it<br> regione.marche.centroimpiegomacerata@emarche.it<br> 0733/604715-602686</div></div>';
    var pec = "non disponibile";
    if(content['Pec'] !== null)
    {
        pec = content['Pec'];
        } */

        return popup_header + popup_content;
    }
</script>
</asp:Content>