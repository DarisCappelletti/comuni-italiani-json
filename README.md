# Comuni italiani json

Ho creato una pagina che permette di effettuare il download di un file json generato in base al file xls impostato sul sito dell'ISTAT su questo <a href="https://www.istat.it/storage/codici-unita-amministrative/Elenco-comuni-italiani.xls">link</a>.
Essendo un link fisso è possibile ottenere sempre un file aggiornato collegandosi a questa <a href="https://www.dariscappelletti.com/comuni-italiani-json/index">pagina<a/>.

# IN PROGRESS - Creazione API

L'idea alla base di queste API e permettere l'estrapolazione dei dati dei comuni italiani da più fonti in modo da avere sempre dei dati completi.
Attualmente ho implementato i dati di queste piattaforme:
- Istat (per l'elenco aggiornato dei comuni)
- IndicePA (per i dettagli specifici di un comune)
- Wikidata (per le coordinate dei comuni)

Attualmente sono presenti 2 chiamate ma prossimamente le unificherò in modo da poter gestire il tutto con una singola chiamata.

## GET api/ComuniItaliani
[`https://www.dariscappelletti.com/api/ComuniItaliani`](https://www.dariscappelletti.com/api/ComuniItaliani)

Dati presi da: Istat, IndicePA

Parametri:
- <strong>Comune (not null)</strong>: Indicando il nome del comune verranno estrapolati i dati di quel comune

Esempio di chiamata: 
[`https://www.dariscappelletti.com/api/ComuniItaliani?comune=ripe%20san%20ginesio`](https://www.dariscappelletti.com/api/ComuniItaliani?comune=ripe%20san%20ginesio)

## GET api/ComuniItalianiCoordinate
[`https://www.dariscappelletti.com/api/ComuniItalianiCoordinate`](https://www.dariscappelletti.com/api/ComuniItalianiCoordinate)

Dati presi da: Istat, WikiData

Parametri:
- <strong>Comune</strong>: Indicando il nome del comune verranno estrapolate le coordinate di quel comune<br>
Non impostando il comune verranno estrapolate le coordinate di tutti i comuni italiani (in questo caso la chiamata impiegherà molto tempo)

Esempio di chiamata: 
[`https://www.dariscappelletti.com/api/ComuniItalianiCoordinate?comune=ripe%20san%20ginesio`](https://www.dariscappelletti.com/api/ComuniItalianiCoordinate?comune=ripe%20san%20ginesio)
