# Report Shopify colori via Streamlit

App Streamlit per estrarre ordini Shopify e creare un report quantità per colore, utile quando il colore è salvato come proprietà personalizzata del line item, ad esempio tramite VOPO / product options app.

## Cosa fa

- Filtra gli ordini per intervallo date.
- Filtra uno o più prodotti per titolo.
- Cerca il colore in:
  1. `customAttributes` / line item properties, tipico per opzioni personalizzate;
  2. `variant.selectedOptions`, se il colore è una variante Shopify nativa;
  3. `variantTitle`, come fallback.
- Raggruppa la quantità totale per prodotto e colore.
- Esporta Excel con:
  - `Riepilogo`
  - `Dettaglio ordini`

## File

```text
app.py
requirements.txt
.streamlit/secrets.example.toml
.gitignore
```

## Secrets Streamlit

In locale puoi creare `.streamlit/secrets.toml`:

```toml
SHOPIFY_SHOP_DOMAIN = "tuo-negozio.myshopify.com"
SHOPIFY_ACCESS_TOKEN = "shpat_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
```

Su Streamlit Community Cloud inserisci gli stessi valori in **Advanced settings > Secrets**.

## Permessi Shopify necessari

L'app deve poter leggere gli ordini via Shopify Admin API.

Minimo:
- `read_orders`

Se devi leggere ordini più vecchi del limite standard di Shopify, potrebbe servirti anche:
- `read_all_orders`

## Avvio locale

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy

1. Crea un repository GitHub.
2. Carica questi file.
3. Vai su Streamlit Community Cloud.
4. Collega il repository.
5. Imposta i secrets.
6. Deploy.

## Nota VOPO

Se il report non trova colori, apri l'espander **Debug: campi colore trovati nei line item**. Guarda il JSON in `Custom attributes` e aggiungi il nome esatto del campo nella casella "Nomi possibili del campo colore".
