# Report Shopify colori via Streamlit

App Streamlit per estrarre ordini Shopify e creare un report quantità per colore, utile quando il colore è salvato come proprietà personalizzata del line item, ad esempio tramite VOPO / product options app.

## Cosa fa

- Filtra gli ordini per intervallo date.
- Filtra uno o più prodotti per titolo.
- Cerca il colore in:
  1. `customAttributes` / line item properties, tipico per opzioni personalizzate;
  2. `variantTitle`, come fallback.

Questa versione evita i campi `product` e `variant`, quindi non richiede `read_products`. Se vuoi leggere anche opzioni variante native Shopify, aggiungi lo scope `read_products` e usa la versione completa.
- Raggruppa la quantità totale per colore, sommando insieme tutti i prodotti filtrati.
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


## Moltiplicazione pezzi per prodotto

Questa versione calcola la quantità reale dei pezzi moltiplicando la quantità ordinata per il numero iniziale nel nome prodotto.

Esempi:

| Prodotto | Quantità ordinata | Totale calcolato |
|---|---:|---:|
| 10 T-shirt Economy | 3 | 30 |
| 20 T-shirt Economy | 17 | 340 |

Se un prodotto non inizia con un numero, il moltiplicatore usato è 1.


## Riepilogo unificato per colore

Questa versione non distingue più il riepilogo finale per prodotto.

Esempio:

| Prodotto | Quantità ordine | Pezzi per prodotto | Colore | Totale calcolato |
|---|---:|---:|---|---:|
| 10 T-shirt Economy | 3 | 10 | Black | 30 |
| 20 T-shirt Economy | 17 | 20 | White | 340 |
| 20 T-shirt Premium | 2 | 20 | Black | 40 |

Il riepilogo finale sarà:

| Colore | Quantità totale |
|---|---:|
| Black | 70 |
| White | 340 |

Il foglio `Dettaglio ordini` mantiene comunque il nome del prodotto, così puoi verificare da dove arrivano i totali.
