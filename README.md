# Excel Order Aggregator

Desktop tool (Tkinter + openpyxl) I wrote to manage merchandise orders for a sports club. It reads an Excel export of online orders and produces a workbook with:

- one sheet per athlete category, with each athlete's order and amount due;
- one aggregate sheet per category (item × size × color, with the names to print);
- an overall summary of items to order and money to collect;
- a separate sheet for photo orders.

## Input format

One order per row, no header row:

| Name | Surname | Athlete name | Category | Order | Payment method |
|---|---|---|---|---|---|

The *Order* cell contains one line per item, in the form produced by the shop's export, followed by the total:

```
Hoodie (Taglia: M, Colore: Blu, Amount: 30.00 EUR)
Total: 30.00 EUR
```

## Run

```bash
pip install openpyxl
python order_aggregator.py
```

The result is saved as `Totali Ordini.xlsx` in the same folder as the input file.

## Known limitations

- Aggregate counts are based on the number of personalizations, so an item ordered with quantity > 1 is counted once.
- Parsing relies on the exact text format of the shop export.
