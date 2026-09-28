import os

import openpyxl
import tkinter as tk
from tkinter import filedialog

class item:
    def __init__(self, nome: str, cognome: str, nome_atleta: str, categoria: str, ordine: str, pagamento: str) -> None:
        self.nome = nome
        self.cognome = cognome
        self.personalizzazione = nome_atleta if nome_atleta else nome 
        self.categoria = categoria
        self.pagamento = pagamento
        self.ordine = ordine.split('Total: ')[0].strip()
        self.articoli: list[article] = []
        self.totale = ordine.split('Total: ')[1].strip()
        for articolo in ordine.split('\n')[:-1]:
            self.articoli.append(match_re(nome, cognome, articolo, self.personalizzazione, categoria))

class article:
    def __init__(self, nome: str, cognome: str, nome_articolo: str, taglia: str, colore: str, personalizzazione: str, amount: str, quantità, categoria: str) -> None:
            self.nome = nome
            self.cognome = cognome
            self.nome_articolo = nome_articolo
            self.taglia = taglia
            self.colore = colore
            self.personalizzazione = personalizzazione
            self.costo = amount
            self.quantita = 0 if not quantità else int(quantità)
            self.categoria = categoria

def match_re( nome: str, cognome: str,articolo_str: str, personalizzazione: str, categoria: str) -> article:
    # Estrai informazioni dall'ordine
    nome_articolo = articolo_str.split('(')[0].strip()
    attributes = articolo_str.split('(')[1].replace(')', '').split(',')
    costo = '0'
    taglia = ''
    quantita = '1'
    colore = ''
    for attribute in attributes:
        if attribute.strip().startswith('Amount:'):
            costo = attribute.split(':')[1].strip()
        elif attribute.strip().startswith('Taglia:') or attribute.strip().startswith('Taglia dell Scarpa:'):
            taglia = attribute.split(':')[1].strip()
        elif attribute.strip().startswith('Quantità:'):
            quantita = attribute.split(':')[1].strip()
        elif attribute.strip().startswith('Colore:'):
            colore = attribute.split(':')[1].strip()
    return article(nome, cognome, nome_articolo, taglia, colore, personalizzazione, costo, quantita, categoria)

def get_items_categories_articles(workbook: openpyxl.Workbook, categorie: list[str], items: list[item], purchases: list[article], output_data: list[str]):
    sheet = workbook.active
    for row in sheet.iter_rows(values_only=True):
        nome = row[0]
        cognome = row[1]
        nome_atleta = row[2]
        categoria = row[3]
        ordine = row[4]
        pagamento = row[5]

        # Aggiungi categoria alla lista
        categorie.append(categoria)

        contenuto_ordine = item(nome, cognome, nome_atleta, categoria, ordine, pagamento)
        # Aggiungi articolo al set di chiavi
        for articolo in contenuto_ordine.articoli:
            purchases.append(articolo)

        # Aggiungi item alla lista
        items.append(contenuto_ordine)

        # Aggiungi i dati a output_data
        output_data.append(f"{nome}{cognome}{nome_atleta}{categoria}{ordine}{pagamento}")

def create_category_sheet(workbook: openpyxl.Workbook, categories: list[str], items: list[item]):
    # Crea una sheet per categoria
    for category in categories:
        catgortySheetFull = workbook.create_sheet(title=category + " atlete")
        # Crea una riga per titolo colonne
        catgortySheetFull.append(["Nome", "Cognome", "Personalizzazione", "Ordine", "Totale da pagare", "Metodo di pagamento"])
        
        # variabile per calcolare totale da raccogliere nel corso
        totale_corso = 0.0
        # Estrai solo le atlete di quella categoria
        filtered_items =  [token for token in items if token.categoria == category]
        # Crea una riga per atleta, con il suo ordine
        for item in filtered_items:
            totale_corso += float(item.totale.replace(' EUR', ''))
            catgortySheetFull.append([item.nome, item.cognome, item.personalizzazione, item.ordine, item.totale, item.pagamento])
        catgortySheetFull.append(["", "", "", "Totale da incassare", str(totale_corso) + " EUR"])

def create_aggregate_per_category_sheet(workbook: openpyxl.Workbook, categories: list[str], purchases: list[article]):
    # Crea una sheet per categoria
    for category in categories:
        # Aggiungi una nuova sheet con i dati sommati
        catgortySheetFull = workbook.create_sheet(title=category + " aggregato")
        # Crea una riga per titolo colonne
        catgortySheetFull.append(["Articolo", "Taglia", "Colore", "Totale", "Personalizzazioni"])
        
        somma_articolo = {}
        # Estrai solo le atlete di quella categoria
        filtered_purchases =  [token for token in purchases if token.categoria == category and not token.nome_articolo.startswith('Foto')]
        # Crea una riga per atleta, con il suo ordine
        for purchase in filtered_purchases:
            # Aggiungi i dati a somma_articolo
            if (purchase.nome_articolo, purchase.taglia, purchase.colore) in somma_articolo:
                somma_articolo[purchase.nome_articolo, purchase.taglia, purchase.colore] += ", " + purchase.personalizzazione
            else:
                somma_articolo[purchase.nome_articolo, purchase.taglia, purchase.colore] = purchase.personalizzazione
            
        for key, personalizzazioni in somma_articolo.items():
            nome_articolo, taglia, colore = key
            catgortySheetFull.append([nome_articolo, taglia, colore, len(personalizzazioni.split(',')), personalizzazioni])

def create_aggregate_sheet(workbook: openpyxl.Workbook, purchases: list[article]):
    # Aggiungi una nuova sheet con i dati sommati
    catgortySheetFull = workbook.create_sheet(title="Totale Ordini")
    # Crea una riga per titolo colonne
    catgortySheetFull.append(["Articolo", "Taglia", "Colore", "Totale", "Totale EUR", "Personalizzazioni"])
    
    somma_articolo_personalizzazione = {}
    somma_articolo_costo = {}
    totale = 0.0
    # Crea una riga per atleta, con il suo ordine
    filtered_purchases =  [token for token in purchases if not token.nome_articolo.startswith('Foto')]
    for purchase in filtered_purchases:
        # Aggiungi i dati a somma_articolo
        totale += float(purchase.costo.replace(' EUR', ''))
        if (purchase.nome_articolo, purchase.taglia, purchase.colore) in somma_articolo_personalizzazione:
            somma_articolo_personalizzazione[purchase.nome_articolo, purchase.taglia, purchase.colore] += ", " + purchase.personalizzazione
            somma_articolo_costo[purchase.nome_articolo, purchase.taglia, purchase.colore] +=  float(purchase.costo.replace(' EUR', ''))
        else:
            somma_articolo_personalizzazione[purchase.nome_articolo, purchase.taglia, purchase.colore] = purchase.personalizzazione
            somma_articolo_costo[purchase.nome_articolo, purchase.taglia, purchase.colore] =  float(purchase.costo.replace(' EUR', ''))
                
    for key, personalizzazioni in somma_articolo_personalizzazione.items():
        nome_articolo, taglia, colore = key
        catgortySheetFull.append([nome_articolo, taglia, colore, len(personalizzazioni.split(',')), somma_articolo_costo[nome_articolo, taglia, colore], personalizzazioni])
    
    catgortySheetFull.append(["", "", "", "Totale da incassare", str(totale) + " EUR"])

def create_foto_sheet(workbook: openpyxl.Workbook, purchases: list[article]):
    # Aggiungi una nuova sheet con i dati sommati
    catgortySheetFull = workbook.create_sheet(title="Foto natalizie")
    # Crea una riga per titolo colonne
    catgortySheetFull.append(["Nome", "Cognome", "Quantità", "Totale"])
    
    totale = 0.0

    # Estrai solo le atlete di quella categoria
    filtered_purchases =  [token for token in purchases if token.nome_articolo.startswith('Foto')]
    # Crea una riga per atleta, con il suo ordine di foto
    for purchase in filtered_purchases:
        totale += float(purchase.costo.replace(' EUR', ''))
        catgortySheetFull.append([purchase.nome, purchase.cognome, purchase.quantita, purchase.costo])
    
    catgortySheetFull.append(["", "", "Totale da incassare", str(totale) + " EUR"])


def separa_dati(file_excel):
    workbook = openpyxl.load_workbook(file_excel)

    output_data = []
    categorie = []
    items = []
    purchases = []
    get_items_categories_articles(workbook, categorie, items, purchases, output_data)
    
    create_category_sheet(workbook, set(categorie), items)
    create_aggregate_per_category_sheet(workbook, set(categorie), purchases)
    create_aggregate_sheet(workbook, purchases)
    create_foto_sheet(workbook, purchases)

    # Salva il file Excel
    output_excel = os.path.join(os.path.dirname(file_excel), "Totali Ordini.xlsx")
    workbook.save(output_excel)

    return output_data, output_excel

def seleziona_file():
    file_excel = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_excel:
        dati_separati, output_excel = separa_dati(file_excel)
        output_text.config(state=tk.NORMAL)
        output_text.delete(1.0, tk.END)
        for linea in dati_separati:
            output_text.insert(tk.END, linea + "\n")
        output_text.config(state=tk.DISABLED)
        output_excel_label.config(text=f"File Excel con somma: {output_excel}")


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Separatore Dati da File Excel con Somma")

    seleziona_file_button = tk.Button(root, text="Seleziona File Excel", command=seleziona_file)
    seleziona_file_button.pack(pady=10)

    output_text = tk.Text(root, height=10, width=50)
    output_text.config(state=tk.DISABLED)
    output_text.pack()

    output_excel_label = tk.Label(root, text="")
    output_excel_label.pack()

    root.mainloop()
