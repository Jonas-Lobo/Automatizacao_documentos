!pip install python-docx

from docx import Document
from datetime import datetime 
import pandas as pd

tabela = pd.read_excel("Teste 1.xlsx")

for linha in tabela.index:
    documento = Document("Notificação_.docx")

    entidade = tabela.loc[linha, "Entidade"]
    item1 = tabela.loc[linha, "Associado"]
    item2 = tabela.loc[linha, "CNPJ"]
    item3 = tabela.loc[linha, "indenização"]
    item4 = tabela.loc[linha, "Total"]

    referencias = {
        "XXXX": entidade,
        "YYYY": item1,
        "ZZZZ": item2,
        "VVVV": str(item3),
        "TTTT": str(item4),
        "DD": str(datetime.now().day),
        "MM": str(datetime.now().month),
        "AAAA": str(datetime.now().year), 
    }

    for paragrafo in documento.paragraphs:
        for codigo in referencias:
            valor = referencias[codigo]
            paragrafo.text = paragrafo.text.replace(codigo, valor)

    documento.save(f"Notificação_ -{entidade}.docx")