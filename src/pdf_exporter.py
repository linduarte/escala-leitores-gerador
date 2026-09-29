"""Exportador de planilhas Excel para arquivos PDF formatados."""

import json
import re
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


def exportar_para_pdf(excel_path: str, pdf_output_path: str):
    """Exporta uma planilha Excel para PDF, substituindo números por nomes."""
    mapa_leitores = _carregar_mapa_leitores()

    # Ajuste o 'header' conforme a linha onde estão os títulos das colunas na planilha (0, 1 ou 2)
    df = pd.read_excel(excel_path, header=1)

    # Substituir números por nomes
    def buscar_nome(celula):
        if pd.isna(celula):
            return ""
        if isinstance(celula, (int, float)):
            chave = str(int(celula))
        else:
            chave = str(celula).strip()
        return mapa_leitores.get(chave, celula)

    df_formatado = df.map(buscar_nome)

    # Renderizar PDF
    _, ax = plt.subplots(figsize=(11.69, 8.27)) # A4 Landscape
    ax.axis("tight")
    ax.axis("off")

    tabela = ax.table(
        cellText=df_formatado.astype(str).values.tolist(),
        colLabels=df_formatado.columns.astype(str).tolist(),
        loc="center",
        cellLoc="center"
    )
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(8)
    tabela.scale(1.2, 1.2)

    plt.savefig(pdf_output_path, bbox_inches="tight", format="pdf")
    plt.close()


def _carregar_mapa_leitores():
    """Carrega o mapeamento entre os números e os nomes dos leitores."""
    json_path = Path("data/leitores.json")
    mapa_leitores = {}
    if not json_path.exists():
        return mapa_leitores

    with open(json_path, "r", encoding="utf-8") as arquivo:
        for item in json.load(arquivo):
            valor = list(item.values())[0] if isinstance(item, dict) else ""
            match = re.search(r"\|\s*(\d+)\s*\|\s*([^|]+)\s*\|", valor)
            if match:
                mapa_leitores[match.group(1).strip()] = match.group(2).strip()
    return mapa_leitores