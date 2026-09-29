"""Exportador de planilhas Excel para arquivos PDF formatados."""

import json
import re
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


def _carregar_mapa_leitores(json_path: Path) -> dict[str, str]:
    """Carrega o mapeamento de números para nomes dos leitores."""
    if not json_path.exists():
        return {}

    with open(json_path, "r", encoding="utf-8") as arquivo:
        dados = json.load(arquivo)

    mapa_leitores: dict[str, str] = {}
    for item in dados:
        valor = list(item.values())[0] if isinstance(item, dict) else ""
        match = re.search(r"\|\s*(\d+)\s*\|\s*([^|]+)\s*\|", valor)
        if match:
            mapa_leitores[match.group(1).strip()] = match.group(2).strip()

    return mapa_leitores


def _buscar_nome(celula, mapa_leitores: dict[str, str]):
    """Substitui valores numéricos por nomes quando houver correspondência."""
    if pd.isna(celula):
        return ""

    chave = (
        str(int(celula)) if isinstance(celula, (int, float)) else str(celula).strip()
    )
    return mapa_leitores.get(chave, celula)


def exportar_para_pdf(excel_path: str, pdf_output_path: str):
    """Exporta uma planilha Excel para um PDF com tabela formatada."""
    mapa_leitores = _carregar_mapa_leitores(Path("data/leitores.json"))

    df_titulo = pd.read_excel(excel_path, nrows=1, header=None)
    titulo_raw = (
        str(df_titulo.iloc[0, 0]) if not df_titulo.empty else "ESCALA DE LEITORES"
    )
    titulo_pagina = (
        titulo_raw.replace("(GABARITO)", "").replace("OCTOBER", "OUTUBRO").strip()
    )

    df = pd.read_excel(excel_path, header=1)
    df_formatado = df.map(lambda celula: _buscar_nome(celula, mapa_leitores))

    fig, ax = plt.subplots(figsize=(11.69, 8.27))
    fig.suptitle(
        titulo_pagina,
        fontsize=16,
        fontweight="bold",
        color="#1f4e78",
        y=0.92,
    )
    ax.axis("off")

    tabela = ax.table(
        cellText=df_formatado.astype(str).values.tolist(),
        colLabels=list(df_formatado.columns),
        loc="center",
        cellLoc="center",
    )
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(9)
    tabela.scale(1.2, 1.6)

    for (linha, _), cell in tabela.get_celld().items():
        if linha == 0:
            cell.set_facecolor("#1f4e78")
            cell.get_text().set_color("white")
            cell.get_text().set_fontweight("bold")
            cell.get_text().set_fontsize(10)

    plt.tight_layout(rect=(0.02, 0.02, 0.98, 0.88))

    output_path = Path(pdf_output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(pdf_output_path, format="pdf", dpi=300)
    plt.close()
