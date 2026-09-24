"""Exportador de planilhas Excel para arquivos PDF formatados."""

import matplotlib.pyplot as plt
import pandas as pd


def exportar_para_pdf(excel_path: str, pdf_output_path: str):
    """Exporta os dados de uma planilha Excel para um arquivo PDF formatado."""
    df = pd.read_excel(excel_path, skiprows=1)  # Pula o título principal

    fig, ax = plt.subplots(figsize=(11.69, 8.27))  # A4 Landscape
    ax.axis("tight")
    ax.axis("off")

    title = excel_path.split("/")[-1].replace(".xlsx", "").replace("_", " ").upper()
    fig.suptitle(title, fontsize=15, fontweight="bold", y=0.96, color="#1A365D")

    headers = list(df.columns)
    data_matrix = [headers] + df.values.tolist()

    table = ax.table(
        cellText=data_matrix,
        cellLoc="center",
        loc="center",
        colWidths=[0.08, 0.08, 0.07, 0.17, 0.17, 0.17, 0.17, 0.17],
    )
    table.auto_set_font_size(False)
    table.set_fontsize(8.5)

    # Estilização das células
    for (r, _), cell in table.get_celld().items():
        cell.set_height(0.026)
        if r == 0:
            cell.set_facecolor("#1A365D")
            cell.get_text().set_color("white")
            cell.get_text().set_fontweight("bold")
        else:
            dia_val = data_matrix[r][1]
            if "Sábado" in str(dia_val) or "Domingo" in str(dia_val):
                cell.set_facecolor("#EDF2F7")
                cell.get_text().set_fontweight("bold")

    fig.subplots_adjust(left=0.03, right=0.97, top=0.90, bottom=0.03)
    fig.savefig(pdf_output_path, format="pdf", dpi=300)
    plt.close(fig)
