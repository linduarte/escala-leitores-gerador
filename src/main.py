"""Command-line interface for generating reader schedule templates."""

import datetime

import click

from .generator import gerar_gabarito_mes
from .pdf_exporter import exportar_para_pdf


@click.group()
def cli():
    """Gerador de Escalas e Gabaritos Paroquiais"""


@cli.command()
@click.option(
    "--mes", type=int, default=datetime.datetime.now().month, help="Mês (1-12)"
)
@click.option(
    "--ano", type=int, default=datetime.datetime.now().year, help="Ano (ex: 2026)"
)
@click.option(
    "--out", type=str, default=None, help="Caminho de saída para o arquivo .xlsx"
)
def gabarito(mes, ano, out):
    """Gera uma planilha gabarito em branco para o mês/ano especificado."""
    if not out:
        out = f"data/gabarito_escala_{mes:02d}_{ano}.xlsx"
    gerar_gabarito_mes(ano, mes, out)


@cli.command()
@click.option(
    "--input",
    "input_path",
    required=True,
    type=str,
    help="Caminho da planilha .xlsx preenchida",
)
@click.option(
    "--out",
    "out_path",
    required=True,
    type=str,
    help="Caminho para salvar o arquivo .pdf",
)
def pdf(input_path, out_path):
    """Exporta a planilha preenchida para um PDF formatado (A4 Landscape, 1 página)."""
    exportar_para_pdf(input_path, out_path)


if __name__ == "__main__":
    cli()
