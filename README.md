# Gerador de Escalas e Gabaritos Paroquiais

Este projeto é uma ferramenta CLI desenvolvida em Python para automatizar a geração de gabaritos em Excel e a exportação das escalas mensais de leitores formatadas em PDF.

---

## 🛠️ Tecnologias e Ferramentas

* **Python** (via `uv` para gestão de dependências e ambiente virtual)
* **Jujutsu (`jj`)** para versionamento
* **Click** para interface de linha de comando (CLI)
* **openpyxl** & **pandas** para manipulação de planilhas
* **matplotlib** para exportação de PDFs

---

## 🚀 Como Usar

### 1. Gerar Gabarito do Mês (Planilha em Branco)

Para criar a planilha gabarito `.xlsx` do mês desejado, execute:

```powershell
uv run escala gabarito --mes 10 --ano 2026
```

Dica: Se o argumento --out não for informado, o arquivo será gerado automaticamente em data/gabarito_escala_10_2026.xlsx. Caso queira especificar o nome ou local, use:

```powershell
uv run escala gabarito --mes 10 --ano 2026 --out data/gabarito_outubro_2026.xlsx
```
2. Preenchimento do Gabarito
Abra a planilha gerada na pasta data/ e preencha as escalas associando os números correspondentes aos leitores (ou o nome direto de convidados).

3. Gerar PDF Final Formatado
Após o preenchimento da planilha, execute o comando de exportação para converter os números nos nomes reais (baseado no data/leitores.json) e gerar o PDF final em formato A4 Landscape:

```powershell
uv run escala pdf --input data/gabarito_outubro_2026.xlsx --out output/escala_outubro_2026.pdf
```

📂 Estrutura de Arquivos Locais
data/leitores.json: Mapeamento local dos números e nomes dos leitores (ignorado no Git).

data/*.xlsx: Gabaritos gerados e preenchidos (ignorados no Git).

output/*.pdf: Escalas finais geradas em PDF para impressão.

