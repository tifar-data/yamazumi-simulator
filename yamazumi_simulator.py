"""
Simulador de Yamazumi em Python

Este script é um exemplo simples de como criar um gráfico Yamazumi para
analisar o balanceamento de estações de trabalho segundo o conceito
japonês. O Yamazumi agrupa as atividades por estação e categoria de
atividade (VA – valor agregado, NVA – não valor agregado, MUDA –
desperdício) para visualizar o tempo total de cada tipo de trabalho
em cada estação e compará-lo com o takt time.

Como funciona
-------------
1. **Entrada**: a planilha Excel deve conter as colunas:
   - ``Estacao``: número ou nome da estação.
   - ``Tempo``: tempo gasto na atividade (em segundos ou minutos).
   - ``Categoria``: classificação da atividade (``VA``, ``NVA`` ou ``MUDA``).

2. **Takt time**: é o ritmo de produção necessário para atender à demanda.
   Pode ser fornecido em minutos com o argumento ``--takt``. Se não
   fornecido, usa-se a média dos tempos totais como referência.

3. **Processamento**: o script soma os tempos de cada categoria por
   estação, calcula o total, desenha um gráfico de barras empilhadas
   (Yamazumi) com cores distintas por categoria e traça uma linha
   horizontal para o takt time.

4. **Saída**: gera uma figura de alta resolução e apresenta no
   console um resumo dos tempos por estação, incluindo a diferença em
   relação ao takt time e a proporção de tempo de valor agregado (VA%).

Para utilizar:

```
python yamazumi_simulator.py --entrada dados_yamazumi.xlsx --saida grafico_yamazumi.png --takt 3.5
```

Autor: ChatGPT
Data: 2025-08-31
"""

from __future__ import annotations

import argparse
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt


def ler_dados_yamazumi(arquivo: str | Path, unidade: str | None = None) -> pd.DataFrame:
    """Lê dados para o Yamazumi a partir de uma planilha Excel.

    A planilha deve conter as colunas 'Estacao', 'Tempo' e 'Categoria'. A função
    converte os tempos para segundos e garante que a categoria esteja em
    maiúsculas.

    Parameters
    ----------
    arquivo : str or Path
        C:\Users\tiago\Desktop\Python Tifar\Balanceamento de linha.
    unidade : str, opcional
        Se ``'minutos'`` ou ``'segundos'``. Se não informado, infere pela coluna ``Tempo``.

    Returns
    -------
    DataFrame
        DataFrame com colunas 'Estacao', 'Tempo_s', 'Categoria'.
    """
    arquivo = Path(arquivo)
    if not arquivo.exists():
        raise FileNotFoundError(f"Arquivo {arquivo} não encontrado")

    df = pd.read_excel(arquivo)
    required_cols = {"estacao", "tempo", "categoria"}
    cols_lower = {c.lower().strip() for c in df.columns}
    if not required_cols.issubset(cols_lower):
        raise ValueError("A planilha deve ter as colunas 'Estacao', 'Tempo' e 'Categoria'")

    # Normalizar nomes
    col_est = [c for c in df.columns if c.lower().strip() == "estacao"][0]
    col_time = [c for c in df.columns if c.lower().strip() == "tempo"][0]
    col_cat = [c for c in df.columns if c.lower().strip() == "categoria"][0]

    tempos = df[col_time].astype(float)
    if unidade:
        unidade = unidade.lower()
        if unidade.startswith("min"):
            tempos_s = tempos * 60.0
        elif unidade.startswith("seg"):
            tempos_s = tempos
        else:
            raise ValueError("Unidade deve ser 'minutos' ou 'segundos'")
    else:
        # Inferir pela unidade do nome
        if "(min" in col_time.lower():
            tempos_s = tempos * 60.0
        else:
            tempos_s = tempos

    categorias = df[col_cat].astype(str).str.upper()
    df_res = pd.DataFrame({
        "Estacao": df[col_est],
        "Tempo_s": tempos_s,
        "Categoria": categorias,
    })
    return df_res


def gerar_grafico_yamazumi(df: pd.DataFrame, takt_s: float | None, saida: str | Path) -> None:
    """Gera um gráfico Yamazumi (barras empilhadas) para as estações.

    Parameters
    ----------
    df : DataFrame
        DataFrame com colunas 'Estacao', 'Tempo_s' e 'Categoria'.
    takt_s : float or None
        Takt time em segundos. Se None, usa a média dos tempos totais como referência.
    saida : str or Path
        Caminho para salvar o gráfico.
    """
    # Determinar ordem das categorias
    categorias_unicas = ["VA", "NVA", "MUDA"]
    # Calcular soma por estação e categoria
    pivot = df.pivot_table(index="Estacao", columns="Categoria", values="Tempo_s", aggfunc="sum", fill_value=0.0)
    pivot = pivot.reindex(columns=categorias_unicas, fill_value=0.0)
    estacoes = pivot.index.astype(str)
    
    # Calcular takt se não fornecido
    total_por_estacao = pivot.sum(axis=1)
    if takt_s is None:
        takt_s = total_por_estacao.mean()

    fig, ax = plt.subplots(figsize=(16, 8))
    # Plotagem empilhada
    bottom = None
    colors = {"VA": "#1f77b4", "NVA": "#ff7f0e", "MUDA": "#d62728"}
    for cat in categorias_unicas:
        ax.bar(estacoes, pivot[cat], bottom=bottom, color=colors.get(cat, None), edgecolor="black", label=cat)
        if bottom is None:
            bottom = pivot[cat]
        else:
            bottom = bottom + pivot[cat]

    # Linha do takt time
    ax.axhline(takt_s, color="green", linestyle="--", linewidth=2, label=f"Takt {takt_s/60:.2f} min")

    # Anotações de gargalo
    max_est = total_por_estacao.idxmax()
    max_val = total_por_estacao.max()
    ax.annotate(f"Gargalo: Estação {max_est} ({max_val:.0f}s)",
                (list(estacoes).index(str(max_est)), max_val),
                textcoords="offset points", xytext=(0, 10), ha='center', color="red", fontsize=9)

    ax.set_title("Gráfico Yamazumi – Distribuição de Tempo por Estação")
    ax.set_xlabel("Estação")
    ax.set_ylabel("Tempo (s)")
    ax.legend(loc="upper right")
    plt.xticks(rotation=90)
    plt.tight_layout()
    saida = Path(saida)
    saida.parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(saida, dpi=300, bbox_inches="tight")
    plt.close()
    # Imprimir resumo, incluindo proporção de tempo VA
    print(f"Takt time considerado: {takt_s/60:.2f} minutos ({takt_s:.0f} s)")
    print("Resumo por estação:")
    # Calcular proporção de VA
    va = pivot["VA"] if "VA" in pivot.columns else pd.Series([0]*len(pivot), index=pivot.index)
    va_ratio = va / total_por_estacao.replace(0, 1)
    for est in pivot.index:
        total = total_por_estacao[est]
        diff = total - takt_s
        status = "acima" if diff > 0 else "abaixo"
        ratio = va_ratio[est] * 100
        print(f"  Estação {est}: {total:.0f} s ( {abs(diff):.0f} s {status} do takt), VA% = {ratio:.1f}%")


def main() -> None:
    parser = argparse.ArgumentParser(description="Cria um gráfico Yamazumi a partir de dados de estação, categoria e tempo")
    parser.add_argument("--entrada", "-i", required=True, help="Arquivo Excel com colunas Estacao, Tempo e Categoria")
    parser.add_argument("--saida", "-o", required=False, default="yamazumi.png", help="Arquivo de saída do gráfico")
    parser.add_argument("--unidade", "-u", choices=["minutos", "segundos"], default=None, help="Unidade de tempo na planilha")
    parser.add_argument("--takt", type=float, default=None, help="Takt time em minutos (opcional). Se omitido, usa a média dos tempos totais.")
    args = parser.parse_args()

    df = ler_dados_yamazumi(args.entrada, unidade=args.unidade)
    takt_s = args.takt * 60.0 if args.takt is not None else None
    gerar_grafico_yamazumi(df, takt_s, args.saida)


if __name__ == "__main__":
    main()