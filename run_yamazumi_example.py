# -*- coding: utf-8 -*-
"""
Simulador Yamazumi - Geração de gráfico a partir de planilha Excel
"""

from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt


def ler_dados_yamazumi(arquivo: str | Path, unidade: str | None = None) -> pd.DataFrame:
    """Lê dados com colunas 'Estacao', 'Tempo' e 'Categoria' e normaliza."""
    arquivo = Path(arquivo)
    if not arquivo.exists():
        raise FileNotFoundError(f"Arquivo {arquivo} não encontrado")

    df = pd.read_excel(arquivo)
    required = {"estacao", "tempo", "categoria"}
    cols_lower = {c.lower().strip() for c in df.columns}
    if not required.issubset(cols_lower):
        raise ValueError("A planilha deve ter as colunas 'Estacao', 'Tempo' e 'Categoria'")

    col_est = [c for c in df.columns if c.lower().strip() == "estacao"][0]
    col_time = [c for c in df.columns if c.lower().strip() == "tempo"][0]
    col_cat = [c for c in df.columns if c.lower().strip() == "categoria"][0]

    tempos = df[col_time].astype(float)
    if unidade:
        if unidade.lower().startswith("min"):
            tempos_s = tempos * 60
        else:
            tempos_s = tempos
    else:
        tempos_s = tempos * 60 if "(min" in col_time.lower() else tempos

    categorias = df[col_cat].astype(str).str.upper()
    df_res = pd.DataFrame({
        "Estacao": df[col_est],
        "Tempo_s": tempos_s,
        "Categoria": categorias
    })

    return df_res


def gerar_grafico_yamazumi(df: pd.DataFrame, takt_s: float | None, saida: str | Path) -> None:
    """Gera gráfico Yamazumi com barras empilhadas."""
    agrupado = df.groupby(["Estacao", "Categoria"])["Tempo_s"].sum().unstack(fill_value=0)
    total_por_estacao = agrupado.sum(axis=1)
    va_por_estacao = agrupado.get("VA", 0)
    estacoes = agrupado.index.tolist()

    # Plot
    fig, ax = plt.subplots(figsize=(16, 8))
    cores = {"VA": "#2ca02c", "NVA": "#ff7f0e", "MUDA": "#d62728"}
    agrupado.plot(kind="bar", stacked=True, ax=ax, color=[cores.get(c, "#ccc") for c in agrupado.columns])

    if takt_s:
        ax.axhline(takt_s, color="blue", linestyle="--", linewidth=2, label=f"Takt: {takt_s:.0f}s")

    # Títulos e legendas
    ax.set_title("Gráfico Yamazumi - Estações de Trabalho", fontsize=16)
    ax.set_xlabel("Estação")
    ax.set_ylabel("Tempo (s)")
    ax.legend(loc="upper right")

    # Destaque do gargalo
    gargalo_idx = total_por_estacao.idxmax()
    ax.annotate("Gargalo", xy=(estacoes.index(gargalo_idx), total_por_estacao.max()),
                xytext=(estacoes.index(gargalo_idx), total_por_estacao.max() + 20),
                arrowprops=dict(facecolor='black', shrink=0.05))

    plt.tight_layout()
    fig.savefig(saida, dpi=300)
    plt.close()

    # Resumo por estação
    print(f"Takt time considerado: {takt_s / 60:.2f} minutos ({takt_s:.0f} s)" if takt_s else "Takt não fornecido.")
    print("Resumo por estação:")
    for est in estacoes:
        total = total_por_estacao[est]
        va = va_por_estacao.get(est, 0)
        delta = f"{total - takt_s:+.0f}" if takt_s else "n/a"
        va_pct = 100 * va / total if total else 0
        print(f"  Estação {est}: {total:.0f}s ({delta}s em relação ao takt), VA% = {va_pct:.1f}%")


# Script de execução com seletor de arquivo
if __name__ == "__main__":
    import tkinter as tk
    from tkinter import filedialog
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    caminho = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Planilhas Excel", "*.xlsx")])
    if not caminho:
        messagebox.showerror("Erro", "Nenhum arquivo selecionado.")
        exit()

    # Rodar análise com takt fixo (por exemplo, 3.5 minutos)
    takt_min = 3.5
    df = ler_dados_yamazumi(caminho)
    gerar_grafico_yamazumi(df, takt_min * 60.0, "yamazumi_output.png")
    print("Gráfico salvo como yamazumi_output.png")
