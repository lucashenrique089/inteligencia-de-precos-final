"""
GERA A LISTA DE PNEUS DO PAINEL | Roma Pneus
=============================================
Para que serve: transforma a planilha de monitoramento (a que você preenche e o
validar_links.py confere) no arquivo produtos_monitorados.csv, que o coletor e o painel usam.

O que faz:
  1. Lê a aba "Itens" da planilha (a conferida, se existir; senão a original)
  2. Para cada loja, usa o LINK SUGERIDO quando o seu estava vazio, errado ou era de busca
     (se a conferência disse "ok" ou "PARECIDO", mantém o seu)
  3. Junta pneus repetidos (o mesmo produto em Varejo e Distribuição vira um só, somando as vendas)
  4. Mantém o número (item) dos 14 pneus antigos, para o histórico de preços continuar valendo
  5. Salva produtos_monitorados.csv (faz um backup do anterior antes)

Como usar:
    python gerar_lista.py "Base_monitoramento_itens_conferida.xlsx"
"""

import os
import re
import shutil
import sys

import pandas as pd

ARQUIVO_SAIDA = "produtos_monitorados.csv"
MAIS_PESQUISADOS = {"924044", "924051", "924045", "923993"}      # os 4 "mais pesquisados" do GA4
TROCAR_QUANDO = ("ERRADO", "LINK DIFERENTE", "link de busca", "link estava vazio", "link não abriu")


def texto(v):
    return "" if pd.isna(v) else str(v).strip()


def numero(v):
    t = texto(v).replace(".0", "")
    return t if t.isdigit() else ""


def escolher(original, conferencia, sugerido):
    """Usa o link sugerido só quando o original estava vazio, errado ou era de busca."""
    if sugerido and (not original or conferencia.startswith(TROCAR_QUANDO)):
        return sugerido
    return original


TROCAS = {"EFFIGRIP": "EfficientGrip", "EFFICIENTGRIP": "EfficientGrip", "WRL": "Wrangler",
          "PERF": "Performance", "PERFORM": "Performance", "PERFORMAN": "Performance",
          "MARAT": "Marathon", "MARA": "Marathon", "PAPALEGUA": "Papa-Légua", "ASYMMETRIC": "Asymmetric"}


def bonito(palavra):
    """EDGE -> Edge; XL, SUV, 86H, 118/116R continuam em maiúsculas."""
    if palavra in TROCAS:
        return TROCAS[palavra]
    if any(c.isdigit() for c in palavra) or len(palavra) <= 3:
        return palavra
    return palavra.title()


def medida_e_modelo(produto):
    p = re.sub(r"\s+", " ", produto.upper()).strip()
    m = re.match(r"(\S+)\s+(.*)", p)
    if not m:
        return p, ""
    return m.group(1), " ".join(bonito(w) for w in m.group(2).split())


def main(arquivo):
    itens = pd.read_excel(arquivo, sheet_name="Itens", dtype=str)
    tem = lambda c: c in itens.columns
    linhas = []
    for _, r in itens.iterrows():
        produto = texto(r["Produto (base de vendas)"])
        url_roma = escolher(texto(r["URL Roma"]), texto(r.get("Conferência Roma")),
                            texto(r.get("Link sugerido Roma")))
        url_ps = escolher(texto(r["URL Pneu Store"]), texto(r.get("Conferência Pneu Store")),
                          texto(r.get("Link sugerido Pneu Store")))
        url_ac = escolher(texto(r["URL Achei Pneus"]), texto(r.get("Conferência Achei")),
                          texto(r.get("Link sugerido Achei")))
        sku_ac = numero(r.get("SKU Achei (descoberto)")) if tem("SKU Achei (descoberto)") else ""
        sku_ac = sku_ac or numero(r["SKU Achei Pneus"])
        cod = (numero(r.get("Código Roma (site)")) if tem("Código Roma (site)") else "") or numero(r["Código Roma"])
        obs = texto(r["Observação"]).lower() + " " + texto(r.get("Conferência Achei")).lower()
        if "MARATHON 2" in produto.upper() and "marathon-3" in texto(r["URL Achei Pneus"]).lower():
            obs += " alternativo"            # a Achei só tem o Cargo Marathon 3 nessa medida
        medida, modelo = medida_e_modelo(produto)
        linhas.append({
            "produto": produto, "medida": medida, "modelo": modelo, "sku_roma": cod,
            "canal": texto(r["Canal"]), "qtd_vendida": int(float(texto(r["Qtd vendida (mai–out)"]) or 0)),
            "tipo": texto(r["Tipo"]), "no_site_roma": texto(r["No site da Roma?"]),
            "roma_preco_cheio": texto(r["Preço no site (R$)"]),
            "url_roma": url_roma, "url_pneustore": url_ps, "url_acheipneus": url_ac, "sku_acheipneus": sku_ac,
            "equivalencia_pneustore": "mesmo produto" if url_ps else "não vende",
            "equivalencia_acheipneus": ("alternativo" if "alternativo" in obs or "parecido" in obs else "mesmo produto")
                                       if (sku_ac or url_ac) else "não vende",
        })
    lista = pd.DataFrame(linhas)

    # Junta repetidos: mesmo código da Roma, ou mesmos links nos dois concorrentes
    lista["chave"] = lista.apply(lambda l: l["sku_roma"] or f"{l['url_pneustore']}|{l['url_acheipneus']}", axis=1)
    agrupado = lista.groupby("chave", sort=False).agg(
        {**{c: "first" for c in lista.columns if c not in ("chave", "canal", "qtd_vendida")},
         "canal": lambda s: " e ".join(dict.fromkeys(s)), "qtd_vendida": "sum"}).reset_index(drop=True)

    # Numeração: os itens antigos mantêm o número; os novos continuam a partir do maior
    antigos = {}
    if os.path.exists(ARQUIVO_SAIDA):
        anterior = pd.read_csv(ARQUIVO_SAIDA, sep=";", dtype=str)
        antigos = dict(zip(anterior["sku_roma"].fillna("").str.replace(".0", "", regex=False), anterior["item"]))
        shutil.copy(ARQUIVO_SAIDA, ARQUIVO_SAIDA.replace(".csv", "_backup.csv"))
    proximo = max([int(i) for i in antigos.values()] + [0]) + 1
    numeros = []
    for sku in agrupado["sku_roma"]:
        if sku and sku in antigos:
            numeros.append(int(antigos[sku]))
        else:
            numeros.append(proximo)
            proximo += 1
    agrupado.insert(0, "item", numeros)
    if antigos:   # itens antigos mantêm o nome de antes (o histórico de preços usa esse nome)
        nomes = anterior.assign(item=anterior["item"].astype(int)).set_index("item")[["medida", "modelo"]]
        for i, l in agrupado.iterrows():
            if l["item"] in nomes.index:
                agrupado.loc[i, ["medida", "modelo"]] = nomes.loc[l["item"]].values
    agrupado["grupo"] = agrupado["sku_roma"].map(lambda s: "Mais pesquisados" if s in MAIS_PESQUISADOS else "Portfólio")
    agrupado["roma_disponivel"] = agrupado["no_site_roma"].map(lambda s: "sim" if s == "Sim" else "nao")
    agrupado = agrupado.sort_values("item")
    colunas = ["item", "grupo", "medida", "modelo", "sku_roma", "roma_preco_cheio", "roma_disponivel",
               "equivalencia_pneustore", "url_roma", "url_pneustore", "url_acheipneus", "sku_acheipneus",
               "equivalencia_acheipneus", "canal", "qtd_vendida", "tipo", "no_site_roma", "produto"]
    agrupado[colunas].to_csv(ARQUIVO_SAIDA, sep=";", index=False, encoding="utf-8-sig")

    print(f"Pronto! {len(agrupado)} pneus em {ARQUIVO_SAIDA} (backup do anterior: produtos_monitorados_backup.csv)")
    print(f"   Varejo: {agrupado['canal'].str.contains('Varejo').sum()} · "
          f"Distribuição: {agrupado['canal'].str.contains('Distribuição').sum()} · "
          f"no site da Roma: {(agrupado['no_site_roma'] == 'Sim').sum()}")
    print(f"   Sem link na Pneu Store: {(agrupado['url_pneustore'] == '').sum()} · "
          f"sem SKU na Achei: {(agrupado['sku_acheipneus'] == '').sum()}")


if __name__ == "__main__":
    padrao = "Base_monitoramento_itens_conferida.xlsx"
    main(sys.argv[1] if len(sys.argv) > 1 else (padrao if os.path.exists(padrao) else "Base_monitoramento_itens.xlsx"))
