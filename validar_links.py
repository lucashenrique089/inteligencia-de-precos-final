"""
CONFERE E COMPLETA OS LINKS DA LISTA DE PNEUS | Roma Pneus
============================================================
Para que serve: a lista de pneus monitorados foi montada à mão (link por link), e é fácil
errar uma medida ou colar o link do produto vizinho. Este script entra em cada site e confere.

O que faz, para cada pneu da planilha:
  ROMA       -> procura pelo CÓDIGO do produto (o mais confiável). Se não tiver código, tenta pelo
                link. Diz se o link está certo e sugere o link correto quando não estiver.
  ACHEI      -> confere se o link abre o pneu da medida certa e descobre o SKU sozinho
                (é o SKU que o painel usa para ler preço e parcelas). Sem link: procura pelo nome.
  PNEU STORE -> abre a página e confere a medida no título. Sem link (ou link de busca):
                procura na busca do site.

Resultado: uma cópia da planilha com colunas novas de CONFERÊNCIA e LINK SUGERIDO.
Nada é trocado sozinho: você revisa e decide.

Como usar (no terminal do VS Code, na pasta do projeto):
    python validar_links.py "Base_monitoramento_itens.xlsx"
Leva uns 5 a 10 minutos (são 3 sites por pneu, com pausas para não ser bloqueado).
Gera: Base_monitoramento_itens_conferida.xlsx
"""

import re
import sys
import time
from urllib.parse import quote, urlparse

import pandas as pd
from curl_cffi import requests

ROMA = "https://romapneus.com.br"
ACHEI = "https://www.acheipneus.com.br"
PNEUSTORE = "https://www.pneustore.com.br"
PAUSA = 1.5


# ============================================================
# MEDIDAS: "185/65R15", "185 65 R15", "18565r15", "185-65-r15" viram ("185","65","15")
# ============================================================

def medidas_no_texto(texto):
    t = str(texto).upper().replace(",", ".")
    achadas = set()
    for m in re.finditer(r"(\d{3})\s*[/\- ]?\s*(\d{2})\s*[\- ]?Z?R\s*[\- ]?(\d{2})(?:[.\-]5)?", t):
        meia = ".5" if re.match(r"[.\-]5", t[m.end() - 2:m.end()]) else ""
        achadas.add((m.group(1), m.group(2), m.group(3) + meia))
    for m in re.finditer(r"(\d{3})\s*[/\- ]?\s*(\d{2})\s*[\- ]?R\s*[\- ]?(\d{2})[.\-]5", t):
        achadas.add((m.group(1), m.group(2), m.group(3) + ".5"))
    for m in re.finditer(r"\b(\d{1,2})[.,](\d{2})\s*-\s*(\d{2})\b", t):     # 9.00-20 (pneu diagonal)
        achadas.add((m.group(1), m.group(2), m.group(3)))
    return achadas


def medida_do_item(produto):
    achadas = medidas_no_texto(produto)
    return sorted(achadas, key=lambda m: -len(m[2]))[0] if achadas else None


def mesma_medida(item, texto):
    if item is None:
        return False
    achadas = medidas_no_texto(texto)
    base = (item[0], item[1], item[2].replace(".5", ""))
    return item in achadas or base in {(a, b, c.replace(".5", "")) for a, b, c in achadas}


# Palavras da linha do pneu, para desempatar quando há vários da mesma medida
LINHAS = {"EDGE TOURING": ["EDGE", "TOURING"], "EDGE SPORT": ["EDGE", "SPORT"], "EDGE SUV": ["EDGE", "SUV"],
          "EAGLE SPORT": ["EAGLE", "SPORT"], "EAGLE TOURING": ["EAGLE", "TOURING"], "F1": ["F1"],
          "SUV": ["EFFICIENTGRIP", "SUV"], "PERF": ["EFFICIENTGRIP", "PERFORMANCE"], "MAXLIFE": ["MAXLIFE"],
          "TERRITORY": ["TERRITORY"], "FORTITUDE": ["FORTITUDE"], "MARATHON 3": ["MARATHON", "3"],
          "MARAT": ["MARATHON"], "MARA": ["MARATHON"], "G32": ["G32"], "REGIONAL": ["REGIONAL"],
          "AGS": ["AGS"], "KMAX": ["KMAX"], "PAPALEGUA": ["PAPA"]}


def palavras_da_linha(produto):
    p = produto.upper()
    for chave, palavras in LINHAS.items():
        if chave in p:
            return palavras
    return []


def nota(produto, nome):
    """Quão parecido é o nome encontrado com o pneu da planilha (0 a 100)."""
    if not mesma_medida(medida_do_item(produto), nome):
        return 0
    nome_u = str(nome).upper().replace("EFFIGRIP", "EFFICIENTGRIP").replace("PERF ", "PERFORMANCE ")
    palavras = palavras_da_linha(produto)
    acertos = sum(p in nome_u for p in palavras)
    indice = re.search(r"\b(\d{2,3})([A-Z])\b", produto.upper())
    bonus = 10 if indice and indice.group(0) in nome_u else 0
    return 50 + (40 * acertos // max(len(palavras), 1)) + bonus


# ============================================================
# CONSULTAS AOS SITES
# ============================================================

def baixar(url, **opcoes):
    for n in range(3):
        try:
            r = requests.get(url, impersonate="chrome", timeout=25, **opcoes)
            if r.status_code == 200:
                return r
        except Exception:
            pass
        time.sleep(2 * (n + 1))
    return None


def graphql(raiz, consulta):
    r = baixar(raiz + "/graphql", params={"query": consulta})
    try:
        return r.json()["data"]["products"]["items"] if r is not None else []
    except Exception:
        return []


def chave_da_url(url):
    return urlparse(str(url)).path.strip("/").split("/")[-1].replace(".html", "")


def roma(produto, codigo, url):
    """Devolve (conferência, link sugerido, código encontrado)."""
    campos = "sku name url_key stock_status"
    achado = None
    if codigo:
        itens = graphql(ROMA, '{products(filter:{sku:{eq:"%s"}}){items{%s}}}' % (codigo, campos))
        achado = itens[0] if itens else None
    if achado is None and url:
        itens = graphql(ROMA, '{products(filter:{url_key:{eq:"%s"}}){items{%s}}}' % (chave_da_url(url), campos))
        achado = itens[0] if itens else None
    if achado is None:
        itens = graphql(ROMA, '{products(search:"%s",pageSize:20){items{%s}}}' % (busca(produto), campos))
        itens = sorted(itens, key=lambda i: -nota(produto, i["name"]))
        if itens and nota(produto, itens[0]["name"]) >= 70:
            achado = itens[0]
    if achado is None:
        return ("NÃO ENCONTRADO no site da Roma (ainda não cadastrado?)", "", "")
    certo = f"{ROMA}/{achado['url_key']}"
    estoque = "com estoque" if achado.get("stock_status") == "IN_STOCK" else "sem estoque"
    if not mesma_medida(medida_do_item(produto), achado["name"]):
        return (f"ATENÇÃO: o produto encontrado é outra medida ({achado['name']})", certo, achado["sku"])
    if url and chave_da_url(url) == achado["url_key"]:
        return (f"ok · {estoque}", "", achado["sku"])
    motivo = "link estava vazio" if not url else "LINK DIFERENTE do cadastrado no site"
    return (f"{motivo} · {achado['name']} · {estoque}", certo, achado["sku"])


def busca(produto):
    medida = medida_do_item(produto)
    texto = f"{medida[0]}/{medida[1]}R{medida[2]}" if medida else produto
    return " ".join([texto] + palavras_da_linha(produto)).replace('"', "")


def achei(produto, url):
    """Devolve (conferência, link sugerido, SKU)."""
    campos = "sku name url_key"
    motivo = "link estava vazio" if not url else "link não abriu"
    if url:
        itens = graphql(ACHEI, '{products(filter:{url_key:{eq:"%s"}}){items{%s}}}' % (chave_da_url(url), campos))
        if itens:
            i = itens[0]
            n = nota(produto, i["name"])
            if n >= 70:
                return ("ok", "", i["sku"])
            if n >= 50:
                return (f"PARECIDO, conferir modelo: {i['name']}", "", i["sku"])
            motivo = f"ERRADO: o link abre outra medida ({i['name']})"
    itens = graphql(ACHEI, '{products(search:"%s",pageSize:20){items{%s}}}' % (busca(produto), campos))
    itens = sorted(itens, key=lambda i: -nota(produto, i["name"]))
    if itens and nota(produto, itens[0]["name"]) >= 50:
        i = itens[0]
        return (f"{motivo} · encontrado: {i['name']} (nota {nota(produto, i['name'])})",
                f"{ACHEI}/{i['url_key']}", i["sku"])
    return ("NÃO VENDE (não encontrado na Achei)" if not url else f"{motivo} e não achei na busca", "", "")


def titulo_pneustore(url):
    r = baixar(url)
    if r is None:
        return None
    m = re.search(r"<h1[^>]*>(.*?)</h1>", r.text, re.S)
    return re.sub(r"<[^>]+>|\s+", " ", m.group(1)).strip() if m else ""


def pneustore(produto, url):
    """Devolve (conferência, link sugerido)."""
    if not url:
        motivo = "link estava vazio"
    elif "/todos" in url or "?" in url:
        motivo = "link de busca"
    else:
        motivo = "link não abriu"
        titulo = titulo_pneustore(url)
        if titulo:
            n = nota(produto, titulo)
            if n >= 70:
                return ("ok", "")
            if n >= 50:
                return (f"PARECIDO, conferir modelo: {titulo}", "")
            motivo = f"ERRADO: o link abre outra medida ({titulo})"
    # Sem link, link de busca ou página que não abriu: procura na busca do site
    r = baixar(f"{PNEUSTORE}/todos?title={quote(busca(produto))}")
    if r is not None:
        links = list(dict.fromkeys(re.findall(r'href="(/pneu-[a-z0-9\-]+-\d{6,})"', r.text)))
        candidatos = sorted(links, key=lambda l: -nota(produto, l))
        if candidatos and nota(produto, candidatos[0]) >= 50:
            return (f"{motivo} · sugestão encontrada na busca (conferir)", PNEUSTORE + candidatos[0])
    return ("NÃO VENDE (não encontrado na Pneu Store)" if not url else f"{motivo} e não achei na busca", "")


# ============================================================
# EXECUÇÃO
# ============================================================

def texto(v):
    return "" if pd.isna(v) else str(v).strip()


def main(arquivo):
    itens = pd.read_excel(arquivo, sheet_name="Itens", dtype=str)
    saida = []
    for n, (_, r) in enumerate(itens.iterrows(), 1):
        produto = texto(r["Produto (base de vendas)"])
        codigo = texto(r["Código Roma"]).replace(".0", "")
        print(f"[{n}/{len(itens)}] {produto}")
        c_roma, s_roma, cod = roma(produto, codigo, texto(r["URL Roma"])); time.sleep(PAUSA)
        c_ps, s_ps = pneustore(produto, texto(r["URL Pneu Store"])); time.sleep(PAUSA)
        c_ac, s_ac, sku = achei(produto, texto(r["URL Achei Pneus"])); time.sleep(PAUSA)
        print(f"     Roma: {c_roma}\n     Pneu Store: {c_ps}\n     Achei: {c_ac}")
        saida.append({"Conferência Roma": c_roma, "Link sugerido Roma": s_roma, "Código Roma (site)": cod,
                      "Conferência Pneu Store": c_ps, "Link sugerido Pneu Store": s_ps,
                      "Conferência Achei": c_ac, "Link sugerido Achei": s_ac, "SKU Achei (descoberto)": sku})
    resultado = pd.concat([itens, pd.DataFrame(saida)], axis=1)
    destino = arquivo.replace(".xlsx", "_conferida.xlsx")
    with pd.ExcelWriter(destino, engine="openpyxl") as escritor:
        resultado.to_excel(escritor, sheet_name="Itens", index=False)
        planilha = escritor.sheets["Itens"]
        for coluna in planilha.columns:
            planilha.column_dimensions[coluna[0].column_letter].width = 28
    problemas = resultado[["Conferência Roma", "Conferência Pneu Store", "Conferência Achei"]].apply(
        lambda col: (~col.str.startswith("ok")).sum())
    print(f"\nPronto! Arquivo salvo: {destino}")
    print("Itens para revisar (não 'ok'):")
    for loja, qtd in problemas.items():
        print(f"   {loja.replace('Conferência ', ''):<11}| {qtd}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "Base_monitoramento_itens.xlsx")
