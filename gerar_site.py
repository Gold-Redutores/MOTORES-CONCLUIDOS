import argparse
import builtins
from pathlib import Path
import pandas as pd
import qrcode
import re
import html
import unicodedata
import json
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

# ALTERAR AQUI depois que o GitHub Pages estiver criado.
# Exemplo: "https://gold-redutores.github.io/DATABOOK-INPASA"
BASE_URL = "https://gold-redutores.github.io/MOTORES-CONCLUIDOS"

BASE_DIR = Path(__file__).resolve().parent
ARQUIVO_EXCEL = BASE_DIR / "dados_equipamentos.xlsx"
PASTA_EQUIPAMENTOS = BASE_DIR / "equipamentos"
PASTA_ATIVOS = builtins.next(
    (
        nome
        for nome in ("Ativos", "assets")
        if (BASE_DIR / nome).is_dir()
    ),
    "assets",
)
PASTA_QRCODES = BASE_DIR / PASTA_ATIVOS / "qrcodes"
PASTA_LOGOS_CLIENTES = Path(PASTA_ATIVOS) / "img" / "clientes"
LOGO_PADRAO_CLIENTE = f"{PASTA_ATIVOS}/img/logo_gold.png"
LOGOS_CLIENTES = {
    "adm do brasil ltda": "adm.png",
    "bunge alimentos s/a": "bunge.svg",
    "cargill agricola s.a": "cargill.svg",
    "cofco international brasil s.a.": "cofco.png",
    "oleoplan rondonia industria de biocombustivel ltda": "oleoplan.png",
    "rumo malha norte s.a": "rumo.svg",
}
EXTENSOES_LOGO = {".png", ".jpg", ".jpeg", ".webp", ".gif"}
COMMONS_API = "https://commons.wikimedia.org/w/api.php"
USER_AGENT = "DataBookDigitalClientLogoUpdater/1.0"
TIMEOUT_BUSCA_LOGO = 20
TERMOS_LEGAIS = {
    "agricola", "alimentos", "brasil", "com", "da", "das", "de", "do",
    "dos", "e", "industria", "ltda", "s", "a", "sa",
}

def limpar_texto(valor):
    if pd.isna(valor):
        return ""
    return builtins.str(valor).strip()


def primeiro_valor_nao_vazio(*valores):
    for valor in valores:
        texto = limpar_texto(valor)
        if texto:
            return texto
    return ""


def tag_para_pasta(tag):
    tag = limpar_texto(tag).upper()
    tag = tag.replace(" ", "-")
    tag = re.sub(r"[^A-Z0-9_-]", "-", tag)
    tag = re.sub(r"-+", "-", tag)
    return tag.strip("-")


def link_relativo(caminho):
    caminho = limpar_texto(caminho)
    if caminho == "":
        return ""
    if caminho.startswith("http://") or caminho.startswith("https://"):
        return caminho
    if caminho.startswith("../../"):
        return caminho

    caminho = caminho.replace("\\", "/").lstrip("/")
    partes = Path(caminho).parts
    if partes and partes[0].casefold() in {"ativos", "assets"}:
        caminho = Path(PASTA_ATIVOS, *partes[1:]).as_posix()

    arquivo = BASE_DIR / caminho
    if not arquivo.is_file() and not Path(caminho).suffix:
        arquivo_pdf = arquivo.with_suffix(".pdf")
        if arquivo_pdf.is_file():
            caminho += ".pdf"
        else:
            return ""
    elif not arquivo.is_file():
        return ""

    return "../../" + caminho


def criar_botao(titulo, caminho, classe="botao"):
    caminho = limpar_texto(caminho)
    if caminho == "":
        return f'<span class="botao inativo">{html.escape(titulo)} indisponível</span>'
    link = link_relativo(caminho)
    if link == "":
        return f'<span class="botao inativo">{html.escape(titulo)} indisponível</span>'
    return f'<a class="{classe}" href="{html.escape(link)}" target="_blank">{html.escape(titulo)}</a>'


def gerar_html_equipamento(dados):
    tag = limpar_texto(dados.get("TAG", ""))
    tag_pasta = tag_para_pasta(tag)

    foto = link_relativo(dados.get("Foto", ""))
    if foto:
        bloco_foto = f'<img src="{html.escape(foto)}" alt="Foto do equipamento {html.escape(tag)}" class="foto-equipamento">'
    else:
        bloco_foto = '<div class="obs">Foto não cadastrada na planilha.</div>'

    campos = [
        ("TAG", "TAG"),
        ("Fabricante", "Fabricante"),
        #("Tipo do equipamento", "Tipo"),
        ("Modelo", "Modelo"),
        ("Potência", "Potencia"),
        #("Relação de redução", "Relacao"),
        ("RPM entrada/saída", "RPM"),
        # ("Tipo de montagem", "Montagem"),
        #("Lubrificante", "Lubrificante"),
        #("Quantidade de óleo", "Quantidade_Oleo"),
        #("Local / Área", "Local"),
        ("Observações", "Observacoes"),
    ]

    linhas = ""
    for titulo, coluna in campos:
        valor = limpar_texto(dados.get(coluna, ""))
        linhas += f"""
        <tr>
          <th>{html.escape(titulo)}</th>
          <td>{html.escape(valor)}</td>
        </tr>"""

    botao_lista = criar_botao("Lista de Peças", dados.get("Lista_Pecas", ""))
    botao_manual = criar_botao("Manual", dados.get("Manual", ""))

    caminho_relatorio = primeiro_valor_nao_vazio(
        dados.get("RELATÓRIO FINAL", ""),
        dados.get("RELATORIO FINAL", ""),
        dados.get("Relatorio Final", ""),
        dados.get("Certificado", ""),
        dados.get("CERTIFICADA", ""),
    )
    botao_relatorio = criar_botao("RELATORIO FINAL", caminho_relatorio)
    botao_fotos = criar_botao("Fotos adicionais", dados.get("Fotos_Adicionais", ""))

    link_publicado = BASE_URL.rstrip("/") + "/equipamentos/" + tag_pasta + "/"

    logo_gold = "https://i.ibb.co/V07bLNT1/GOLD-MAQUINAS.png"

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Mini Data Book - {html.escape(tag)}</title>
  <link rel="stylesheet" href="../../{PASTA_ATIVOS}/css/style.css">
</head>
<body class="pagina-equipamento">

  <header class="topo">
    <div class="header-row">
      <div class="logos">
        <img class="logo-gold" src="{html.escape(logo_gold)}" alt="Gold Redutores">
      </div>
      <div class="header-text">
        <div class="tipo">{html.escape(limpar_texto(dados.get("Tipo", "Equipamento")))}</div>
        <div class="tag">{html.escape(tag)}</div>
      </div>
    </div>
  </header>

  <main class="container">
    <section class="photo-card">
      {bloco_foto}
    </section>

    <section class="card">
      <h2>Dados principais</h2>
      <table class="tabela-dados">
        {linhas}
      </table>
    </section>

    <section class="card">
      <h2>Acessos rápidos</h2>
      <div class="botoes">
        {botao_relatorio}
        {botao_lista}
        {botao_manual}
        {botao_fotos}
        <a href="../../index.html" class="botao secundario">Voltar para lista geral</a>
      </div>
    </section>

    <section class="card">
      <h2>Link desta página</h2>
      <div class="obs">{html.escape(link_publicado)}</div>
    </section>
  </main>

  <footer>
    Gold Máquinas - Mini Data Book Digital | Página individual por equipamento
  </footer>

</body>
</html>
"""


def gerar_index_principal(df):
    cards = ""

    for _, row in df.iterrows():
        tag = limpar_texto(row.get("TAG", ""))
        if not tag:
            continue

        tag_pasta = tag_para_pasta(tag)
        tipo = limpar_texto(row.get("Tipo", "Equipamento"))
        modelo = limpar_texto(row.get("Modelo", ""))
        local = limpar_texto(row.get("Local", ""))
        busca = (tag + " " + tipo + " " + modelo + " " + local).lower()

        cards += f"""
      <a class="equipamento-card" href="equipamentos/{html.escape(tag_pasta)}/index.html" data-busca="{html.escape(busca)}">
        <strong>{html.escape(tag)}</strong>
        <span>{html.escape(tipo)}</span>
        <span>{html.escape(modelo)}</span>
        <small>{html.escape(local)}</small>
      </a>
"""

    logo_gold = "https://i.ibb.co/zWgjxdtj/logo.png"
    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Data Book Digital - Equipamentos</title>
  <link rel="stylesheet" href="{PASTA_ATIVOS}/css/style.css">
  <style>
    /* Estilos Integrados do Pure CSS3 Slider - Copiado do Relatório */
    .header-slider-wrapper {{ position: sticky; top: 0; z-index: 1000; }}
    #slider {{ width: 90%; max-width: 1080px; height: 70px; position: relative; overflow: hidden; margin: 8px auto 15px auto; border: 1px solid #333; background: #fff; }}
    .slides {{ display: flex; width: 400%; height: 100%; position: relative; -webkit-animation: slide 5s infinite; -moz-animation: slide 5s infinite; animation: slide 5s infinite; }}
    .slider {{ flex: 0 0 25%; height: 100%; position: relative; z-index: 1; overflow: hidden; box-sizing: border-box; border-right: 2px solid #333; }}
    .image {{ width: 100%; height: 100%; position: relative; display: flex; align-items: center; justify-content: center; }}
    .image img {{ width: 70%; height: 80%; object-fit: contain; display: block; }}
    .switch {{ width: 120px; height: 10px; position: absolute; bottom: 18px; z-index: 99; left: 20px; }}
    .switch>ul {{ list-style: none; overflow: hidden; padding: 0; margin: 0; }}
    .switch>ul>li {{ width: 6px; height: 6px; border-radius: 50%; background: #333; float: left; margin-right: 9px; cursor: pointer; }}
    .on {{ width: 100%; height: 100%; border-radius: 50%; background: rgba(23, 117, 177, 1); position: relative; -webkit-animation: on 5s infinite; -moz-animation: on 5s infinite; animation: on 5s infinite; }}
    @-webkit-keyframes slide {{ 0%, 22% {{ transform: translateX(0); }} 25%, 47% {{ transform: translateX(-25%); }} 50%, 72% {{ transform: translateX(-50%); }} 75%, 100% {{ transform: translateX(-75%); }} }}
    @keyframes slide {{ 0%, 22% {{ transform: translateX(0); }} 25%, 47% {{ transform: translateX(-25%); }} 50%, 72% {{ transform: translateX(-50%); }} 75%, 100% {{ transform: translateX(-75%); }} }}
    @-webkit-keyframes on {{ 0%, 22% {{ margin-left: 0; }} 25%, 47% {{ margin-left: 15px; }} 50%, 100% {{ margin-left: 30px; }} }}
    @keyframes on {{ 0%, 22% {{ margin-left: 0; }} 25%, 47% {{ margin-left: 15px; }} 50%, 100% {{ margin-left: 30px; }} }}
  </style>
</head>
<body>

  <div class="header-slider-wrapper">
    <header class="topo">
      <div class="header-row">
        <div class="logos">
          <img class="logo-gold" src="{html.escape(logo_gold)}" alt="Gold Redutores">
        </div>
        <div class="header-text">
          <div class="tipo">Data Book Digital</div>
          <div class="tag">Equipamentos</div>
        </div>
      </div>
    </header>

    <!-- Carrossel de Imagens -->
    <div id="slider">
        <div class="slides">
            <div class="slider">
                <div class="image">
                    <img src="https://i.ibb.co/V07bLNT1/GOLD-MAQUINAS.png" alt="Logo Gold Redutores">
                </div>
            </div>
            <div class="slider">
                <div class="image">
                    <img src="https://i.ibb.co/r9DRsZq/GOLD-REDUTORES.png" alt="GOLD REDUTORES" border="0">
                </div>
            </div>
            <div class="slider">
                <div class="image">
                    <img src="https://i.ibb.co/3mVp95wL/GOLD-AUTOMA-O.png" alt="GOLD AUTOMAÇÃO" border="0">
                </div>
            </div>
            <div class="slider" aria-hidden="true">
              <div class="image">
                <img src="https://i.ibb.co/V07bLNT1/GOLD-MAQUINAS.png" alt="">
              </div>
            </div>
        </div>
        <div class="switch">
            <ul>
                <li><div class="on"></div></li>
                <li></li>
                <li></li>
            </ul>
        </div>
    </div>
  </div>

  <main class="container">
    <section class="card">
      <h2>Equipamentos cadastrados</h2>
      <p>Digite a TAG, modelo, tipo ou local para localizar o equipamento.</p>
      <input class="busca" id="busca" type="text" placeholder="Exemplo: ME-631011, motoredutor, WEG...">
    </section>

    <section class="grid-equipamentos" id="listaEquipamentos">
      {cards}
    </section>
  </main>

  <footer>Gold Redutores - Data Book Digital</footer>

  <script>
    const campoBusca = document.getElementById("busca");
    const cards = document.querySelectorAll(".equipamento-card");

    campoBusca.addEventListener("input", function() {{
      const termo = campoBusca.value.toLowerCase().trim();
      cards.forEach(card => {{
        const texto = card.getAttribute("data-busca");
        card.style.display = texto.includes(termo) ? "block" : "none";
      }});
    }});
  </script>

</body>
</html>
"""


def coluna_cliente(df):
    return builtins.next(
        (coluna for coluna in df.columns if builtins.str(coluna).strip().casefold() == "cliente"),
        None,
    )


def clientes_e_pastas(df, coluna):
    clientes_por_chave = {}
    for valor in df[coluna]:
        cliente = limpar_texto(valor)
        if cliente:
            clientes_por_chave.setdefault(cliente.casefold(), cliente)

    pastas_usadas = builtins.set()
    resultado = {}
    for chave, cliente in builtins.sorted(
        clientes_por_chave.items(), key=lambda item: (item[1].casefold(), item[1])
    ):
        base = unicodedata.normalize("NFKD", cliente)
        base = base.encode("ascii", "ignore").decode("ascii").upper()
        base = re.sub(r"[^A-Z0-9_-]+", "-", base)
        base = re.sub(r"-+", "-", base).strip("-") or "CLIENTE"

        pasta = base
        sufixo = 2
        while pasta in pastas_usadas:
            pasta = f"{base}-{sufixo}"
            sufixo += 1
        pastas_usadas.add(pasta)
        resultado[chave] = (cliente, pasta)
    return resultado


def caminho_logo_cliente(chave, pasta):
    arquivo_logo = LOGOS_CLIENTES.get(chave)
    if arquivo_logo:
        caminho = PASTA_LOGOS_CLIENTES / arquivo_logo
        caminho_completo = BASE_DIR / caminho
        if caminho_completo.is_file():
            return caminho.as_posix()
        builtins.print(
            f"Logo cadastrado não encontrado: {caminho_completo}; "
            f"será usado o logo padrão."
        )
        return LOGO_PADRAO_CLIENTE

    diretorio_logos = BASE_DIR / PASTA_LOGOS_CLIENTES
    candidatos = builtins.sorted(
        caminho
        for caminho in diretorio_logos.glob(f"{pasta}-logo.*")
        if caminho.suffix.casefold() in EXTENSOES_LOGO
    )
    if candidatos:
        return candidatos[0].relative_to(BASE_DIR).as_posix()
    return LOGO_PADRAO_CLIENTE


def normalizar_nome_logo(valor):
    texto = unicodedata.normalize("NFKD", builtins.str(valor))
    texto = "".join(caractere for caractere in texto if not unicodedata.combining(caractere))
    return re.sub(r"[^a-z0-9]+", " ", texto.casefold()).strip()


def buscar_logo_na_internet(cliente):
    termos = [
        termo
        for termo in normalizar_nome_logo(cliente).split()
        if termo not in TERMOS_LEGAIS
    ]
    if not termos:
        termos = normalizar_nome_logo(cliente).split()
    if not termos:
        return None

    parametros = urlencode(
        {
            "action": "query",
            "generator": "search",
            "gsrsearch": " ".join(termos + ["logo"]),
            "gsrnamespace": 6,
            "gsrlimit": 15,
            "prop": "imageinfo",
            "iiprop": "url",
            "iiurlwidth": 600,
            "format": "json",
        }
    )
    requisicao = Request(
        f"{COMMONS_API}?{parametros}",
        headers={"User-Agent": USER_AGENT},
    )
    with urlopen(requisicao, timeout=TIMEOUT_BUSCA_LOGO) as resposta:
        dados = json.load(resposta)

    melhores = []
    termos_marca = builtins.set(termos)
    for pagina in dados.get("query", {}).get("pages", {}).values():
        titulo = normalizar_nome_logo(pagina.get("title", ""))
        palavras = builtins.set(titulo.split())
        if "logo" not in palavras or termos[0] not in palavras:
            continue

        correspondencias = builtins.len(termos_marca & palavras)
        if builtins.len(termos) > 1 and correspondencias < 2:
            continue

        imagens = pagina.get("imageinfo", [])
        if imagens and imagens[0].get("thumburl"):
            melhores.append((correspondencias, titulo, imagens[0]["thumburl"]))

    if not melhores:
        return None
    melhores.sort(key=lambda item: (-item[0], item[1]))
    return melhores[0][2]


def extensao_segura_imagem(conteudo):
    if conteudo.startswith(b"\x89PNG\r\n\x1a\n"):
        return ".png"
    if conteudo.startswith(b"\xff\xd8\xff"):
        return ".jpg"
    if conteudo.startswith((b"GIF87a", b"GIF89a")):
        return ".gif"
    if builtins.len(conteudo) >= 12 and conteudo[:4] == b"RIFF" and conteudo[8:12] == b"WEBP":
        return ".webp"
    return ""


def baixar_logo_cliente(url, cliente, pasta):
    requisicao = Request(url, headers={"User-Agent": USER_AGENT})
    with urlopen(requisicao, timeout=TIMEOUT_BUSCA_LOGO) as resposta:
        conteudo = resposta.read(5_000_001)
    if builtins.len(conteudo) > 5_000_000:
        raise builtins.ValueError("A imagem encontrada ultrapassa o limite de 5 MB.")

    extensao = extensao_segura_imagem(conteudo)
    if not extensao:
        raise builtins.ValueError("O resultado da busca não é um logo em formato de imagem compatível.")

    diretorio = BASE_DIR / PASTA_LOGOS_CLIENTES
    diretorio.mkdir(parents=True, exist_ok=True)
    destino = diretorio / f"{pasta}-logo{extensao}"
    destino.write_bytes(conteudo)
    builtins.print(f"Logo encontrado para {cliente}: {destino.relative_to(BASE_DIR)}")


def procurar_logos_novos(clientes):
    for chave, (cliente, pasta) in clientes.items():
        if caminho_logo_cliente(chave, pasta) != LOGO_PADRAO_CLIENTE:
            continue
        if chave in LOGOS_CLIENTES:
            continue
        try:
            url_logo = buscar_logo_na_internet(cliente)
            if not url_logo:
                builtins.print(f"Logo confiável não encontrado para {cliente}; será usado o logo Gold.")
                continue
            baixar_logo_cliente(url_logo, cliente, pasta)
        except (
            HTTPError,
            URLError,
            builtins.TimeoutError,
            json.JSONDecodeError,
            builtins.ValueError,
        ) as erro:
            builtins.print(f"Não foi possível obter o logo de {cliente}: {erro}. Será usado o logo Gold.")


def cartoes_equipamentos(df, links_relativos=False):
    cards = ""
    for _, row in df.iterrows():
        tag = limpar_texto(row.get("TAG", ""))
        if not tag:
            continue

        tag_pasta = tag_para_pasta(tag)
        tipo = limpar_texto(row.get("Tipo", "Equipamento"))
        modelo = limpar_texto(row.get("Modelo", ""))
        local = limpar_texto(row.get("Local", ""))
        busca = (tag + " " + tipo + " " + modelo + " " + local).lower()
        if links_relativos:
            href = f"../../equipamentos/{tag_pasta}/index.html"
        else:
            href = f"equipamentos/{tag_pasta}/index.html"

        cards += f"""
      <a class="equipamento-card" href="{html.escape(href, quote=True)}" data-busca="{html.escape(busca, quote=True)}">
        <strong>{html.escape(tag)}</strong>
        <span>{html.escape(tipo)}</span>
        <span>{html.escape(modelo)}</span>
        <small>{html.escape(local)}</small>
      </a>
"""
    return cards


def cabecalho_indice(titulo):
    return f"""
  <div class="header-slider-wrapper">
    <header class="topo">
      <div class="header-row">
        <div class="logos">
          <img class="logo-gold" src="https://i.ibb.co/zWgjxdtj/logo.png" alt="Gold Redutores">
        </div>
        <div class="header-text">
          <div class="tipo">Data Book Digital</div>
          <div class="tag">{html.escape(titulo)}</div>
        </div>
      </div>
    </header>
    <div id="slider">
      <div class="slides">
        <div class="slider"><div class="image"><img src="https://i.ibb.co/V07bLNT1/GOLD-MAQUINAS.png" alt="Logo Gold Redutores"></div></div>
        <div class="slider"><div class="image"><img src="https://i.ibb.co/r9DRsZq/GOLD-REDUTORES.png" alt="Gold Redutores"></div></div>
        <div class="slider"><div class="image"><img src="https://i.ibb.co/3mVp95wL/GOLD-AUTOMA-O.png" alt="Gold Automação"></div></div>
        <div class="slider" aria-hidden="true"><div class="image"><img src="https://i.ibb.co/V07bLNT1/GOLD-MAQUINAS.png" alt=""></div></div>
      </div>
      <div class="switch"><ul><li><div class="on"></div></li><li></li><li></li></ul></div>
    </div>
  </div>"""


def gerar_pagina_indice(titulo, titulo_secao, descricao, cards, caminho_css, botao_voltar=""):
    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Data Book Digital - {html.escape(titulo)}</title>
  <link rel="stylesheet" href="{html.escape(caminho_css, quote=True)}">
  <style>
    .header-slider-wrapper {{ position: sticky; top: 0; z-index: 1000; }}
    #slider {{ width: 90%; max-width: 1080px; height: 70px; position: relative; overflow: hidden; margin: 8px auto 15px; border: 1px solid #333; background: #fff; }}
    .slides {{ display: flex; width: 400%; height: 100%; position: relative; animation: slide 5s infinite; }}
    .slider {{ flex: 0 0 25%; height: 100%; overflow: hidden; box-sizing: border-box; border-right: 2px solid #333; }}
    .image {{ width: 100%; height: 100%; display: flex; align-items: center; justify-content: center; }}
    .image img {{ width: 70%; height: 80%; object-fit: contain; display: block; }}
    .switch {{ width: 120px; height: 10px; position: absolute; bottom: 18px; left: 20px; }}
    .switch ul {{ list-style: none; overflow: hidden; padding: 0; margin: 0; }}
    .switch li {{ width: 6px; height: 6px; border-radius: 50%; background: #333; float: left; margin-right: 9px; }}
    .on {{ width: 100%; height: 100%; border-radius: 50%; background: #1775b1; animation: on 5s infinite; }}
    .logo-cliente {{ display: block; width: 100%; height: 82px; margin: 0 auto 16px; padding: 8px; object-fit: contain; background: #fff; border-radius: 8px; }}
    @keyframes slide {{ 0%, 22% {{ transform: translateX(0); }} 25%, 47% {{ transform: translateX(-25%); }} 50%, 72% {{ transform: translateX(-50%); }} 75%, 100% {{ transform: translateX(-75%); }} }}
    @keyframes on {{ 0%, 22% {{ margin-left: 0; }} 25%, 47% {{ margin-left: 15px; }} 50%, 100% {{ margin-left: 30px; }} }}
  </style>
</head>
<body>
{cabecalho_indice(titulo)}
  <main class="container">
    <section class="card">
      <h2>{html.escape(titulo_secao)}</h2>
      <p>{html.escape(descricao)}</p>
      {"<div class='botoes'>" + botao_voltar + "</div>" if botao_voltar else ""}
      {"<input class='busca' id='busca' type='text' placeholder='Exemplo: OS 6251, motoredutor, WEG...'>" if titulo_secao != "Clientes cadastrados" else ""}
    </section>
    <section class="grid-equipamentos" id="listaEquipamentos">
      {cards}
    </section>
  </main>
  <footer>Gold Soluções e Serviços - Data Book Digital</footer>
  {"<script>const campoBusca = document.getElementById('busca'); const cards = document.querySelectorAll('.equipamento-card'); campoBusca.addEventListener('input', function() { const termo = campoBusca.value.toLowerCase().trim(); cards.forEach(card => { card.style.display = card.getAttribute('data-busca').includes(termo) ? 'block' : 'none'; }); });</script>" if titulo_secao != "Clientes cadastrados" else ""}
</body>
</html>
"""


def gerar_paginas_clientes(df):
    coluna = coluna_cliente(df)
    if coluna is None:
        raise builtins.ValueError("A planilha precisa ter uma coluna chamada Cliente.")

    clientes = clientes_e_pastas(df, coluna)
    procurar_logos_novos(clientes)
    paginas_individuais_ausentes = 0
    for _, row in df.iterrows():
        if not limpar_texto(row.get(coluna, "")):
            continue
        tag = limpar_texto(row.get("TAG", ""))
        if not tag:
            continue
        pasta_equipamento = PASTA_EQUIPAMENTOS / tag_para_pasta(tag)
        pagina_equipamento = pasta_equipamento / "index.html"
        if not pagina_equipamento.exists():
            pasta_equipamento.mkdir(parents=True, exist_ok=True)
            pagina_equipamento.write_text(
                gerar_html_equipamento(row),
                encoding="utf-8",
            )
            paginas_individuais_ausentes += 1

    cards_clientes = ""
    for chave, (cliente, pasta) in clientes.items():
        quantidade = builtins.sum(
            limpar_texto(valor).casefold() == chave for valor in df[coluna]
        )
        caminho_logo = caminho_logo_cliente(chave, pasta)
        logo_html = (
            f'<img class="logo-cliente" src="{html.escape(caminho_logo, quote=True)}" '
            f'alt="Logo {html.escape(cliente, quote=True)}" loading="lazy">'
        )
        cards_clientes += f"""
      <a class="equipamento-card" href="clientes/{html.escape(pasta, quote=True)}/index.html">
        {logo_html}
        <strong>{html.escape(cliente)}</strong>
        <span>{quantidade} equipamento(s)</span>
      </a>
"""

    index_clientes = gerar_pagina_indice(
        "Clientes",
        "Clientes cadastrados",
        "Selecione um cliente para visualizar somente os equipamentos vinculados a ele.",
        cards_clientes,
        f"{PASTA_ATIVOS}/css/style.css",
    )
    (BASE_DIR / "index.html").write_text(index_clientes, encoding="utf-8")

    for chave, (cliente, pasta) in clientes.items():
        equipamentos_cliente = df[
            df[coluna].map(limpar_texto).str.casefold() == chave
        ]
        cards = cartoes_equipamentos(equipamentos_cliente, links_relativos=True)
        voltar = '<a class="botao secundario" href="../../index.html">Voltar para lista geral</a>'
        pagina = gerar_pagina_indice(
            cliente,
            "Equipamentos cadastrados",
            f"Equipamentos cadastrados para {cliente}.",
            cards,
            f"../../{PASTA_ATIVOS}/css/style.css",
            voltar,
        )
        destino = BASE_DIR / "clientes" / pasta
        destino.mkdir(parents=True, exist_ok=True)
        (destino / "index.html").write_text(pagina, encoding="utf-8")

    return builtins.len(clientes)


def main(somente_clientes=False):
    arquivo_excel = ARQUIVO_EXCEL
    if not arquivo_excel.exists():
        fallback = BASE_DIR / "dados_equipamentos.xlsx"
        if fallback.exists():
            arquivo_excel = fallback
        else:
            raise builtins.FileNotFoundError(f"Arquivo não encontrado: {arquivo_excel}")

    df = pd.read_excel(arquivo_excel)

    if "TAG" not in df.columns:
        raise builtins.ValueError("A planilha precisa ter uma coluna chamada TAG.")

    if somente_clientes:
        total_clientes = gerar_paginas_clientes(df)
        builtins.print(f"Página inicial criada: index.html ({total_clientes} clientes).")
        return

    PASTA_EQUIPAMENTOS.mkdir(exist_ok=True)
    PASTA_QRCODES.mkdir(parents=True, exist_ok=True)

    total = 0

    for _, row in df.iterrows():
        tag = limpar_texto(row.get("TAG", ""))
        if not tag:
            continue

        tag_pasta = tag_para_pasta(tag)
        pasta = PASTA_EQUIPAMENTOS / tag_pasta
        pasta.mkdir(parents=True, exist_ok=True)

        html_equip = gerar_html_equipamento(row)
        (pasta / "index.html").write_text(html_equip, encoding="utf-8")

        link_equipamento = BASE_URL.rstrip("/") + f"/equipamentos/{tag_pasta}/"

        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=10,
            border=4,
        )
        qr.add_data(link_equipamento)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        img.save(PASTA_QRCODES / f"QR-{tag_pasta}.png")

        total += 1

    total_clientes = gerar_paginas_clientes(df)

    builtins.print(f"Concluído: {total} páginas criadas.")
    builtins.print(f"QR Codes salvos em {PASTA_QRCODES.relative_to(BASE_DIR)}")
    builtins.print(f"Página inicial criada: index.html ({total_clientes} clientes).")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Gera as páginas do Data Book Digital.")
    parser.add_argument(
        "--clientes-only",
        action="store_true",
        help="Atualiza index.html e as páginas dos clientes sem regenerar todos os equipamentos.",
    )
    main(somente_clientes=parser.parse_args().clientes_only)
