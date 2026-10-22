
# Data Book Digital - Excel + GitHub Pages

Este pacote gera páginas HTML individuais para equipamentos a partir de uma planilha Excel.

## Como usar

1. Preencha `dados_equipamentos.xlsx`.
2. Coloque fotos em `assets/img/` ou `Ativos/img/`.
3. Coloque PDFs em `assets/pdf/` ou `Ativos/pdf/`.
4. Abra `gerar_site.py`.
5. Confira se `BASE_URL` corresponde ao endereço publicado no GitHub Pages.
6. Rode:

```bash
pip install pandas openpyxl qrcode pillow
python gerar_site.py
```

Para atualizar somente a página inicial de clientes e suas páginas, sem regenerar todos os equipamentos:

```bash
python gerar_site.py --clientes-only
```

Esse modo atualiza `index.html` e uma página em `clientes/` para cada valor preenchido na coluna `Cliente`. Cada página lista apenas os equipamentos daquele cliente. Se faltar a página individual de algum equipamento vinculado, ela será criada sem reescrever as páginas já existentes.

Na página de clientes, linhas sem valor em `Cliente` são ignoradas. Os logos são armazenados em `Ativos/img/clientes/`; para clientes sem logo cadastrado no gerador, é usado `Ativos/img/logo_gold.png`. Os logos ADM, Bunge, Oleoplan e Rumo vieram dos respectivos sites oficiais; os logos Cargill e COFCO vieram do Wikimedia Commons.

O mesmo `gerar_site.py` pesquisa automaticamente logos para clientes novos no Wikimedia Commons, salva os arquivos encontrados em `Ativos/img/clientes/` e atualiza `index.html` e as páginas de cada cliente. Linhas com `Cliente` vazio continuam ignoradas. Se a pesquisa não encontrar um logo correspondente ou estiver indisponível, será usado o logo Gold; logotipos já cadastrados são preservados.

```bash
python gerar_site.py
```

O gerador usa `Ativos/` quando essa pasta existe; caso contrário, usa `assets/`. O `index.html` na raiz é a página inicial de clientes para uso no GitHub Pages. O `gerar_site.py` padrão regenera também as páginas individuais dos equipamentos e QR Codes; use `--clientes-only` se atualizou apenas a coluna de clientes/logos. Referências a PDFs na planilha podem ser informadas sem a extensão `.pdf`; o gerador a completa quando o arquivo correspondente existe. Arquivos locais inexistentes aparecem como indisponíveis, em vez de gerar links quebrados.

Depois de gerar, publique no repositório GitHub, na branch e pasta configuradas no GitHub Pages, o `index.html` e as pastas `clientes/`, `equipamentos/` e `assets/` ou `Ativos/`. Atualizar os arquivos apenas no computador não altera o site publicado. Confirme também que a origem do Pages está configurada para a branch `main` e a pasta raiz (`/`).

## Estrutura esperada

```text
MOTORES-CONCLUIDOS/
├── dados_equipamentos.xlsx
├── gerar_site.py
├── index.html
├── assets/ (ou Ativos/)
│   ├── css/style.css
│   ├── img/
│   ├── pdf/
│   └── qrcodes/
└── equipamentos/
    └── ME-631011/
        └── index.html
```

## QR Code

Cada QR Code aponta para:

```text
https://SEU_USUARIO.github.io/NOME_REPOSITORIO/equipamentos/TAG/
```

Exemplo:

```text
https://gold-redutores.github.io/MOTORES-CONCLUIDOS/equipamentos/OS-4938-910384/
```
