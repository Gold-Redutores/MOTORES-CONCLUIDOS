
# Data Book Digital - Excel + GitHub Pages

Este pacote gera páginas HTML individuais para equipamentos a partir de uma planilha Excel.

## Como usar

1. Preencha `dados_equipamentos.xlsx`.
2. Coloque fotos em `assets/img/`.
3. Coloque PDFs em `assets/pdf/`.
4. Abra `gerar_site.py`.
5. Troque a linha `BASE_URL` pelo link real do GitHub Pages.
6. Rode:

```bash
pip install pandas openpyxl qrcode pillow
python gerar_site.py
```

## Estrutura esperada

```text
DATABOOK_INPASA_EXCEL_GITHUB/
├── dados_equipamentos.xlsx
├── gerar_site.py
├── index.html
├── assets/
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
https://gold-redutores.github.io/DATABOOK-INPASA/equipamentos/ME-631011/
```
