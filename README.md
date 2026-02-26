# Preencher Fotos em Excel

Ferramenta em Python para preencher automaticamente imagens em planilhas Excel com base no nome do arquivo informado na propria planilha.

O projeto possui:

- Motor de processamento (detectar colunas `Nome`/`Foto`, localizar imagens e inserir no Excel)
- Interface grafica simples (2 campos para selecionar pastas)
- Empacotamento em `.exe` com `PyInstaller`

## Visao Geral

Fluxo do sistema:

1. Ler uma pasta de planilhas (`PLANILHAS`) e uma pasta de imagens (`FOTOS`)
2. Indexar as imagens por nome de arquivo (recursivo)
3. Abrir cada planilha no Excel (COM)
4. Detectar automaticamente colunas de origem/destino por cabecalho
5. Inserir as imagens nas celulas correspondentes
6. Salvar a planilha e emitir resumo no log

## Requisitos

- Windows
- Microsoft Excel instalado (obrigatorio)
- Python 3.x (apenas para desenvolvimento/build)

Para usar somente o executavel, nao precisa de Python.

## Estrutura do Projeto

```text
ADRIEL_DEMANDA/
|- Inserir foto.py              # Ponto de entrada (CLI/UI)
|- preencher_fotos_core.py      # Motor principal de processamento
|- preencher_fotos_ui.py        # Interface grafica (Tkinter)
|- build_exe.bat                # Script para gerar o .exe
|- PreencherFotosExcel.spec     # Configuracao do PyInstaller
|- PARA_RODAR/                  # Pasta de uso final
|  |- PreencherFotosExcel.exe
|  |- Executar_Preenchimento.bat
|  |- README_USO.txt
|  |- PLANILHAS/
|  `- FOTOS/
`- docs/
   |- ARQUITETURA.md
   `- OPERACAO.md
```

## Como Usar (Usuario Final)

### Opcao 1: Executavel + pastas padrao

Na pasta `PARA_RODAR`, use:

- `PLANILHAS/` para as planilhas (`.xlsx`, `.xlsm`, `.xls`)
- `FOTOS/` para as imagens (`.png`, `.jpg`, `.jpeg`, `.bmp`, `.gif`, `.webp`)

Pode usar subpastas em ambos os casos.

Depois execute:

- `PARA_RODAR/PreencherFotosExcel.exe`
ou
- `PARA_RODAR/Executar_Preenchimento.bat`

### Opcao 2: Interface com selecao manual

Ao abrir o `.exe`, a interface permite informar:

- Pasta das planilhas
- Pasta das fotos

Com botoes de selecao, alem do botao para usar as pastas padrao.

## Como Rodar em Desenvolvimento

### Interface grafica (padrao)

```powershell
python "Inserir foto.py"
```

### Modo console (CLI)

```powershell
python "Inserir foto.py" --cli
```

Observacao:
- O modo CLI alterna a entrada para console, mas ainda pode usar as pastas padrao ou abrir dialogs de selecao.

## Como o Sistema Detecta as Colunas

O motor procura cabecalhos nas abas com palavras como:

- Origem (nome do arquivo): `nome`, `arquivo`
- Destino (onde entra imagem): `foto`, `imagem`, `link`

Quando encontra indices (`Nome 01`, `Foto 01`, etc.), ele prioriza o pareamento por indice.
Se nao houver indice, faz pareamento por ordem das colunas.

## Build do Executavel

### Instalar dependencias (uma vez)

```powershell
python -m pip install pyinstaller pywin32
```

### Gerar o `.exe`

```powershell
.\build_exe.bat
```

Saida esperada:

- `dist/PreencherFotosExcel.exe`

Depois copie/substitua em:

- `PARA_RODAR/PreencherFotosExcel.exe`

## Arquivos Principais (Repositorio)

- `Inserir foto.py`: entrada curta que escolhe UI ou CLI
- `preencher_fotos_core.py`: regras de deteccao, leitura de planilhas, insercao de imagens e relatorios
- `preencher_fotos_ui.py`: tela com os 2 campos de pasta e log
- `build_exe.bat`: build automatizado com `PyInstaller`
- `PreencherFotosExcel.spec`: configuracao do build (inclui imports `pywin32`)

## Troubleshooting Rapido

### 1) Aba foi ignorada

Motivo comum:
- cabecalhos fora do padrao detectavel (`Nome`, `Foto`, `Imagem`, `Link`)

Acao:
- ajustar cabecalhos na planilha
- ou evoluir o detector no `preencher_fotos_core.py`

### 2) Fotos faltantes

Motivo comum:
- nome da imagem na planilha nao bate exatamente com o arquivo real

Acao:
- conferir log e corrigir o nome na planilha ou no arquivo

### 3) Excel nao abre / erro COM

Verificar:
- Microsoft Excel instalado
- permissao para abrir a planilha
- arquivo nao aberto por outro usuario em lock

### 4) Build falha com pywin32

O `build_exe.bat` ja inclui imports necessarios:
- `win32com`
- `pythoncom`
- `pywintypes`
- `win32timezone`

## Git (Recomendado)

O repositorio pode versionar:

- codigo fonte
- scripts de build
- documentacao

Evite versionar:

- `build/`, `dist/`
- caches Python
- planilhas reais de trabalho
- imagens reais de trabalho
- executavel gerado (opcional)

Veja `.gitignore` para o padrao sugerido.

## Proximos Passos (Opcional)

- Adicionar modo teste (`dry-run`, sem salvar)
- Exportar relatorio `.csv`
- Melhorar detector de cabecalhos para layouts mais variados
- Criar testes unitarios para a deteccao de mapeamento
