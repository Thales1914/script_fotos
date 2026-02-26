# Arquitetura do Projeto

## Objetivo

Separar responsabilidades para facilitar manutencao:

- entrada (CLI/UI)
- motor de processamento
- interface grafica

## Modulos

## `Inserir foto.py`

Ponto de entrada do projeto.

Responsabilidades:

- abrir interface grafica por padrao
- permitir modo CLI com `--cli`

## `preencher_fotos_core.py`

Motor principal.

Principais responsabilidades:

- localizar pastas e arquivos
- detectar mapeamento `Nome -> Foto` nas abas
- abrir planilhas via Excel COM
- inserir imagens e salvar
- gerar resumo por aba / planilha / lote

Blocos logicos principais:

- Utilitarios
  - normalizacao de texto e nomes
  - conversao de coluna numerica para letra
- Deteccao de cabecalho
  - classificacao de cabecalhos
  - pareamento por indice (`01`, `02`, `03`)
- Operacao em Excel
  - inserir shape
  - ajustar imagem dentro da celula
  - substituir shapes gerados anteriormente
- Processamento em lote
  - indexar imagens recursivamente
  - processar todas as planilhas recursivamente

## `preencher_fotos_ui.py`

Interface grafica (Tkinter).

Responsabilidades:

- exibir dois campos de pasta (planilhas/fotos)
- abrir seletores de pasta
- iniciar processamento em thread separada
- mostrar log e status

Detalhes importantes:

- o processamento roda em thread para nao travar a interface
- a UI usa fila (`queue`) para receber logs e resultado final
- inicializa COM na thread de trabalho (`pythoncom.CoInitialize`)

## Fluxo de Execucao

1. Usuario abre o programa
2. Interface define as pastas (manual ou padrao)
3. UI chama `inserir_imagens_em_lote(...)` do core
4. Core indexa imagens
5. Core abre planilhas e processa abas
6. Core devolve resumo final
7. UI atualiza status e exibe mensagem de conclusao

## Regras de Deteccao de Colunas

Cabecalhos de origem (nome de arquivo):

- `nome`
- `arquivo`

Cabecalhos de destino (foto):

- `foto`
- `imagem`
- `link`

Estrategia:

1. Tentar parear por indice (`Nome 01` com `Foto 01`)
2. Se faltar indice, parear por ordem de coluna

## Estrategia de Reexecucao

O sistema nomeia as imagens inseridas com prefixo interno (`AUTOIMG_...`) e tenta remover a anterior antes de inserir novamente na mesma celula-alvo.

Isso ajuda a evitar duplicacao quando a ferramenta e executada mais de uma vez.

