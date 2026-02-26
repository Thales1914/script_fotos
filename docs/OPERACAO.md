# Operacao e Publicacao no Git

## Uso Local (Operacao)

### Estrutura recomendada para uso

```text
PARA_RODAR/
|- PreencherFotosExcel.exe
|- Executar_Preenchimento.bat
|- README_USO.txt
|- PLANILHAS/
`- FOTOS/
```

### Passos de uso

1. Colocar planilhas em `PLANILHAS/`
2. Colocar imagens em `FOTOS/`
3. Executar o `.exe` (ou `.bat`)
4. Verificar log e resumo final

## Build / Rebuild do Executavel

### Dependencias

```powershell
python -m pip install pyinstaller pywin32
```

### Gerar

```powershell
.\build_exe.bat
```

### Atualizar pasta de uso

Copiar:

- `dist/PreencherFotosExcel.exe`

para:

- `PARA_RODAR/PreencherFotosExcel.exe`

## Publicacao no Git (Recomendado)

Versionar:

- `Inserir foto.py`
- `preencher_fotos_core.py`
- `preencher_fotos_ui.py`
- `build_exe.bat`
- `PreencherFotosExcel.spec`
- `README.md`
- `docs/`

Evitar versionar:

- `build/`
- `dist/`
- `__pycache__/`
- `PARA_RODAR/PLANILHAS/*` (dados de trabalho)
- `PARA_RODAR/FOTOS/*` (dados de trabalho)
- `PARA_RODAR/PreencherFotosExcel.exe` (binario gerado, opcional)

## Sugestao de Primeiro Commit

```powershell
git init
git add .
git status
git commit -m "feat: ferramenta de preenchimento automatico de fotos em Excel"
```

## Observacao sobre Dados Reais

Se as planilhas/imagens tiverem dados sensiveis ou arquivos grandes, mantenha fora do Git.

Para compartilhar a ferramenta com usuarios finais, envie apenas a pasta `PARA_RODAR` preparada (ou gere um zip de distribuicao).
