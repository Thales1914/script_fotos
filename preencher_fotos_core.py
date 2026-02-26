import os
import re
import sys
from collections import defaultdict
from tkinter import Tk, filedialog

import win32com.client


PASTA_PLANILHAS_PADRAO = "PLANILHAS"
PASTA_IMAGENS_PADRAO = "FOTOS"

EXTENSOES_PLANILHA = {".xlsx", ".xlsm", ".xls"}
EXTENSOES_IMAGEM = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}

MAX_LINHAS_CABECALHO = 30
MAX_COLUNAS_CABECALHO = 250
MARGEM_IMAGEM_X = 2
MARGEM_IMAGEM_Y = 2
PREFIXO_SHAPE = "AUTOIMG"


def criar_root_tk():
    root = Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass
    return root


def emitir_log(log_fn, mensagem=""):
    if log_fn is None:
        return
    log_fn(str(mensagem))


def obter_pasta_base_execucao():
    """Retorna a pasta ao lado do script (ou do .exe, quando empacotado)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def selecionar_pasta(titulo, pasta_inicial):
    root = criar_root_tk()
    try:
        return filedialog.askdirectory(
            title=titulo,
            initialdir=pasta_inicial,
            mustexist=True,
        )
    finally:
        root.destroy()


def normalizar_texto(valor):
    if valor is None:
        return ""
    texto = str(valor).strip()
    if not texto:
        return ""
    return re.sub(r"\s+", " ", texto).lower()


def normalizar_chave_arquivo(valor):
    if valor is None:
        return ""
    nome = os.path.basename(str(valor).strip())
    if not nome:
        return ""
    return re.sub(r"\s+", " ", nome).lower()


def tem_conteudo(valor):
    if valor is None:
        return False
    if isinstance(valor, str):
        return bool(valor.strip())
    return True


def col_num_para_letra(numero):
    resultado = ""
    n = int(numero)
    while n > 0:
        n, resto = divmod(n - 1, 26)
        resultado = chr(65 + resto) + resultado
    return resultado or "A"


def obter_limites_usados(ws):
    try:
        usado = ws.UsedRange
        primeira_linha = int(usado.Row)
        primeira_coluna = int(usado.Column)
        total_linhas = int(usado.Rows.Count)
        total_colunas = int(usado.Columns.Count)

        ultima_linha = primeira_linha + total_linhas - 1
        ultima_coluna = primeira_coluna + total_colunas - 1
        return max(1, ultima_linha), max(1, ultima_coluna)
    except Exception:
        return 1, 1


def extrair_indice_rotulo(texto_normalizado):
    match = re.search(r"(\d+)", texto_normalizado)
    if not match:
        return None
    try:
        return int(match.group(1))
    except Exception:
        return None


def classificar_cabecalho(texto):
    texto_norm = normalizar_texto(texto)
    if not texto_norm:
        return None

    # Prioriza "nome" para nao classificar "nome da imagem" como coluna de foto.
    if "nome" in texto_norm or "arquivo" in texto_norm:
        return {"tipo": "nome", "indice": extrair_indice_rotulo(texto_norm), "texto": texto_norm}

    if "foto" in texto_norm or "imagem" in texto_norm or "link" in texto_norm:
        return {"tipo": "foto", "indice": extrair_indice_rotulo(texto_norm), "texto": texto_norm}

    return None


def montar_pares_cabecalho(candidatos):
    nomes = sorted(
        [c for c in candidatos if c["tipo"] == "nome"],
        key=lambda item: item["coluna"],
    )
    fotos = sorted(
        [c for c in candidatos if c["tipo"] == "foto"],
        key=lambda item: item["coluna"],
    )

    if not nomes or not fotos:
        return []

    pares = []
    colunas_nomes_usadas = set()
    colunas_fotos_usadas = set()

    nomes_por_indice = {}
    fotos_por_indice = {}

    for item in nomes:
        indice = item.get("indice")
        if indice is not None and indice not in nomes_por_indice:
            nomes_por_indice[indice] = item

    for item in fotos:
        indice = item.get("indice")
        if indice is not None and indice not in fotos_por_indice:
            fotos_por_indice[indice] = item

    for indice in sorted(set(nomes_por_indice) & set(fotos_por_indice)):
        nome_col = nomes_por_indice[indice]
        foto_col = fotos_por_indice[indice]
        pares.append(
            {
                "indice": indice,
                "name_col": nome_col["coluna"],
                "target_col": foto_col["coluna"],
                "name_header": nome_col["texto"],
                "target_header": foto_col["texto"],
                "pareado_por_indice": True,
            }
        )
        colunas_nomes_usadas.add(nome_col["coluna"])
        colunas_fotos_usadas.add(foto_col["coluna"])

    nomes_restantes = [n for n in nomes if n["coluna"] not in colunas_nomes_usadas]
    fotos_restantes = [f for f in fotos if f["coluna"] not in colunas_fotos_usadas]

    for nome_col, foto_col in zip(nomes_restantes, fotos_restantes):
        pares.append(
            {
                "indice": None,
                "name_col": nome_col["coluna"],
                "target_col": foto_col["coluna"],
                "name_header": nome_col["texto"],
                "target_header": foto_col["texto"],
                "pareado_por_indice": False,
            }
        )

    pares.sort(key=lambda item: item["name_col"])
    return pares


def detectar_mapeamento_na_aba(ws):
    ultima_linha, ultima_coluna = obter_limites_usados(ws)
    linhas_scan = min(MAX_LINHAS_CABECALHO, max(1, ultima_linha))
    colunas_scan = min(MAX_COLUNAS_CABECALHO, max(1, ultima_coluna))

    melhor = None

    for linha in range(1, linhas_scan + 1):
        candidatos = []
        for coluna in range(1, colunas_scan + 1):
            valor = ws.Cells(linha, coluna).Value2
            classificado = classificar_cabecalho(valor)
            if not classificado:
                continue
            classificado["coluna"] = coluna
            candidatos.append(classificado)

        pares = montar_pares_cabecalho(candidatos)
        if not pares:
            continue

        qtd_nomes = sum(1 for c in candidatos if c["tipo"] == "nome")
        qtd_fotos = sum(1 for c in candidatos if c["tipo"] == "foto")
        qtd_pares_indice = sum(1 for p in pares if p["pareado_por_indice"])
        score = (len(pares), qtd_pares_indice, min(qtd_nomes, qtd_fotos), -linha)

        if melhor is None or score > melhor["score"]:
            melhor = {
                "score": score,
                "header_row": linha,
                "pairs": pares,
                "ultima_linha": ultima_linha,
            }

    return melhor


def nome_shape_automatico(linha, coluna_letra):
    return f"{PREFIXO_SHAPE}_{coluna_letra}{linha}"


def remover_shape_se_existir(ws, nome_shape):
    try:
        ws.Shapes(nome_shape).Delete()
        return True
    except Exception:
        return False


def ajustar_shape_na_celula(shape, celula):
    largura_celula = float(celula.Width)
    altura_celula = float(celula.Height)
    esquerda_celula = float(celula.Left)
    topo_celula = float(celula.Top)

    largura_util = max(2.0, largura_celula - (2 * MARGEM_IMAGEM_X))
    altura_util = max(2.0, altura_celula - (2 * MARGEM_IMAGEM_Y))

    try:
        shape.LockAspectRatio = True
    except Exception:
        pass

    largura_original = float(shape.Width)
    altura_original = float(shape.Height)

    if largura_original > 0 and altura_original > 0:
        escala = min(largura_util / largura_original, altura_util / altura_original)
        if escala <= 0:
            escala = 1.0
        shape.Width = max(1.0, largura_original * escala)
        shape.Height = max(1.0, altura_original * escala)

    shape.Left = esquerda_celula + max(0.0, (largura_celula - float(shape.Width)) / 2)
    shape.Top = topo_celula + max(0.0, (altura_celula - float(shape.Height)) / 2)

    # 1 = xlMoveAndSize
    try:
        shape.Placement = 1
    except Exception:
        pass


def inserir_imagem_na_celula(ws, caminho_imagem, linha, coluna_destino):
    celula = ws.Cells(linha, coluna_destino)
    shape = ws.Shapes.AddPicture(
        Filename=os.path.abspath(caminho_imagem),
        LinkToFile=False,
        SaveWithDocument=True,
        Left=float(celula.Left) + MARGEM_IMAGEM_X,
        Top=float(celula.Top) + MARGEM_IMAGEM_Y,
        Width=-1,
        Height=-1,
    )

    ajustar_shape_na_celula(shape, celula)
    return shape


def indexar_imagens_recursivamente(pasta_imagens):
    indice = {}
    duplicados = defaultdict(list)
    total_imagens = 0

    for raiz, _, arquivos in os.walk(pasta_imagens):
        for nome in arquivos:
            _, ext = os.path.splitext(nome)
            if ext.lower() not in EXTENSOES_IMAGEM:
                continue

            total_imagens += 1
            caminho = os.path.join(raiz, nome)
            chave = normalizar_chave_arquivo(nome)
            if not chave:
                continue

            if chave in indice:
                if not duplicados[chave]:
                    duplicados[chave].append(indice[chave])
                duplicados[chave].append(caminho)
                continue

            indice[chave] = caminho

    return {
        "indice": indice,
        "total_imagens": total_imagens,
        "nomes_duplicados": duplicados,
    }


def listar_planilhas_recursivamente(pasta_planilhas):
    encontrados = []
    for raiz, _, arquivos in os.walk(pasta_planilhas):
        for nome in arquivos:
            if nome.startswith("~$"):
                continue
            _, ext = os.path.splitext(nome)
            if ext.lower() in EXTENSOES_PLANILHA:
                encontrados.append(os.path.join(raiz, nome))
    return sorted(encontrados)


def detectar_pastas_padrao(pasta_base, log_fn=print):
    pasta_planilhas = os.path.join(pasta_base, PASTA_PLANILHAS_PADRAO)
    pasta_imagens = os.path.join(pasta_base, PASTA_IMAGENS_PADRAO)

    if os.path.isdir(pasta_planilhas) and os.path.isdir(pasta_imagens):
        emitir_log(
            log_fn,
            "Pastas padrao encontradas ao lado do script: "
            f'"{PASTA_PLANILHAS_PADRAO}" e "{PASTA_IMAGENS_PADRAO}".'
        )
        return pasta_planilhas, pasta_imagens

    return None, None


def selecionar_origens(pasta_base, log_fn=print):
    pasta_planilhas, pasta_imagens = detectar_pastas_padrao(pasta_base, log_fn=log_fn)

    if not pasta_planilhas:
        pasta_planilhas = selecionar_pasta(
            "Selecione a pasta com as planilhas para preencher",
            pasta_base,
        )

    if not pasta_imagens:
        pasta_imagens = selecionar_pasta(
            "Selecione a pasta com as imagens",
            pasta_base,
        )

    return pasta_planilhas, pasta_imagens


def processar_aba(ws, indice_imagens):
    deteccao = detectar_mapeamento_na_aba(ws)
    if not deteccao:
        return {
            "aba": ws.Name,
            "detectada": False,
            "motivo": "Nenhum mapeamento Nome/Foto detectado",
            "linhas_com_nomes": 0,
            "fotos_referenciadas": 0,
            "fotos_inseridas": 0,
            "fotos_faltantes": 0,
            "mapeamentos": [],
        }

    header_row = deteccao["header_row"]
    ultima_linha = deteccao["ultima_linha"]
    pares = deteccao["pairs"]

    resumo = {
        "aba": ws.Name,
        "detectada": True,
        "header_row": header_row,
        "ultima_linha": ultima_linha,
        "linhas_com_nomes": 0,
        "fotos_referenciadas": 0,
        "fotos_inseridas": 0,
        "fotos_faltantes": 0,
        "shapes_substituidas": 0,
        "mapeamentos": [
            (
                col_num_para_letra(p["name_col"]),
                col_num_para_letra(p["target_col"]),
            )
            for p in pares
        ],
        "faltantes_exemplo": [],
    }

    linha_inicial_dados = header_row + 1
    if linha_inicial_dados > ultima_linha:
        return resumo

    for linha in range(linha_inicial_dados, ultima_linha + 1):
        nomes_na_linha = []
        for par in pares:
            valor_nome = ws.Cells(linha, par["name_col"]).Value2
            chave = normalizar_chave_arquivo(valor_nome)
            nomes_na_linha.append((par, chave))

        if not any(chave for _, chave in nomes_na_linha):
            continue

        resumo["linhas_com_nomes"] += 1

        for par, chave_arquivo in nomes_na_linha:
            if not chave_arquivo:
                continue

            resumo["fotos_referenciadas"] += 1
            caminho_imagem = indice_imagens.get(chave_arquivo)
            col_destino_letra = col_num_para_letra(par["target_col"])

            if not caminho_imagem:
                resumo["fotos_faltantes"] += 1
                if len(resumo["faltantes_exemplo"]) < 10:
                    resumo["faltantes_exemplo"].append(
                        {
                            "linha": linha,
                            "origem": col_num_para_letra(par["name_col"]),
                            "destino": col_destino_letra,
                            "arquivo": chave_arquivo,
                        }
                    )
                continue

            nome_shape = nome_shape_automatico(linha, col_destino_letra)
            if remover_shape_se_existir(ws, nome_shape):
                resumo["shapes_substituidas"] += 1

            shape = inserir_imagem_na_celula(ws, caminho_imagem, linha, par["target_col"])
            try:
                shape.Name = nome_shape
            except Exception:
                pass

            resumo["fotos_inseridas"] += 1

    return resumo


def processar_planilha(excel, caminho_planilha, indice_imagens):
    resultado = {
        "arquivo": caminho_planilha,
        "aberto": False,
        "salvo": False,
        "erro": None,
        "abas": [],
        "totais": {
            "abas_detectadas": 0,
            "linhas_com_nomes": 0,
            "fotos_referenciadas": 0,
            "fotos_inseridas": 0,
            "fotos_faltantes": 0,
            "shapes_substituidas": 0,
        },
    }

    wb = None
    try:
        wb = excel.Workbooks.Open(os.path.abspath(caminho_planilha))
        resultado["aberto"] = True

        if bool(getattr(wb, "ReadOnly", False)):
            resultado["erro"] = "Arquivo aberto em modo somente leitura (nao sera salvo)."

        for ws in wb.Worksheets:
            try:
                resumo_aba = processar_aba(ws, indice_imagens)
            except Exception as exc_aba:
                resumo_aba = {
                    "aba": ws.Name,
                    "detectada": False,
                    "motivo": f"Erro ao processar aba: {exc_aba}",
                    "linhas_com_nomes": 0,
                    "fotos_referenciadas": 0,
                    "fotos_inseridas": 0,
                    "fotos_faltantes": 0,
                    "shapes_substituidas": 0,
                    "mapeamentos": [],
                    "faltantes_exemplo": [],
                }
            resultado["abas"].append(resumo_aba)

            if resumo_aba.get("detectada"):
                resultado["totais"]["abas_detectadas"] += 1

            for chave in (
                "linhas_com_nomes",
                "fotos_referenciadas",
                "fotos_inseridas",
                "fotos_faltantes",
                "shapes_substituidas",
            ):
                resultado["totais"][chave] += int(resumo_aba.get(chave, 0))

        houve_alteracao = (
            resultado["totais"]["fotos_inseridas"] > 0
            or resultado["totais"]["shapes_substituidas"] > 0
        )

        if houve_alteracao and not bool(getattr(wb, "ReadOnly", False)):
            wb.Save()
            resultado["salvo"] = True

    except Exception as exc:
        resultado["erro"] = str(exc)
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass

    return resultado


def imprimir_resumo_planilha(resultado, log_fn=print):
    emitir_log(log_fn, "")
    emitir_log(log_fn, f'Arquivo: {resultado["arquivo"]}')

    if resultado["erro"] and not resultado["aberto"]:
        emitir_log(log_fn, f'  Erro ao abrir/processar: {resultado["erro"]}')
        return

    if resultado["erro"] and resultado["aberto"]:
        emitir_log(log_fn, f'  Aviso: {resultado["erro"]}')

    emitir_log(
        log_fn,
        "  Totais: "
        f'abas_detectadas={resultado["totais"]["abas_detectadas"]}, '
        f'linhas_com_nomes={resultado["totais"]["linhas_com_nomes"]}, '
        f'fotos_referenciadas={resultado["totais"]["fotos_referenciadas"]}, '
        f'fotos_inseridas={resultado["totais"]["fotos_inseridas"]}, '
        f'fotos_faltantes={resultado["totais"]["fotos_faltantes"]}, '
        f'shapes_substituidas={resultado["totais"]["shapes_substituidas"]}, '
        f'salvo={resultado["salvo"]}'
    )

    for aba in resultado["abas"]:
        if not aba.get("detectada"):
            emitir_log(log_fn, f'  - Aba "{aba["aba"]}": ignorada ({aba["motivo"]})')
            continue

        mapeamentos_txt = ", ".join([f"{orig}->{dest}" for orig, dest in aba["mapeamentos"]])
        emitir_log(
            log_fn,
            f'  - Aba "{aba["aba"]}": '
            f'cabecalho_linha={aba["header_row"]}, '
            f'mapeamentos=[{mapeamentos_txt}], '
            f'linhas={aba["linhas_com_nomes"]}, '
            f'referenciadas={aba["fotos_referenciadas"]}, '
            f'inseridas={aba["fotos_inseridas"]}, '
            f'faltantes={aba["fotos_faltantes"]}'
        )

        if aba["faltantes_exemplo"]:
            for item in aba["faltantes_exemplo"][:3]:
                emitir_log(
                    log_fn,
                    "    faltante: "
                    f'linha={item["linha"]}, '
                    f'{item["origem"]}->{item["destino"]}, '
                    f'arquivo="{item["arquivo"]}"'
                )


def inserir_imagens_em_lote(pasta_planilhas=None, pasta_imagens=None, log_fn=print):
    pasta_base = obter_pasta_base_execucao()
    log = log_fn or print

    if not pasta_planilhas or not pasta_imagens:
        pasta_planilhas, pasta_imagens = selecionar_origens(pasta_base, log_fn=log)

    if not pasta_planilhas:
        emitir_log(log, "Nenhuma pasta de planilhas selecionada. Encerrando.")
        return None

    if not pasta_imagens:
        emitir_log(log, "Nenhuma pasta de imagens selecionada. Encerrando.")
        return None

    if not os.path.isdir(pasta_planilhas):
        emitir_log(log, f"Pasta de planilhas nao encontrada: {pasta_planilhas}")
        return None

    if not os.path.isdir(pasta_imagens):
        emitir_log(log, f"Pasta de imagens nao encontrada: {pasta_imagens}")
        return None

    emitir_log(log, f"Pasta de planilhas: {pasta_planilhas}")
    emitir_log(log, f"Pasta de imagens:   {pasta_imagens}")

    planilhas = listar_planilhas_recursivamente(pasta_planilhas)
    if not planilhas:
        emitir_log(log, "Nenhuma planilha encontrada para processar.")
        return None

    dados_imagens = indexar_imagens_recursivamente(pasta_imagens)
    indice_imagens = dados_imagens["indice"]
    nomes_duplicados = dados_imagens["nomes_duplicados"]

    if not indice_imagens:
        emitir_log(log, "Nenhuma imagem encontrada para indexacao.")
        return None

    emitir_log(
        log,
        "Leitura concluida: "
        f'planilhas={len(planilhas)}, '
        f'imagens_indexadas={len(indice_imagens)}, '
        f'imagens_total={dados_imagens["total_imagens"]}, '
        f'nomes_duplicados={len(nomes_duplicados)}'
    )

    if nomes_duplicados:
        emitir_log(
            log,
            "Aviso: existem nomes de imagem duplicados em subpastas. O script usa a primeira ocorrencia.",
        )

    excel = None
    resultados = []

    totais_gerais = {
        "planilhas": len(planilhas),
        "planilhas_salvas": 0,
        "abas_detectadas": 0,
        "linhas_com_nomes": 0,
        "fotos_referenciadas": 0,
        "fotos_inseridas": 0,
        "fotos_faltantes": 0,
        "shapes_substituidas": 0,
        "erros": 0,
    }

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        for caminho_planilha in planilhas:
            resultado = processar_planilha(excel, caminho_planilha, indice_imagens)
            resultados.append(resultado)
            imprimir_resumo_planilha(resultado, log_fn=log)

            if resultado["salvo"]:
                totais_gerais["planilhas_salvas"] += 1
            if resultado["erro"] and not resultado["aberto"]:
                totais_gerais["erros"] += 1

            for chave in (
                "abas_detectadas",
                "linhas_com_nomes",
                "fotos_referenciadas",
                "fotos_inseridas",
                "fotos_faltantes",
                "shapes_substituidas",
            ):
                totais_gerais[chave] += int(resultado["totais"].get(chave, 0))

    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

    emitir_log(log, "")
    emitir_log(log, "Resumo geral:")
    emitir_log(
        log,
        f'  planilhas={totais_gerais["planilhas"]}, '
        f'planilhas_salvas={totais_gerais["planilhas_salvas"]}, '
        f'abas_detectadas={totais_gerais["abas_detectadas"]}, '
        f'linhas_com_nomes={totais_gerais["linhas_com_nomes"]}, '
        f'fotos_referenciadas={totais_gerais["fotos_referenciadas"]}, '
        f'fotos_inseridas={totais_gerais["fotos_inseridas"]}, '
        f'fotos_faltantes={totais_gerais["fotos_faltantes"]}, '
        f'shapes_substituidas={totais_gerais["shapes_substituidas"]}, '
        f'erros={totais_gerais["erros"]}'
    )

    return {
        "totais": totais_gerais,
        "resultados": resultados,
        "pasta_planilhas": pasta_planilhas,
        "pasta_imagens": pasta_imagens,
    }

