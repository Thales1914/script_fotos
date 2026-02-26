import sys

from preencher_fotos_core import inserir_imagens_em_lote
from preencher_fotos_ui import executar_interface


if __name__ == "__main__":
    if "--cli" in sys.argv:
        inserir_imagens_em_lote()
    else:
        executar_interface()
