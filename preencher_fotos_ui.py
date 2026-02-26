import os
import queue
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pythoncom

from preencher_fotos_core import (
    PASTA_IMAGENS_PADRAO,
    PASTA_PLANILHAS_PADRAO,
    inserir_imagens_em_lote,
    obter_pasta_base_execucao,
)


class MiniInterfacePreenchimento:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Preenchimento de Fotos em Planilhas")
        self.root.geometry("920x620")
        self.root.minsize(820, 520)

        self.pasta_base = obter_pasta_base_execucao()
        self.var_planilhas = tk.StringVar()
        self.var_imagens = tk.StringVar()
        self.var_status = tk.StringVar(value="Selecione as pastas e clique em Iniciar.")

        self._fila_eventos = queue.Queue()
        self._thread_execucao = None
        self._em_execucao = False

        self._montar_interface()
        self._aplicar_pastas_padrao_sem_alerta()
        self.root.after(120, self._processar_eventos_fila)

    def _montar_interface(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        frame_topo = ttk.Frame(self.root, padding=12)
        frame_topo.grid(row=0, column=0, sticky="nsew")
        frame_topo.columnconfigure(1, weight=1)

        ttk.Label(frame_topo, text="Pasta das planilhas").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
        self.entry_planilhas = ttk.Entry(frame_topo, textvariable=self.var_planilhas)
        self.entry_planilhas.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        self.btn_planilhas = ttk.Button(frame_topo, text="Selecionar...", command=self._selecionar_planilhas)
        self.btn_planilhas.grid(row=0, column=2, sticky="ew", padx=(8, 0), pady=(0, 8))

        ttk.Label(frame_topo, text="Pasta das fotos").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
        self.entry_imagens = ttk.Entry(frame_topo, textvariable=self.var_imagens)
        self.entry_imagens.grid(row=1, column=1, sticky="ew", pady=(0, 8))
        self.btn_imagens = ttk.Button(frame_topo, text="Selecionar...", command=self._selecionar_imagens)
        self.btn_imagens.grid(row=1, column=2, sticky="ew", padx=(8, 0), pady=(0, 8))

        frame_acoes = ttk.Frame(self.root, padding=(12, 0, 12, 8))
        frame_acoes.grid(row=1, column=0, sticky="ew")
        frame_acoes.columnconfigure(3, weight=1)

        self.btn_padrao = ttk.Button(frame_acoes, text="Usar pastas padrao", command=self._usar_pastas_padrao)
        self.btn_padrao.grid(row=0, column=0, sticky="w")

        self.btn_limpar_log = ttk.Button(frame_acoes, text="Limpar log", command=self._limpar_log)
        self.btn_limpar_log.grid(row=0, column=1, sticky="w", padx=(8, 0))

        self.btn_iniciar = ttk.Button(frame_acoes, text="Iniciar preenchimento", command=self._iniciar_processamento)
        self.btn_iniciar.grid(row=0, column=2, sticky="e", padx=(12, 0))

        frame_log = ttk.LabelFrame(self.root, text="Log de processamento", padding=(8, 8))
        frame_log.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 8))
        frame_log.columnconfigure(0, weight=1)
        frame_log.rowconfigure(0, weight=1)

        self.txt_log = tk.Text(frame_log, wrap="word", height=20)
        self.txt_log.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(frame_log, orient="vertical", command=self.txt_log.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.txt_log.configure(yscrollcommand=scrollbar.set)

        frame_rodape = ttk.Frame(self.root, padding=(12, 0, 12, 12))
        frame_rodape.grid(row=3, column=0, sticky="ew")
        frame_rodape.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(frame_rodape, mode="indeterminate")
        self.progress.grid(row=0, column=0, sticky="ew")
        ttk.Label(frame_rodape, textvariable=self.var_status).grid(row=1, column=0, sticky="w", pady=(6, 0))

    def _aplicar_pastas_padrao_sem_alerta(self):
        plan = os.path.join(self.pasta_base, PASTA_PLANILHAS_PADRAO)
        imgs = os.path.join(self.pasta_base, PASTA_IMAGENS_PADRAO)
        if os.path.isdir(plan):
            self.var_planilhas.set(plan)
        if os.path.isdir(imgs):
            self.var_imagens.set(imgs)

    def _usar_pastas_padrao(self):
        plan = os.path.join(self.pasta_base, PASTA_PLANILHAS_PADRAO)
        imgs = os.path.join(self.pasta_base, PASTA_IMAGENS_PADRAO)
        self.var_planilhas.set(plan)
        self.var_imagens.set(imgs)
        self._log_gui("Pastas padrao preenchidas nos campos.")

    def _selecionar_planilhas(self):
        caminho = filedialog.askdirectory(
            title="Selecione a pasta com as planilhas",
            initialdir=self.var_planilhas.get() or self.pasta_base,
            mustexist=True,
            parent=self.root,
        )
        if caminho:
            self.var_planilhas.set(caminho)

    def _selecionar_imagens(self):
        caminho = filedialog.askdirectory(
            title="Selecione a pasta com as imagens",
            initialdir=self.var_imagens.get() or self.pasta_base,
            mustexist=True,
            parent=self.root,
        )
        if caminho:
            self.var_imagens.set(caminho)

    def _limpar_log(self):
        self.txt_log.delete("1.0", "end")

    def _log_gui(self, mensagem):
        self.txt_log.insert("end", f"{mensagem}\n")
        self.txt_log.see("end")
        self.root.update_idletasks()

    def _log_worker(self, mensagem):
        self._fila_eventos.put(("log", mensagem))

    def _set_em_execucao(self, em_execucao):
        self._em_execucao = em_execucao
        estado = "disabled" if em_execucao else "normal"
        for widget in (
            self.entry_planilhas,
            self.entry_imagens,
            self.btn_planilhas,
            self.btn_imagens,
            self.btn_padrao,
            self.btn_limpar_log,
            self.btn_iniciar,
        ):
            try:
                widget.configure(state=estado)
            except Exception:
                pass

        if em_execucao:
            self.btn_iniciar.configure(text="Processando...")
            self.progress.start(10)
            self.var_status.set("Processando planilhas...")
        else:
            self.btn_iniciar.configure(text="Iniciar preenchimento")
            self.progress.stop()

    def _iniciar_processamento(self):
        if self._em_execucao:
            return

        pasta_planilhas = self.var_planilhas.get().strip()
        pasta_imagens = self.var_imagens.get().strip()

        if not pasta_planilhas or not pasta_imagens:
            messagebox.showwarning(
                "Pastas obrigatorias",
                "Preencha os dois campos: pasta das planilhas e pasta das fotos.",
                parent=self.root,
            )
            return

        if not os.path.isdir(pasta_planilhas):
            messagebox.showerror("Pasta invalida", f"Pasta de planilhas nao encontrada:\n{pasta_planilhas}", parent=self.root)
            return

        if not os.path.isdir(pasta_imagens):
            messagebox.showerror("Pasta invalida", f"Pasta de imagens nao encontrada:\n{pasta_imagens}", parent=self.root)
            return

        self._set_em_execucao(True)
        self._log_gui("")
        self._log_gui("=== Inicio do processamento ===")
        self._log_gui(f"Planilhas: {pasta_planilhas}")
        self._log_gui(f"Imagens:   {pasta_imagens}")

        self._thread_execucao = threading.Thread(
            target=self._executar_processamento_em_thread,
            args=(pasta_planilhas, pasta_imagens),
            daemon=True,
        )
        self._thread_execucao.start()

    def _executar_processamento_em_thread(self, pasta_planilhas, pasta_imagens):
        try:
            pythoncom.CoInitialize()
            resultado = inserir_imagens_em_lote(
                pasta_planilhas=pasta_planilhas,
                pasta_imagens=pasta_imagens,
                log_fn=self._log_worker,
            )
            self._fila_eventos.put(("done", resultado))
        except Exception:
            self._fila_eventos.put(("error", traceback.format_exc()))
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _processar_eventos_fila(self):
        try:
            while True:
                tipo, payload = self._fila_eventos.get_nowait()
                if tipo == "log":
                    self._log_gui(payload)
                elif tipo == "done":
                    self._finalizar_processamento(payload)
                elif tipo == "error":
                    self._finalizar_com_erro(payload)
        except queue.Empty:
            pass
        finally:
            self.root.after(120, self._processar_eventos_fila)

    def _finalizar_processamento(self, resultado):
        self._set_em_execucao(False)
        self._log_gui("=== Fim do processamento ===")

        if not resultado:
            self.var_status.set("Processamento encerrado sem executar.")
            return

        totais = resultado.get("totais", {})
        self.var_status.set(
            "Concluido: "
            f'planilhas={totais.get("planilhas_salvas", 0)}/{totais.get("planilhas", 0)} salvas, '
            f'fotos_inseridas={totais.get("fotos_inseridas", 0)}, '
            f'fotos_faltantes={totais.get("fotos_faltantes", 0)}'
        )
        messagebox.showinfo(
            "Processamento concluido",
            self.var_status.get(),
            parent=self.root,
        )

    def _finalizar_com_erro(self, detalhes):
        self._set_em_execucao(False)
        self.var_status.set("Erro durante o processamento.")
        self._log_gui("")
        self._log_gui("ERRO:")
        for linha in str(detalhes).splitlines():
            self._log_gui(linha)
        messagebox.showerror(
            "Erro",
            "Ocorreu um erro durante o processamento. Veja o log para detalhes.",
            parent=self.root,
        )

    def executar(self):
        self.root.mainloop()


def executar_interface():
    app = MiniInterfacePreenchimento()
    app.executar()

