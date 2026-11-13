"""
Interface Gráfica — Central de Automação de Dados
===================================================
Executa a pipeline de scripts em sequência, exibindo os logs de cada etapa
**em tempo real** (linha a linha), tanto no modo de desenvolvimento quanto
no executável empacotado pelo PyInstaller.

Ordem de execução:
  1. extrator_saldo.py
  2. atualizar_bruta.py
  3. gerador_conf.py
  4. gerador_atrasos.py
  5. atualizar_conf.py
  6. atualizar_atr.py
"""

import customtkinter as ctk
import subprocess
import threading
import sys
import os
import shutil
import time
from datetime import datetime, timedelta
from pathlib import Path

from src.paths import RESOURCE_ROOT, APP_ROOT, ensure_portable_data_files

# Importações explícitas para o PyInstaller detectar e embutir
import selenium            # noqa: F401
import selenium.common     # noqa: F401
import selenium.webdriver   # noqa: F401
import pandas              # noqa: F401
import numpy               # noqa: F401
import openpyxl            # noqa: F401
import requests            # noqa: F401

ensure_portable_data_files()

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# ══════════════════════════════════════════════════════════════════════════════
# Constantes
# ══════════════════════════════════════════════════════════════════════════════
SCRIPTS = [
    "extrator_saldo.py",
    "atualizar_bruta.py",
    "gerador_conf.py",
    "gerador_atrasos.py",
    "atualizar_conf.py",
    "atualizar_atr.py",
]

NOMES_ETAPAS = [
    "Extrator de Saldo",
    "Atualizar Planilha Bruta",
    "Gerador de Conferência",
    "Gerador de Atrasos",
    "Atualizar Conferência",
    "Atualizar Atrasos",
]

DESCRICOES_ETAPAS = [
    "Requests → Consulta o ERP e salva os relatórios consolidados",
    "Pandas → Atualiza planilha bruta com dados extraídos",
    "Pandas/Openpyxl → Gera planilha de conferência D+2",
    "Selenium/Pandas → Extrai atrasos da Web e gera relatório",
    "Requests → Envia dados de conferência para Google Sheets",
    "Requests → Envia relatório de atrasos para Google Sheets",
]


def calcular_dias_uteis_anteriores(data_base, dias_uteis=2):
    """Calcula dias úteis retroativos para sugerir as datas iniciais."""
    data_atual = data_base
    contagem = 0
    while contagem < dias_uteis:
        data_atual -= timedelta(days=1)
        if data_atual.weekday() < 5:
            contagem += 1
    return data_atual


def formatar_duracao(segundos):
    """Formata segundos em mm:ss."""
    m, s = divmod(int(segundos), 60)
    return f"{m:02d}:{s:02d}"


def verificar_chrome():
    """Verifica se o Google Chrome está acessível no sistema."""
    locais_chrome = [
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
    ]
    for local in locais_chrome:
        if local and os.path.exists(local):
            return True
    # Fallback: tenta via shutil
    return shutil.which("chrome") is not None or shutil.which("google-chrome") is not None


# ══════════════════════════════════════════════════════════════════════════════
# Classe Principal
# ══════════════════════════════════════════════════════════════════════════════
class AppExecutor(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Central de Automação de Dados  v5.0")
        self.geometry("1200x920")
        self.minsize(1050, 800)
        self.configure(fg_color="#0b0f1a")

        self.processo_atual = None
        self.executando = False
        self.interrompido = False
        self.etapas_widgets = []  # lista de dicts {frame, label_nome, label_desc, label_tempo, label_status}
        self.inicio_etapa = None

        self._criar_interface()

    # ──────────────────────────────────────────────────────────────────────────
    # Interface
    # ──────────────────────────────────────────────────────────────────────────
    def _criar_interface(self):
        self.grid_columnconfigure(0, weight=1)

        # ── Título ──
        titulo = ctk.CTkLabel(
            self, text="⚡ Pipeline de Extração e Atualização",
            font=ctk.CTkFont(size=26, weight="bold"),
            text_color="#e0e7ff",
        )
        titulo.grid(row=0, column=0, pady=(22, 8), sticky="n")

        # ── Painel de Entrada ──
        self.frame_topo = ctk.CTkFrame(self, corner_radius=18, fg_color="#111827",
                                        border_width=1, border_color="#1e3a5f")
        self.frame_topo.grid(row=1, column=0, padx=28, pady=(0, 12), sticky="ew")
        for c in range(4):
            self.frame_topo.grid_columnconfigure(c, weight=1)

        hoje = datetime.now()
        data_ini = calcular_dias_uteis_anteriores(hoje, 2).strftime("%d/%m/%Y")
        data_fim = hoje.strftime("%d/%m/%Y")

        campos = [
            ("Data inicial", data_ini, False),
            ("Data final", data_fim, False),
            ("Usuário (Intranet)", "", False),
            ("Senha", "", True),
        ]
        self.entries = {}
        for col, (rotulo, padrao, senha) in enumerate(campos):
            ctk.CTkLabel(self.frame_topo, text=rotulo, font=ctk.CTkFont(weight="bold"),
                         text_color="#94a3b8").grid(row=0, column=col, padx=14, pady=(14, 4), sticky="w")
            entry = ctk.CTkEntry(self.frame_topo, width=200, height=38,
                                  border_color="#334155", fg_color="#0f172a",
                                  text_color="#e2e8f0",
                                  show="•" if senha else "")
            entry.insert(0, padrao)
            entry.grid(row=1, column=col, padx=14, pady=(0, 14), sticky="ew")
            self.entries[rotulo] = entry

        # ── Status global + barra de progresso ──
        self.lbl_status = ctk.CTkLabel(
            self, text="Aguardando inicialização...",
            font=ctk.CTkFont(size=14, weight="bold"), text_color="#94a3b8",
        )
        self.lbl_status.grid(row=2, column=0, pady=(4, 2), sticky="n")

        self.progress_bar = ctk.CTkProgressBar(self, width=940, height=14,
                                                progress_color="#6366f1",
                                                fg_color="#1e1b4b", corner_radius=7)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=3, column=0, padx=28, pady=(2, 8), sticky="ew")

        # ── Painel central com etapas + console ──
        self.frame_conteudo = ctk.CTkFrame(self, corner_radius=18, fg_color="#111827",
                                            border_width=1, border_color="#1e3a5f")
        self.frame_conteudo.grid(row=4, column=0, padx=28, pady=(0, 8), sticky="nsew")
        self.frame_conteudo.grid_columnconfigure(0, weight=0)
        self.frame_conteudo.grid_columnconfigure(1, weight=1)
        self.frame_conteudo.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # ── Coluna esquerda: etapas ──
        frame_esq = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent", width=360)
        frame_esq.grid(row=0, column=0, padx=(14, 6), pady=14, sticky="nsew")
        frame_esq.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame_esq, text="Etapas do Pipeline",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color="#c7d2fe").grid(row=0, column=0, pady=(0, 8), sticky="w")

        for idx, (nome, desc) in enumerate(zip(NOMES_ETAPAS, DESCRICOES_ETAPAS)):
            f = ctk.CTkFrame(frame_esq, fg_color="#1e2333", corner_radius=10,
                             border_width=1, border_color="#2d3748")
            f.grid(row=idx + 1, column=0, sticky="ew", pady=3)
            f.grid_columnconfigure(1, weight=1)

            lbl_num = ctk.CTkLabel(f, text=f" {idx + 1} ", font=ctk.CTkFont(size=13, weight="bold"),
                                    fg_color="#334155", corner_radius=6,
                                    text_color="#cbd5e1", width=30)
            lbl_num.grid(row=0, column=0, rowspan=2, padx=(8, 6), pady=6)

            lbl_nome = ctk.CTkLabel(f, text=nome, font=ctk.CTkFont(size=13, weight="bold"),
                                     text_color="#e2e8f0", anchor="w")
            lbl_nome.grid(row=0, column=1, padx=(0, 8), pady=(6, 0), sticky="ew")

            lbl_desc = ctk.CTkLabel(f, text=desc, font=ctk.CTkFont(size=10),
                                     text_color="#64748b", anchor="w")
            lbl_desc.grid(row=1, column=1, padx=(0, 8), pady=(0, 2), sticky="ew")

            lbl_status = ctk.CTkLabel(f, text="", font=ctk.CTkFont(size=10, weight="bold"),
                                       text_color="#94a3b8", anchor="w")
            lbl_status.grid(row=2, column=1, padx=(0, 8), pady=(0, 6), sticky="ew")

            lbl_tempo = ctk.CTkLabel(f, text="", font=ctk.CTkFont(size=10),
                                      text_color="#64748b", anchor="e")
            lbl_tempo.grid(row=0, column=2, padx=(0, 10), pady=(6, 0), sticky="e")

            self.etapas_widgets.append({
                "frame": f, "label_nome": lbl_nome, "label_desc": lbl_desc,
                "label_status": lbl_status, "label_tempo": lbl_tempo, "label_num": lbl_num,
            })

        # ── Coluna direita: console ──
        frame_dir = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        frame_dir.grid(row=0, column=1, padx=(6, 14), pady=14, sticky="nsew")
        frame_dir.grid_columnconfigure(0, weight=1)
        frame_dir.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(frame_dir, text="Console de Saída (Tempo Real)",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color="#c7d2fe").grid(row=0, column=0, pady=(0, 6), sticky="w")

        self.txt_console = ctk.CTkTextbox(
            frame_dir, font=ctk.CTkFont(family="Consolas", size=11),
            fg_color="#0a0e17", text_color="#a3e635",
            border_color="#1e293b", corner_radius=10,
            border_width=1,
        )
        self.txt_console.grid(row=1, column=0, sticky="nsew")

        # ── Botões ──
        frame_botoes = ctk.CTkFrame(self, fg_color="transparent")
        frame_botoes.grid(row=5, column=0, pady=(4, 18))

        self.btn_iniciar = ctk.CTkButton(
            frame_botoes, text="▶  Iniciar Pipeline",
            command=self.iniciar_fluxo,
            fg_color="#22c55e", hover_color="#16a34a",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=48, width=260, corner_radius=12,
        )
        self.btn_iniciar.grid(row=0, column=0, padx=12)

        self.btn_parar = ctk.CTkButton(
            frame_botoes, text="⏹  Interromper",
            command=self.interromper_fluxo,
            fg_color="#dc2626", hover_color="#b91c1c",
            state="disabled",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=48, width=260, corner_radius=12,
        )
        self.btn_parar.grid(row=0, column=1, padx=12)

    # ──────────────────────────────────────────────────────────────────────────
    # Controle de etapas
    # ──────────────────────────────────────────────────────────────────────────
    def _resetar_etapas(self):
        for w in self.etapas_widgets:
            w["frame"].configure(fg_color="#1e2333", border_color="#2d3748")
            w["label_nome"].configure(text_color="#e2e8f0")
            w["label_status"].configure(text="", text_color="#94a3b8")
            w["label_tempo"].configure(text="")
            w["label_num"].configure(fg_color="#334155", text_color="#cbd5e1")

    def _marcar_etapa(self, idx, estado, duracao=None):
        w = self.etapas_widgets[idx]
        cores = {
            "executando": {"bg": "#1e3a5f", "border": "#6366f1", "num_bg": "#6366f1",
                           "nome": "#e0e7ff", "status": "⏳ Executando...",
                           "status_cor": "#a5b4fc"},
            "concluido":  {"bg": "#14532d", "border": "#22c55e", "num_bg": "#22c55e",
                           "nome": "#d1fae5", "status": "✅ Concluído",
                           "status_cor": "#86efac"},
            "falha":      {"bg": "#7f1d1d", "border": "#ef4444", "num_bg": "#ef4444",
                           "nome": "#fecaca", "status": "❌ Falhou",
                           "status_cor": "#fca5a5"},
            "cancelado":  {"bg": "#713f12", "border": "#f59e0b", "num_bg": "#f59e0b",
                           "nome": "#fef3c7", "status": "⚠️ Cancelado",
                           "status_cor": "#fcd34d"},
        }
        c = cores.get(estado, cores["executando"])
        w["frame"].configure(fg_color=c["bg"], border_color=c["border"])
        w["label_nome"].configure(text_color=c["nome"])
        w["label_num"].configure(fg_color=c["num_bg"], text_color="#ffffff")
        w["label_status"].configure(text=c["status"], text_color=c["status_cor"])
        if duracao is not None:
            w["label_tempo"].configure(text=f"⏱ {formatar_duracao(duracao)}")

    # ──────────────────────────────────────────────────────────────────────────
    # Log e status
    # ──────────────────────────────────────────────────────────────────────────
    def _iniciar_processador_logs(self):
        """Lê os logs da fila e insere no console em lotes para não travar a UI."""
        if hasattr(self, "_log_queue"):
            import queue
            lote = []
            try:
                for _ in range(50): # Limite de 50 linhas por frame (60fps)
                    lote.append(self._log_queue.get_nowait())
            except queue.Empty:
                pass

            if lote:
                self.txt_console.insert("end", "\n".join(lote) + "\n")
                self.txt_console.see("end")

        self.after(50, self._iniciar_processador_logs)

    def _log(self, msg):
        """Thread-safe: adiciona na fila de logs."""
        if not hasattr(self, "_log_queue"):
            import queue
            self._log_queue = queue.Queue()
            self._iniciar_processador_logs()
        self._log_queue.put(msg)

    def _atualizar_status(self, texto, progresso):
        self.after(0, self._atualizar_status_main, texto, progresso)

    def _atualizar_status_main(self, texto, progresso):
        self.lbl_status.configure(text=texto, text_color="#c7d2fe")
        self.progress_bar.set(max(0.0, min(1.0, progresso)))

    # ──────────────────────────────────────────────────────────────────────────
    # Fluxo principal
    # ──────────────────────────────────────────────────────────────────────────
    def iniciar_fluxo(self):
        dt_ini = self.entries["Data inicial"].get().strip()
        dt_fim = self.entries["Data final"].get().strip()

        # Validação de datas
        try:
            datetime.strptime(dt_ini, "%d/%m/%Y")
            datetime.strptime(dt_fim, "%d/%m/%Y")
        except ValueError:
            self._log("❌ ERRO: Datas no formato inválido. Use DD/MM/AAAA.")
            return

        # Validação de credenciais
        usuario = self.entries["Usuário (Intranet)"].get().strip()
        senha = self.entries["Senha"].get().strip()
        if not usuario or not senha:
            self._log("⚠️ Preencha usuário e senha da Intranet antes de iniciar.")
            return

        # Verificação do Chrome
        if not verificar_chrome():
            self._log("❌ Google Chrome não encontrado no sistema!")
            self._log("   → Instale o Chrome em: https://www.google.com/chrome/")
            self._log("   → O Chrome é necessário para as etapas que usam Selenium, como a geração de atrasos.")
            return

        if not self.executando:
            self.executando = True
            self.interrompido = False
            self.btn_iniciar.configure(state="disabled")
            self.btn_parar.configure(state="normal")
            self.txt_console.delete("1.0", "end")
            self._resetar_etapas()
            threading.Thread(target=self._rodar_pipeline, daemon=True).start()

    def interromper_fluxo(self):
        if self.executando:
            self.interrompido = True
            self._log("\n⚠️ [SOLICITAÇÃO DE INTERRUPÇÃO ENVIADA]...")
            if self.processo_atual and self.processo_atual.poll() is None:
                self.processo_atual.terminate()
            self.btn_parar.configure(state="disabled")

    def _obter_python_executavel(self):
        """
        Retorna o interpretador Python correto para executar subprocessos.
        - Modo desenvolvimento: sys.executable (python.exe)
        - Modo empacotado: usa o python embarcado ou o próprio executável com flag
        """
        if getattr(sys, "frozen", False):
            # No modo empacotado, usamos o próprio exe com --run-script
            return sys.executable
        return sys.executable

    def _rodar_pipeline(self):
        total = len(SCRIPTS)
        inicio_total = time.time()

        self._log("═" * 72)
        self._log(f"🚀 PIPELINE INICIADO — {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        self._log(f"   Período: {self.entries['Data inicial'].get()} → {self.entries['Data final'].get()}")
        self._log(f"   Usuário: {self.entries['Usuário (Intranet)'].get()}")
        self._log(f"   Etapas: {total}")
        self._log("═" * 72)

        for idx, script_nome in enumerate(SCRIPTS):
            if self.interrompido:
                break

            caminho_script = str(RESOURCE_ROOT / "src" / "etl" / script_nome)
            self.after(0, self._marcar_etapa, idx, "executando")
            self._atualizar_status(
                f"Etapa {idx + 1}/{total}: {NOMES_ETAPAS[idx]}...",
                idx / total,
            )

            self._log(f"\n{'─' * 72}")
            self._log(f"▶ ETAPA {idx + 1}/{total}: {NOMES_ETAPAS[idx]}")
            self._log(f"  Script: {script_nome}")
            self._log(f"  Descrição: {DESCRICOES_ETAPAS[idx]}")
            self._log(f"  Início: {datetime.now().strftime('%H:%M:%S')}")
            self._log(f"{'─' * 72}")

            inicio_etapa = time.time()
            max_tentativas = 3
            sucesso = False

            for tentativa in range(1, max_tentativas + 1):
                if self.interrompido:
                    break
                    
                if tentativa > 1:
                    self._log(f"\n🔄 Tentativa {tentativa} de {max_tentativas} para a etapa '{NOMES_ETAPAS[idx]}'...")
                    self.after(0, self._marcar_etapa, idx, "executando")
                    
                try:
                    codigo_retorno = self._executar_script(caminho_script, idx)

                    if self.interrompido:
                        break

                    if codigo_retorno == 0:
                        sucesso = True
                        break
                    else:
                        self._log(f"\n❌ FALHA na tentativa {tentativa} (código: {codigo_retorno})")

                except Exception as e:
                    if self.processo_atual and self.processo_atual.poll() is None:
                        self.processo_atual.terminate()
                        self.processo_atual.wait()
                    self._log(f"\n❌ ERRO INESPERADO na tentativa {tentativa}: {type(e).__name__}: {e}")

            duracao = time.time() - inicio_etapa

            if self.interrompido:
                self.after(0, self._marcar_etapa, idx, "cancelado", duracao)
                self._log(f"\n⛔ Etapa cancelada após {formatar_duracao(duracao)}")
                break

            if not sucesso:
                self.after(0, self._marcar_etapa, idx, "falha", duracao)
                self._log(f"\n❌ FALHA DEFINITIVA na etapa '{NOMES_ETAPAS[idx]}' após {max_tentativas} tentativas.")
                self._log(f"   Duração total: {formatar_duracao(duracao)}")
                self._finalizar_processo()
                return

            self.after(0, self._marcar_etapa, idx, "concluido", duracao)
            self._log(f"\n✅ Etapa concluída em {formatar_duracao(duracao)}")

        duracao_total = time.time() - inicio_total

        if self.interrompido:
            self._atualizar_status(
                f"Interrompido pelo usuário — {formatar_duracao(duracao_total)}",
                max(0.0, idx / total) if idx else 0.0,
            )
            self._log(f"\n⚠️ Pipeline interrompido. Tempo total: {formatar_duracao(duracao_total)}")
        else:
            self._atualizar_status(
                f"✅ Concluído com sucesso! — {formatar_duracao(duracao_total)}",
                1.0,
            )
            self._log(f"\n{'═' * 72}")
            self._log(f"🎉 PIPELINE CONCLUÍDO COM SUCESSO!")
            self._log(f"   Tempo total: {formatar_duracao(duracao_total)}")
            self._log(f"   Finalizado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
            self._log(f"{'═' * 72}")

        self._finalizar_processo()

    def _executar_script(self, caminho_script, idx):
        """
        Executa um script como subprocesso, lendo stdout/stderr em tempo real.
        Funciona tanto em modo dev quanto empacotado.
        """
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        env["GUI_DATA_INICIO"] = self.entries["Data inicial"].get().strip()
        env["GUI_DATA_FIM"] = self.entries["Data final"].get().strip()
        env["GUI_USER"] = self.entries["Usuário (Intranet)"].get().strip()
        env["GUI_PASS"] = self.entries["Senha"].get().strip()

        if getattr(sys, "frozen", False):
            # Modo empacotado: executa o script via o próprio .exe com --run-script
            cmd = [sys.executable, "--run-script", caminho_script]
        else:
            # Modo desenvolvimento: executa via python normalmente
            cmd = [sys.executable, "-u", caminho_script]
            pythonpath = env.get("PYTHONPATH")
            env["PYTHONPATH"] = os.pathsep.join(
                path for path in (str(APP_ROOT), pythonpath) if path
            )

        self.processo_atual = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            env=env,
            encoding="utf-8",
            errors="replace",
            cwd=str(APP_ROOT),
        )

        # Lê linha a linha em tempo real
        for linha in iter(self.processo_atual.stdout.readline, ""):
            if self.interrompido:
                break
            texto = linha.rstrip()
            if texto:
                self._log(f"  │ {texto}")

        self.processo_atual.stdout.close()
        codigo_retorno = self.processo_atual.wait()
        return codigo_retorno

    def _finalizar_processo(self):
        self.executando = False
        self.processo_atual = None
        self.after(0, self._habilitar_botoes)

    def _habilitar_botoes(self):
        self.btn_iniciar.configure(state="normal")
        self.btn_parar.configure(state="disabled")


# ══════════════════════════════════════════════════════════════════════════════
# Ponto de Entrada — com suporte a --run-script para modo empacotado
# ══════════════════════════════════════════════════════════════════════════════
def main():
    """
    Quando o .exe é chamado com --run-script <caminho>, ele executa o script
    Python diretamente (via runpy) em vez de abrir a GUI.
    Isso permite que o exe empacotado funcione como "interpretador" dos
    scripts incluídos, garantindo saída em tempo real via subprocess.
    """
    if len(sys.argv) >= 3 and sys.argv[1] == "--run-script":
        import runpy
        import io
        # Garante stdout unbuffered para tempo real
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")

        script_path = sys.argv[2]
        try:
            runpy.run_path(script_path, run_name="__main__")
        except SystemExit as e:
            sys.exit(e.code if isinstance(e.code, int) else 1)
        except Exception as e:
            print(f"❌ {type(e).__name__}: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc()
            sys.exit(1)
        sys.exit(0)

    # Modo normal: abre a GUI
    app = AppExecutor()
    app.mainloop()


if __name__ == "__main__":
    main()
