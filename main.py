import logging
import re
import threading
import tkinter as tk
import unicodedata
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Dict, List

import nltk
import pandas as pd
from nltk.stem import RSLPStemmer
from rapidfuzz import fuzz


class ExcelKeywordSearcherGUI:
    def __init__(self):
        # ---- Estado ----
        self.df: pd.DataFrame | None = None
        self.arquivo_path: str | None = None
        self.resultados: Dict | None = None
        self.norm_cache: dict[str, pd.Series] = {}
        self.stem_cache: dict[str, pd.Series] = {}

        # ---- Logging ----
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s %(levelname)s %(message)s"
        )
        self._ensure_nltk()

        # ---- Stemmer ----
        self.stemmer = RSLPStemmer()

        # ---- UI raiz ----
        self.root = tk.Tk()
        self.root.title("🔍 Buscador de Palavras-chave em Excel")
        self.root.geometry("900x760")
        self.root.resizable(True, True)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        self._build_ui()

    # ===================== Utilidades de texto =====================
    def normalizar_texto(self, texto: str) -> str:
        if not isinstance(texto, str):
            texto = str(texto)
        t = unicodedata.normalize('NFD', texto)
        t = ''.join(ch for ch in t if unicodedata.category(ch) != 'Mn')
        return t.lower()

    def stem_pt(self, s: str) -> str:
        tokens = re.findall(r"\w+", self.normalizar_texto(s))
        return " ".join(self.stemmer.stem(t) for t in tokens)

    # ===================== UI =====================
    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(6, weight=1)

        ttk.Label(main, text="🔍 Buscador de Palavras-chave em Excel",
                  font=('Arial', 16, 'bold')).grid(row=0, column=0, columnspan=3, pady=(0, 12))

        # Arquivo
        lf_arq = ttk.LabelFrame(main, text="📂 Arquivo Excel", padding=10)
        lf_arq.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        lf_arq.columnconfigure(1, weight=1)

        ttk.Label(lf_arq, text="Arquivo:").grid(row=0, column=0, sticky="w")
        self.arquivo_var = tk.StringVar()
        ttk.Entry(lf_arq, textvariable=self.arquivo_var, state="readonly").grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(lf_arq, text="Selecionar", command=self.selecionar_arquivo).grid(row=0, column=2)
        self.info_arquivo = ttk.Label(lf_arq, text="Nenhum arquivo selecionado", foreground="gray")
        self.info_arquivo.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))

        # Aba
        lf_aba = ttk.LabelFrame(main, text="📋 Aba do Excel", padding=10)
        lf_aba.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        lf_aba.columnconfigure(1, weight=1)

        ttk.Label(lf_aba, text="Aba:").grid(row=0, column=0, sticky="w")
        self.aba_var = tk.StringVar()
        self.aba_combo = ttk.Combobox(lf_aba, textvariable=self.aba_var, state="readonly")
        self.aba_combo.grid(row=0, column=1, sticky="ew", padx=6)
        self.aba_combo.bind('<<ComboboxSelected>>', self.trocar_aba)

        # Palavras
        lf_pal = ttk.LabelFrame(main, text="🔎 Palavras-chave", padding=10)
        lf_pal.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        lf_pal.columnconfigure(0, weight=1)

        ttk.Label(lf_pal, text="Digite separadas por vírgula:").grid(row=0, column=0, sticky="w")
        self.palavras_var = tk.StringVar()
        ttk.Entry(lf_pal, textvariable=self.palavras_var).grid(row=1, column=0, sticky="ew", pady=4)
        ttk.Label(lf_pal, text="Ex.: desarquivamento, arquivamento, arquivar", foreground="gray").grid(row=2, column=0, sticky="w")

        # Opções
        lf_ops = ttk.LabelFrame(main, text="⚙️ Opções", padding=10)
        lf_ops.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        for i in range(6):
            lf_ops.columnconfigure(i, weight=1)

        # Modo
        ttk.Label(lf_ops, text="Modo de busca:").grid(row=0, column=0, sticky="w")
        self.modo_busca = tk.StringVar(value="similar")  # exato | regex | fuzzy | stem
        ttk.Combobox(lf_ops, textvariable=self.modo_busca, state="readonly",
                     values=["exato", "padrão", "similaridade", "radical"]).grid(row=0, column=1, sticky="ew", padx=(6, 12))

        # Limiar fuzzy
        ttk.Label(lf_ops, text="% de Similaridade:").grid(row=0, column=2, sticky="w")
        self.limiar_fuzzy = tk.IntVar(value=80)
        
        self.lbl_limiar = ttk.Label(lf_ops, text=str(self.limiar_fuzzy.get()))
        self.lbl_limiar.grid(row=0, column=4, sticky="w")
        
        self.slider = ttk.Scale(
            lf_ops, from_=60, to=95, orient="horizontal",
            command=self._on_slider_change
        )
        
        self.slider = ttk.Scale(lf_ops, from_=60, to=95, orient="horizontal",
                                command=lambda v: self.lbl_limiar.config(text=str(int(float(v)))))
        self.slider.grid(row=0, column=3, sticky="ew", padx=6)
        self.slider.set(self.limiar_fuzzy.get())

        # Colunas específicas
        self.usar_colunas_especificas = tk.BooleanVar()
        ttk.Checkbutton(lf_ops, text="Buscar apenas em colunas específicas",
                        variable=self.usar_colunas_especificas,
                        command=self.toggle_colunas_especificas).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        ttk.Label(lf_ops, text="Colunas:").grid(row=1, column=2, sticky="w", pady=(6, 0))
        self.colunas_var = tk.StringVar()
        self.colunas_entry = ttk.Entry(lf_ops, textvariable=self.colunas_var, state='disabled')
        self.colunas_entry.grid(row=1, column=3, columnspan=2, sticky="ew", padx=6, pady=(6, 0))
        self.btn_mostrar_colunas = ttk.Button(lf_ops, text="Ver Colunas", command=self.mostrar_colunas, state='disabled')
        self.btn_mostrar_colunas.grid(row=1, column=5, sticky="e", pady=(6, 0))

        # Ações
        acts = ttk.Frame(main)
        acts.grid(row=5, column=0, columnspan=3, pady=8)
        ttk.Button(acts, text="🔍 Buscar", command=self.executar_busca, style='Accent.TButton').pack(side="left", padx=5)
        ttk.Button(acts, text="🗑️ Limpar", command=self.limpar_campos).pack(side="left", padx=5)
        self.btn_salvar = ttk.Button(acts, text="💾 Salvar Resultados", command=self.salvar_resultados, state='disabled')
        self.btn_salvar.pack(side="left", padx=5)

        # Resultados
        lf_res = ttk.LabelFrame(main, text="📊 Resultados", padding=10)
        lf_res.grid(row=6, column=0, columnspan=3, sticky="nsew")
        lf_res.columnconfigure(0, weight=1)
        lf_res.rowconfigure(0, weight=1)

        self.resultado_text = scrolledtext.ScrolledText(lf_res, height=18, font=('Consolas', 10), wrap="word")
        self.resultado_text.grid(row=0, column=0, sticky="nsew")

        self.progress = ttk.Progressbar(main, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        
    def _on_slider_change(self, v):
        try:
            val = int(float(v))
        except Exception:
            val = self.limiar_fuzzy.get()
        self.limiar_fuzzy.set(val)
        self.lbl_limiar.config(text=str(val))


    # ===================== Arquivo/Aba =====================
    def selecionar_arquivo(self):
        arq = filedialog.askopenfilename(
            title="Selecionar arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if not arq:
            return
        self.arquivo_var.set(arq)
        self.arquivo_path = arq
        self._carregar_arquivo_info()

    def _carregar_arquivo_info(self):
        try:
            excel_file = pd.ExcelFile(self.arquivo_path)
            self.aba_combo['values'] = excel_file.sheet_names
            self.aba_combo.set(excel_file.sheet_names[0])
            self.df = pd.read_excel(self.arquivo_path, sheet_name=excel_file.sheet_names[0], engine="openpyxl")
            self._clear_caches()
            self.info_arquivo.config(
                text=f"✅ {len(excel_file.sheet_names)} aba(s) | {self.df.shape[0]} linhas × {self.df.shape[1]} colunas",
                foreground="green"
            )
        except Exception as e:
            logging.exception("Erro ao carregar arquivo")
            messagebox.showerror("Erro", f"Erro ao carregar arquivo:\n{e}")
            self.info_arquivo.config(text="❌ Erro ao carregar arquivo", foreground="red")

    def trocar_aba(self, _evt=None):
        if not self.arquivo_path:
            return
        try:
            nome_aba = self.aba_var.get()
            self.df = pd.read_excel(self.arquivo_path, sheet_name=nome_aba, engine="openpyxl")
            self._clear_caches()
            self.info_arquivo.config(
                text=f"✅ Aba '{nome_aba}': {self.df.shape[0]} linhas × {self.df.shape[1]} colunas",
                foreground="green"
            )
        except Exception as e:
            logging.exception("Erro ao carregar aba")
            messagebox.showerror("Erro", f"Erro ao carregar aba:\n{e}")

    # ===================== Colunas específicas =====================
    def toggle_colunas_especificas(self):
        state = 'normal' if self.usar_colunas_especificas.get() else 'disabled'
        self.colunas_entry.config(state=state)
        self.btn_mostrar_colunas.config(state=state)
        if state == 'disabled':
            self.colunas_var.set("")

    def mostrar_colunas(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Carregue um arquivo Excel primeiro.")
            return

        popup = tk.Toplevel(self.root)
        popup.title("Colunas Disponíveis")
        popup.geometry("420x520")
        popup.resizable(True, True)

        frame = ttk.Frame(popup, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Colunas disponíveis:", font=('Arial', 12, 'bold')).pack(anchor="w")

        lista_frame = ttk.Frame(frame)
        lista_frame.pack(fill="both", expand=True, pady=(10, 0))

        scrollbar = ttk.Scrollbar(lista_frame)
        scrollbar.pack(side="right", fill="y")

        listbox = tk.Listbox(lista_frame, yscrollcommand=scrollbar.set, font=('Consolas', 10), selectmode="extended")
        for c in self.df.columns:
            listbox.insert("end", c)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)

        btns = ttk.Frame(frame)
        btns.pack(fill="x", pady=(10, 0))

        def copiar_sel():
            sel = [listbox.get(i) for i in listbox.curselection()]
            if sel:
                self.colunas_var.set(", ".join(sel))
                popup.destroy()

        def copiar_todas():
            self.colunas_var.set(", ".join(self.df.columns.tolist()))
            popup.destroy()

        ttk.Button(btns, text="Copiar Selecionadas", command=copiar_sel).pack(side="left", padx=4)
        ttk.Button(btns, text="Copiar Todas", command=copiar_todas).pack(side="left", padx=4)
        ttk.Button(btns, text="Fechar", command=popup.destroy).pack(side="right")

        ttk.Label(frame, text="Dica: Ctrl/Cmd + clique para múltiplas seleções", foreground="gray").pack(anchor="w", pady=6)

    # ===================== Execução de busca =====================
    def executar_busca(self):
        if not self.arquivo_path:
            messagebox.showwarning("Aviso", "Selecione um arquivo Excel.")
            return
        palavras_texto = self.palavras_var.get().strip()
        if not palavras_texto:
            messagebox.showwarning("Aviso", "Digite pelo menos uma palavra-chave.")
            return

        # sincroniza slider -> variável
        try:
            self.limiar_fuzzy.set(int(float(self.slider.get())))
        except Exception:
            pass

        self.btn_salvar.config(state='disabled')
        self.btn_buscar_state(False)
        self.progress.start()
        self.resultado_text.delete(1.0, "end")
        self.resultado_text.insert("end", "🔎 Executando busca...\n")

        t = threading.Thread(target=self._buscar_thread, daemon=True)
        t.start()

    def _buscar_thread(self):
        try:
            palavras = [p.strip() for p in self.palavras_var.get().split(",") if p.strip()]
            cols = None
            if self.usar_colunas_especificas.get():
                cols = [c.strip() for c in self.colunas_var.get().split(",") if c.strip()]
            self.resultados = self.buscar_palavras_chave(palavras, cols)
            self.root.after(0, self._finalizar_busca)
        except Exception:
            logging.exception("Erro durante a busca")
            self.root.after(0, lambda: self._erro_busca(str(e)))

    def _finalizar_busca(self):
        self.progress.stop()
        self.btn_buscar_state(True)
        self.exibir_resultados()
        if self.resultados and self.resultados.get('total_ocorrencias', 0) > 0:
            self.btn_salvar.config(state='normal')

    def _erro_busca(self, erro: str):
        self.progress.stop()
        self.btn_buscar_state(True)
        messagebox.showerror("Erro na Busca", f"Ocorreu um erro:\n{erro}")
        self.resultado_text.delete(1.0, "end")
        self.resultado_text.insert("end", f"❌ Erro na busca: {erro}")

    def btn_buscar_state(self, enable: bool):
        # protege caso o botão ainda não exista em alguns temas
        for w in self.root.winfo_children():
            pass
        try:
            # encontra pelo texto
            for child in self.root.winfo_children():
                pass
        except Exception:
            pass

    # ===================== Núcleo de busca =====================
    def buscar_palavras_chave(self, palavras_chave: List[str], colunas_especificas: List[str] | None = None) -> Dict:
        if self.df is None or self.df.empty:
            return {'palavras_encontradas': {}, 'total_ocorrencias': 0, 'resumo': {}}

        modo = self.modo_busca.get()
        limiar = self.limiar_fuzzy.get()
        logging.info("Iniciando busca | modo=%s limiar=%s", modo, limiar)

        resultados = {'palavras_encontradas': {}, 'total_ocorrencias': 0, 'resumo': {}}

        # colunas alvo
        colunas_busca = [c for c in (colunas_especificas or self.df.columns.tolist()) if c in self.df.columns]

        # caches por coluna
        self._prepare_caches(colunas_busca, need_stem=(modo == "stem"))

        # prepara consultas
        consultas: list[tuple[str, str]] = []
        for p in palavras_chave:
            if not p:
                continue
            proc = self.stem_pt(p) if modo == "stem" else self.normalizar_texto(p)
            consultas.append((p, proc))

        for palavra_original, alvo_proc in consultas:
            resultados['palavras_encontradas'][palavra_original] = []
            for c in colunas_busca:
                base = self.stem_cache[c] if modo == "stem" else self.norm_cache[c]

                if modo == "exato":
                    mask = base.str.contains(re.escape(alvo_proc), na=False)
                    idxs = base[mask].index

                elif modo == "regex":
                    try:
                        mask = base.str.contains(alvo_proc, na=False)
                    except re.error:
                        mask = pd.Series(False, index=base.index)
                    idxs = base[mask].index

                elif modo == "fuzzy":
                    # pré-filtro barato por trigram
                    if len(alvo_proc) >= 3:
                        trig = re.escape(alvo_proc[:3])
                        pre = base.str.contains(trig, na=False)
                        cand_idx = base[pre].index
                    else:
                        cand_idx = base.index

                    idxs = []
                    for i in cand_idx:
                        if fuzz.partial_ratio(alvo_proc, base.at[i]) >= limiar:
                            idxs.append(i)

                elif modo == "stem":
                    mask = base.str.contains(re.escape(alvo_proc), na=False)
                    idxs = base[mask].index

                else:
                    idxs = []

                for i in idxs:
                    valor_original = self.df.at[i, c]
                    linha_completa = self.df.loc[i, :].to_dict()
                    pos = str(valor_original).lower().find(palavra_original.lower())
                    resultados['palavras_encontradas'][palavra_original].append({
                        'linha': i + 2,  # Excel header + index base 1
                        'coluna': c,
                        'valor_original': valor_original,
                        'posicao_encontrada': pos,
                        'linha_completa': linha_completa
                    })
                    resultados['total_ocorrencias'] += 1

        for palavra, ocorr in resultados['palavras_encontradas'].items():
            resultados['resumo'][palavra] = len(ocorr)

        logging.info("Busca finalizada | ocorrências=%d", resultados['total_ocorrencias'])
        return resultados

    def _prepare_caches(self, cols: list[str], need_stem: bool):
        for c in cols:
            if c not in self.norm_cache:
                serie = self.df[c].astype(str)
                self.norm_cache[c] = serie.map(self.normalizar_texto)
            if need_stem and c not in self.stem_cache:
                serie = self.df[c].astype(str)
                self.stem_cache[c] = serie.map(self.stem_pt)

    def _clear_caches(self):
        self.norm_cache.clear()
        self.stem_cache.clear()

    # ===================== Exibir/Salvar =====================
    def exibir_resultados(self):
        self.resultado_text.delete(1.0, "end")
        if not self.resultados or self.resultados.get('total_ocorrencias', 0) == 0:
            self.resultado_text.insert("end", "❌ Nenhuma palavra-chave foi encontrada.\n")
            self.resultado_text.insert("end", "\nTente:\n• Palavras mais simples\n• Modo fuzzy ou stem\n• Ajustar o limiar\n")
            return

        self.resultado_text.insert("end", "=" * 64 + "\n")
        self.resultado_text.insert("end", "📊 RESULTADOS DA BUSCA\n")
        self.resultado_text.insert("end", "=" * 64 + "\n\n")
        self.resultado_text.insert("end", f"✅ Total de ocorrências: {self.resultados['total_ocorrencias']}\n\n")

        self.resultado_text.insert("end", "📋 Resumo por palavra:\n")
        for palavra, qtd in self.resultados['resumo'].items():
            if qtd > 0:
                self.resultado_text.insert("end", f"  • '{palavra}': {qtd}\n")

        self.resultado_text.insert("end", "\n📍 Detalhes:\n")
        for palavra, ocorr in self.resultados['palavras_encontradas'].items():
            if not ocorr:
                continue
            self.resultado_text.insert("end", f"\n🔎 Palavra: '{palavra}'\n" + "-" * 40 + "\n")
            for i, item in enumerate(ocorr, 1):
                conteudo = str(item['valor_original'])
                if len(conteudo) > 140:
                    conteudo = conteudo[:140] + "..."
                self.resultado_text.insert("end", f"  {i}. Linha {item['linha']}, Coluna '{item['coluna']}'\n")
                self.resultado_text.insert("end", f"     Conteúdo: {conteudo}\n")

    def salvar_resultados(self):
        if not self.resultados or self.resultados.get('total_ocorrencias', 0) == 0:
            messagebox.showwarning("Aviso", "Não há resultados para salvar.")
            return

        saida = filedialog.asksaveasfilename(
            title="Salvar resultados",
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if not saida:
            return

        try:
            registros = []
            for palavra, ocorr in self.resultados['palavras_encontradas'].items():
                for item in ocorr:
                    base = {
                        'PALAVRA_BUSCADA': palavra,
                        'LINHA_ENCONTRADA': item['linha'],
                        'COLUNA_ENCONTRADA': item['coluna'],
                        'CONTEUDO_ENCONTRADO': item['valor_original']
                    }
                    for nome_coluna, valor in item['linha_completa'].items():
                        base[f'ORIGINAL_{nome_coluna}'] = valor
                    registros.append(base)

            df_out = pd.DataFrame(registros)
            cols_first = ['PALAVRA_BUSCADA', 'LINHA_ENCONTRADA', 'COLUNA_ENCONTRADA', 'CONTEUDO_ENCONTRADO']
            cols_orig = sorted([c for c in df_out.columns if c.startswith("ORIGINAL_")])
            df_out = df_out[cols_first + cols_orig]
            df_out.to_excel(saida, index=False)

            messagebox.showinfo(
                "Sucesso",
                f"✅ Resultados salvos\n\nArquivo: {saida}\nLinhas: {len(df_out)}\nColunas totais: {len(df_out.columns)}"
            )
        except Exception as e:
            logging.exception("Erro ao salvar")
            messagebox.showerror("Erro", f"Erro ao salvar arquivo:\n{e}")

    # ===================== Miscelânea =====================
    def limpar_campos(self):
        self.arquivo_var.set("")
        self.palavras_var.set("")
        self.colunas_var.set("")
        self.usar_colunas_especificas.set(False)
        self.toggle_colunas_especificas()
        self.resultado_text.delete(1.0, "end")
        self.btn_salvar.config(state='disabled')
        self.info_arquivo.config(text="Nenhum arquivo selecionado", foreground='gray')
        self.aba_combo['values'] = []
        self.aba_var.set("")
        self.df = None
        self.arquivo_path = None
        self.resultados = None
        self._clear_caches()

    def _ensure_nltk(self):
        try:
            nltk.data.find('stemmers/rslp')
        except LookupError:
            nltk.download('rslp')

    def executar(self):
        self.root.mainloop()


def main():
    try:
        app = ExcelKeywordSearcherGUI()
        app.executar()
    except Exception as e:
        logging.exception("Falha ao iniciar")
        print(f"Erro ao iniciar aplicação: {e}")
        input("Pressione Enter para sair...")


if __name__ == "__main__":
    main()
