import pandas as pd
import unicodedata
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import List, Dict
import os
import threading

class ExcelKeywordSearcherGUI:
    def __init__(self):
        """
        Inicializa a interface gráfica do buscador
        """
        self.df = None
        self.arquivo_path = None
        self.resultados = None
        
        # Cria a janela principal
        self.root = tk.Tk()
        self.root.title("🔍 Buscador de Palavras-chave em Excel")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # Configura o estilo
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.criar_interface()
        
    def normalizar_texto(self, texto: str) -> str:
        """
        Remove acentos, converte para minúsculas e normaliza o texto
        """
        if not isinstance(texto, str):
            texto = str(texto)
        
        texto_sem_acento = unicodedata.normalize('NFD', texto)
        texto_sem_acento = ''.join(char for char in texto_sem_acento 
                                  if unicodedata.category(char) != 'Mn')
        
        return texto_sem_acento.lower()
    
    def criar_interface(self):
        """
        Cria todos os elementos da interface gráfica
        """
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configura redimensionamento
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Título
        titulo = ttk.Label(main_frame, text="🔍 Buscador de Palavras-chave em Excel", 
                          font=('Arial', 16, 'bold'))
        titulo.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Seção 1: Seleção de arquivo
        arquivo_frame = ttk.LabelFrame(main_frame, text="📂 Arquivo Excel", padding="10")
        arquivo_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        arquivo_frame.columnconfigure(1, weight=1)
        
        ttk.Label(arquivo_frame, text="Arquivo:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        
        self.arquivo_var = tk.StringVar()
        self.arquivo_entry = ttk.Entry(arquivo_frame, textvariable=self.arquivo_var, 
                                      state='readonly', width=50)
        self.arquivo_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        
        self.btn_selecionar = ttk.Button(arquivo_frame, text="Selecionar Arquivo", 
                                        command=self.selecionar_arquivo)
        self.btn_selecionar.grid(row=0, column=2)
        
        # Info do arquivo
        self.info_arquivo = ttk.Label(arquivo_frame, text="Nenhum arquivo selecionado", 
                                     foreground='gray')
        self.info_arquivo.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
        
        # Seção 2: Seleção de aba
        aba_frame = ttk.LabelFrame(main_frame, text="📋 Aba do Excel", padding="10")
        aba_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        aba_frame.columnconfigure(1, weight=1)
        
        ttk.Label(aba_frame, text="Aba:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        
        self.aba_var = tk.StringVar()
        self.aba_combo = ttk.Combobox(aba_frame, textvariable=self.aba_var, 
                                     state='readonly', width=30)
        self.aba_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        
        # Seção 3: Palavras-chave
        palavras_frame = ttk.LabelFrame(main_frame, text="🔍 Palavras-chave para Buscar", padding="10")
        palavras_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        palavras_frame.columnconfigure(0, weight=1)
        
        ttk.Label(palavras_frame, text="Digite as palavras separadas por vírgula:").grid(row=0, column=0, sticky=tk.W)
        
        self.palavras_var = tk.StringVar()
        self.palavras_entry = ttk.Entry(palavras_frame, textvariable=self.palavras_var, 
                                       font=('Arial', 11), width=60)
        self.palavras_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Exemplo
        exemplo_label = ttk.Label(palavras_frame, 
                                 text="Exemplo: Responsável, gerente, coordenador", 
                                 foreground='gray', font=('Arial', 9))
        exemplo_label.grid(row=2, column=0, sticky=tk.W, pady=(2, 0))
        
        # Seção 4: Opções avançadas
        opcoes_frame = ttk.LabelFrame(main_frame, text="⚙️ Opções Avançadas", padding="10")
        opcoes_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        opcoes_frame.columnconfigure(1, weight=1)
        
        # Checkbox para colunas específicas
        self.usar_colunas_especificas = tk.BooleanVar()
        checkbox_colunas = ttk.Checkbutton(opcoes_frame, 
                                          text="Buscar apenas em colunas específicas", 
                                          variable=self.usar_colunas_especificas,
                                          command=self.toggle_colunas_especificas)
        checkbox_colunas.grid(row=0, column=0, columnspan=2, sticky=tk.W)
        
        # Entry para colunas específicas
        ttk.Label(opcoes_frame, text="Colunas:").grid(row=1, column=0, sticky=tk.W, padx=(20, 5))
        
        self.colunas_var = tk.StringVar()
        self.colunas_entry = ttk.Entry(opcoes_frame, textvariable=self.colunas_var, 
                                      state='disabled', width=40)
        self.colunas_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        
        # Botão para mostrar colunas disponíveis
        self.btn_mostrar_colunas = ttk.Button(opcoes_frame, text="Ver Colunas", 
                                             command=self.mostrar_colunas, state='disabled')
        self.btn_mostrar_colunas.grid(row=1, column=2)
        
        # Seção 5: Botões de ação
        botoes_frame = ttk.Frame(main_frame)
        botoes_frame.grid(row=5, column=0, columnspan=3, pady=(10, 0))
        
        self.btn_buscar = ttk.Button(botoes_frame, text="🔍 Buscar", 
                                    command=self.executar_busca, 
                                    style='Accent.TButton')
        self.btn_buscar.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_limpar = ttk.Button(botoes_frame, text="🗑️ Limpar", 
                                    command=self.limpar_campos)
        self.btn_limpar.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_salvar = ttk.Button(botoes_frame, text="💾 Salvar Resultados", 
                                    command=self.salvar_resultados, state='disabled')
        self.btn_salvar.pack(side=tk.LEFT)
        
        # Seção 6: Resultados
        resultados_frame = ttk.LabelFrame(main_frame, text="📊 Resultados", padding="10")
        resultados_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        resultados_frame.columnconfigure(0, weight=1)
        resultados_frame.rowconfigure(0, weight=1)
        
        # Área de texto para resultados com scroll
        self.resultado_text = scrolledtext.ScrolledText(resultados_frame, 
                                                       height=15, width=80, 
                                                       font=('Consolas', 10),
                                                       wrap=tk.WORD)
        self.resultado_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Barra de progresso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Configura redimensionamento da seção de resultados
        main_frame.rowconfigure(6, weight=1)
        
    def selecionar_arquivo(self):
        """
        Abre diálogo para selecionar arquivo Excel
        """
        arquivo = filedialog.askopenfilename(
            title="Selecionar arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            self.arquivo_var.set(arquivo)
            self.arquivo_path = arquivo
            self.carregar_arquivo_info()
    
    def carregar_arquivo_info(self):
        """
        Carrega informações do arquivo e popula combobox de abas
        """
        try:
            # Carrega o arquivo para obter informações
            excel_file = pd.ExcelFile(self.arquivo_path)
            
            # Popula combobox com nomes das abas
            self.aba_combo['values'] = excel_file.sheet_names
            self.aba_combo.set(excel_file.sheet_names[0])  # Seleciona primeira aba
            
            # Carrega a primeira aba para mostrar informações
            self.df = pd.read_excel(self.arquivo_path, sheet_name=excel_file.sheet_names[0])
            
            # Atualiza info do arquivo
            info_texto = (f"✅ Arquivo carregado: {len(excel_file.sheet_names)} aba(s), "
                         f"{self.df.shape[0]} linhas x {self.df.shape[1]} colunas")
            self.info_arquivo.config(text=info_texto, foreground='green')
            
            # Bind para atualizar quando trocar de aba
            self.aba_combo.bind('<<ComboboxSelected>>', self.trocar_aba)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar arquivo:\n{str(e)}")
            self.info_arquivo.config(text="❌ Erro ao carregar arquivo", foreground='red')
    
    def trocar_aba(self, event=None):
        """
        Carrega nova aba quando usuário seleciona diferente
        """
        try:
            nome_aba = self.aba_var.get()
            self.df = pd.read_excel(self.arquivo_path, sheet_name=nome_aba)
            
            info_texto = (f"✅ Aba '{nome_aba}': "
                         f"{self.df.shape[0]} linhas x {self.df.shape[1]} colunas")
            self.info_arquivo.config(text=info_texto, foreground='green')
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar aba:\n{str(e)}")
    
    def toggle_colunas_especificas(self):
        """
        Habilita/desabilita entrada de colunas específicas
        """
        if self.usar_colunas_especificas.get():
            self.colunas_entry.config(state='normal')
            self.btn_mostrar_colunas.config(state='normal')
        else:
            self.colunas_entry.config(state='disabled')
            self.btn_mostrar_colunas.config(state='disabled')
            self.colunas_var.set("")
    
    def mostrar_colunas(self):
        """
        Mostra janela com colunas disponíveis
        """
        if self.df is None:
            messagebox.showwarning("Aviso", "Carregue um arquivo Excel primeiro!")
            return
        
        # Cria janela popup
        popup = tk.Toplevel(self.root)
        popup.title("Colunas Disponíveis")
        popup.geometry("400x500")
        popup.resizable(True, True)
        
        # Frame principal
        frame = ttk.Frame(popup, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Colunas disponíveis no arquivo:", 
                 font=('Arial', 12, 'bold')).pack(anchor=tk.W)
        
        # Lista de colunas
        lista_frame = ttk.Frame(frame)
        lista_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(lista_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Listbox
        listbox = tk.Listbox(lista_frame, yscrollcommand=scrollbar.set, 
                            font=('Consolas', 10))
        
        for coluna in self.df.columns:
            listbox.insert(tk.END, coluna)
        
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Botões
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        def copiar_selecionadas():
            selecionadas = [listbox.get(i) for i in listbox.curselection()]
            if selecionadas:
                self.colunas_var.set(", ".join(selecionadas))
                popup.destroy()
        
        def copiar_todas():
            self.colunas_var.set(", ".join(self.df.columns.tolist()))
            popup.destroy()
        
        ttk.Button(btn_frame, text="Copiar Selecionadas", 
                  command=copiar_selecionadas).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Copiar Todas", 
                  command=copiar_todas).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Fechar", 
                  command=popup.destroy).pack(side=tk.RIGHT)
        
        # Instruções
        ttk.Label(frame, text="Dica: Ctrl+clique para selecionar múltiplas colunas", 
                 foreground='gray').pack(anchor=tk.W, pady=(5, 0))
    
    def executar_busca(self):
        """
        Executa a busca em thread separada para não travar a interface
        """
        # Validações
        if not self.arquivo_path:
            messagebox.showwarning("Aviso", "Selecione um arquivo Excel primeiro!")
            return
        
        palavras_texto = self.palavras_var.get().strip()
        if not palavras_texto:
            messagebox.showwarning("Aviso", "Digite pelo menos uma palavra-chave!")
            return
        
        # Desabilita botão e mostra progresso
        self.btn_buscar.config(state='disabled')
        self.progress.start()
        self.resultado_text.delete(1.0, tk.END)
        self.resultado_text.insert(tk.END, "🔍 Executando busca...\n")
        
        # Executa busca em thread separada
        thread = threading.Thread(target=self._buscar_thread)
        thread.daemon = True
        thread.start()
    
    def _buscar_thread(self):
        """
        Thread que executa a busca propriamente dita
        """
        try:
            # Processa palavras-chave
            palavras_chave = [palavra.strip() for palavra in self.palavras_var.get().split(',')]
            palavras_chave = [p for p in palavras_chave if p]  # Remove vazias
            
            # Processa colunas específicas
            colunas_especificas = None
            if self.usar_colunas_especificas.get():
                colunas_texto = self.colunas_var.get().strip()
                if colunas_texto:
                    colunas_especificas = [col.strip() for col in colunas_texto.split(',')]
                    colunas_especificas = [col for col in colunas_especificas if col in self.df.columns]
            
            # Executa busca
            self.resultados = self.buscar_palavras_chave(palavras_chave, colunas_especificas)
            
            # Atualiza interface na thread principal
            self.root.after(0, self._finalizar_busca)
            
        except Exception as e:
            self.root.after(0, lambda: self._erro_busca(str(e)))
    
    def _finalizar_busca(self):
        """
        Finaliza busca e atualiza interface
        """
        self.progress.stop()
        self.btn_buscar.config(state='normal')
        
        # Exibe resultados
        self.exibir_resultados()
        
        # Habilita botão salvar se houver resultados
        if self.resultados['total_ocorrencias'] > 0:
            self.btn_salvar.config(state='normal')
    
    def _erro_busca(self, erro):
        """
        Trata erros na busca
        """
        self.progress.stop()
        self.btn_buscar.config(state='normal')
        messagebox.showerror("Erro na Busca", f"Erro durante a busca:\n{erro}")
        
        self.resultado_text.delete(1.0, tk.END)
        self.resultado_text.insert(tk.END, f"❌ Erro na busca: {erro}")
    
    def buscar_palavras_chave(self, palavras_chave: List[str], 
                             colunas_especificas: List[str] = None) -> Dict:
        """
        Busca palavras-chave no DataFrame (mesma lógica do script original)
        """
        palavras_normalizadas = [self.normalizar_texto(palavra) for palavra in palavras_chave]
        
        resultados = {
            'palavras_encontradas': {},
            'total_ocorrencias': 0,
            'resumo': {}
        }
        
        # Define colunas para buscar
        if colunas_especificas:
            colunas_busca = [col for col in colunas_especificas if col in self.df.columns]
        else:
            colunas_busca = self.df.columns.tolist()
        
        # Para cada palavra-chave
        for idx, palavra_original in enumerate(palavras_chave):
            palavra_normalizada = palavras_normalizadas[idx]
            resultados['palavras_encontradas'][palavra_original] = []
            
            # Para cada coluna
            for coluna in colunas_busca:
                # Para cada linha
                for linha_idx, valor_celula in enumerate(self.df[coluna]):
                    if pd.isna(valor_celula):
                        continue
                    
                    valor_normalizado = self.normalizar_texto(str(valor_celula))
                    
                    # Verifica se a palavra está presente
                    if palavra_normalizada in valor_normalizado:
                        # Pega todos os dados da linha onde encontrou a palavra
                        linha_completa = {}
                        for col_name in self.df.columns:
                            linha_completa[col_name] = self.df.iloc[linha_idx][col_name]
                        
                        resultado_item = {
                            'linha': linha_idx + 2,  # +2 porque Excel começa em 1 e tem header
                            'coluna': coluna,
                            'valor_original': valor_celula,
                            'posicao_encontrada': valor_normalizado.find(palavra_normalizada),
                            'linha_completa': linha_completa  # Adiciona toda a linha
                        }
                        
                        resultados['palavras_encontradas'][palavra_original].append(resultado_item)
                        resultados['total_ocorrencias'] += 1
        
        # Gera resumo
        for palavra, ocorrencias in resultados['palavras_encontradas'].items():
            resultados['resumo'][palavra] = len(ocorrencias)
        
        return resultados
    
    def exibir_resultados(self):
        """
        Exibe resultados na área de texto
        """
        self.resultado_text.delete(1.0, tk.END)
        
        if self.resultados['total_ocorrencias'] == 0:
            self.resultado_text.insert(tk.END, "❌ Nenhuma palavra-chave foi encontrada!\n")
            self.resultado_text.insert(tk.END, "\nDicas:\n")
            self.resultado_text.insert(tk.END, "• Verifique se as palavras estão corretas\n")
            self.resultado_text.insert(tk.END, "• Tente palavras mais simples\n")
            self.resultado_text.insert(tk.END, "• Verifique se está na aba correta\n")
            return
        
        # Cabeçalho
        self.resultado_text.insert(tk.END, "="*60 + "\n")
        self.resultado_text.insert(tk.END, "📊 RESULTADOS DA BUSCA\n")
        self.resultado_text.insert(tk.END, "="*60 + "\n\n")
        
        self.resultado_text.insert(tk.END, f"✅ Total de ocorrências: {self.resultados['total_ocorrencias']}\n\n")
        
        # Resumo
        self.resultado_text.insert(tk.END, "📋 Resumo por palavra:\n")
        for palavra, quantidade in self.resultados['resumo'].items():
            if quantidade > 0:
                self.resultado_text.insert(tk.END, f"  • '{palavra}': {quantidade} ocorrência(s)\n")
        
        self.resultado_text.insert(tk.END, "\n📍 Detalhes das ocorrências:\n")
        
        # Detalhes
        for palavra, ocorrencias in self.resultados['palavras_encontradas'].items():
            if ocorrencias:
                self.resultado_text.insert(tk.END, f"\n🔍 Palavra: '{palavra}'\n")
                self.resultado_text.insert(tk.END, "-" * 40 + "\n")
                
                for i, item in enumerate(ocorrencias, 1):
                    self.resultado_text.insert(tk.END, f"  {i}. Linha {item['linha']}, Coluna '{item['coluna']}'\n")
                    # Limita tamanho do conteúdo mostrado
                    conteudo = str(item['valor_original'])
                    if len(conteudo) > 100:
                        conteudo = conteudo[:100] + "..."
                    self.resultado_text.insert(tk.END, f"     Conteúdo: {conteudo}\n\n")
    
    def salvar_resultados(self):
        """
        Salva resultados em arquivo Excel com todas as colunas do relatório original
        """
        if not self.resultados or self.resultados['total_ocorrencias'] == 0:
            messagebox.showwarning("Aviso", "Não há resultados para salvar!")
            return
        
        # Diálogo para salvar arquivo
        arquivo_saida = filedialog.asksaveasfilename(
            title="Salvar resultados",
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        
        if not arquivo_saida:
            return
        
        try:
            # Prepara dados com TODAS as colunas do relatório original
            dados_para_salvar = []
            
            for palavra, ocorrencias in self.resultados['palavras_encontradas'].items():
                for item in ocorrencias:
                    # Cria registro base com informações da busca
                    registro = {
                        'PALAVRA_BUSCADA': palavra,
                        'LINHA_ENCONTRADA': item['linha'],
                        'COLUNA_ENCONTRADA': item['coluna'],
                        'CONTEUDO_ENCONTRADO': item['valor_original']
                    }
                    
                    # Adiciona TODAS as colunas da linha original
                    for nome_coluna, valor_coluna in item['linha_completa'].items():
                        # Evita duplicar a coluna onde encontrou (já está em CONTEUDO_ENCONTRADO)
                        nome_coluna_limpo = f"ORIGINAL_{nome_coluna}"
                        registro[nome_coluna_limpo] = valor_coluna
                    
                    dados_para_salvar.append(registro)
            
            # Cria DataFrame e salva
            df_resultados = pd.DataFrame(dados_para_salvar)
            
            # Reordena colunas: informações da busca primeiro, depois dados originais
            colunas_busca = ['PALAVRA_BUSCADA', 'LINHA_ENCONTRADA', 'COLUNA_ENCONTRADA', 'CONTEUDO_ENCONTRADO']
            colunas_originais = [col for col in df_resultados.columns if col.startswith('ORIGINAL_')]
            colunas_ordenadas = colunas_busca + sorted(colunas_originais)
            
            df_resultados = df_resultados[colunas_ordenadas]
            df_resultados.to_excel(arquivo_saida, index=False)
            
            # Mensagem de sucesso detalhada
            total_colunas = len(df_resultados.columns)
            colunas_originais_count = len(colunas_originais)
            
            mensagem = (f"✅ Resultados salvos com sucesso!\n\n"
                       f"📁 Arquivo: {arquivo_saida}\n"
                       f"📊 Linhas: {len(df_resultados)} resultados\n"
                       f"📋 Colunas: {total_colunas} total ({colunas_originais_count} do relatório original)\n\n"
                       f"🔍 Estrutura do arquivo salvo:\n"
                       f"• Informações da busca (4 colunas)\n"
                       f"• Todas as colunas do relatório original ({colunas_originais_count} colunas)")
            
            messagebox.showinfo("Sucesso", mensagem)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar arquivo:\n{str(e)}")
            print(f"Erro detalhado: {e}")  # Para debug
    
    def limpar_campos(self):
        """
        Limpa todos os campos da interface
        """
        self.arquivo_var.set("")
        self.palavras_var.set("")
        self.colunas_var.set("")
        self.usar_colunas_especificas.set(False)
        self.toggle_colunas_especificas()
        self.resultado_text.delete(1.0, tk.END)
        self.btn_salvar.config(state='disabled')
        self.info_arquivo.config(text="Nenhum arquivo selecionado", foreground='gray')
        self.aba_combo['values'] = []
        self.aba_var.set("")
        self.df = None
        self.arquivo_path = None
        self.resultados = None
    
    def executar(self):
        """
        Inicia a interface gráfica
        """
        self.root.mainloop()

def main():
    """
    Função principal - inicia a aplicação
    """
    try:
        app = ExcelKeywordSearcherGUI()
        app.executar()
    except Exception as e:
        print(f"Erro ao iniciar aplicação: {e}")
        input("Pressione Enter para sair...")

if __name__ == "__main__":
    main()