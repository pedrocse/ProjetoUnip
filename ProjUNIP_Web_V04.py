import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, simpledialog
import os
import re
import json
import time
from datetime import datetime, timedelta

# Tentar importar python-docx para ler arquivos .docx
try:
    from docx import Document

    DOCX_DISPONIVEL = True
except ImportError:
    DOCX_DISPONIVEL = False
    print("Biblioteca python-docx não encontrada. Instale com: pip install python-docx")


class QuestoesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplicativo de Questões")
        self.root.geometry("1200x750")
        #self.root.state('zoomed')
        #self.root.resizable(True, True)  # Permite redimensionar largura e altura
        self.root.configure(bg='#2c3e50')

        # Variáveis
        self.df = None
        self.questao_atual = 0
        self.questoes_agrupadas = []
        self.alternativa_selecionada = None
        self.resposta_verificada = False

        # Estatísticas
        self.total_acertos = 0
        self.total_erros = 0
        self.questoes_respondidas = set()

        # Temporizador
        self.timer_ativo = False
        self.tempo_inicial = None
        self.tempo_restante = None
        self.timer_id = None
        self.tempo_prova = 60 * 60  # 60 minutos em segundos (padrão)

        # Arquivo de progresso
        self.arquivo_progresso = "progresso_questoes.json"

        # Diretório de teorias
        self.diretorio_teorias = "teorias"

        # Diretório de imagens das questões
        self.diretorio_imagens = "imagens"

        # Criar widgets
        self.criar_widgets()

        # Carregar progresso anterior
        self.carregar_progresso()

    def criar_widgets(self):
        """Cria todos os widgets da interface"""

        # Limpar tudo primeiro
        for widget in self.root.winfo_children():
            widget.destroy()

        # Frame superior - título
        frame_topo = tk.Frame(self.root, bg='#34495e', height=80)
        frame_topo.pack(fill='x')

        titulo = tk.Label(
            frame_topo,
            text="📚 Aplicativo de Questões - TOMOs",
            font=("Arial", 24, "bold"),
            bg='#34495e',
            fg='white'
        )
        titulo.pack(pady=10)

        # Frame de estatísticas e temporizador
        frame_stats = tk.Frame(frame_topo, bg='#34495e')
        frame_stats.pack(fill='x', padx=20, pady=5)

        # Estatísticas
        self.lbl_estatisticas = tk.Label(
            frame_stats,
            text="📊 Acertos: 0 | Erros: 0 | Total: 0",
            font=("Arial", 11, "bold"),
            bg='#34495e',
            fg='#f39c12'
        )
        self.lbl_estatisticas.pack(side='left', padx=10)

        # Temporizador
        self.lbl_temporizador = tk.Label(
            frame_stats,
            text="⏱️ Tempo: 00:00",
            font=("Arial", 11, "bold"),
            bg='#34495e',
            fg='#f39c12'
        )
        self.lbl_temporizador.pack(side='right', padx=10)

        # Frame principal
        self.frame_principal = tk.Frame(self.root, bg='#2c3e50')
        self.frame_principal.pack(fill='both', expand=True, padx=20, pady=20)

        # Frame para carregar arquivo
        self.frame_carregar = tk.Frame(self.frame_principal, bg='#2c3e50')
        self.frame_carregar.pack(fill='both', expand=True)

        lbl_instrucao = tk.Label(
            self.frame_carregar,
            text="📁 Selecione o arquivo Excel com as questões:",
            font=("Arial", 12),
            bg='#2c3e50',
            fg='white'
        )
        lbl_instrucao.pack(pady=20)

        # Entrada para nome do arquivo
        frame_arquivo = tk.Frame(self.frame_carregar, bg='#2c3e50')
        frame_arquivo.pack(pady=10)

        self.entry_arquivo = tk.Entry(frame_arquivo, width=50, font=("Arial", 11))
        self.entry_arquivo.insert(0, "TOMOs com questões.xlsx")
        self.entry_arquivo.pack(side='left', padx=5)

        btn_procurar = tk.Button(
            frame_arquivo,
            text="📂 Procurar",
            command=self.procurar_arquivo,
            bg='#3498db',
            fg='white',
            font=("Arial", 10, "bold"),
            cursor='hand2'
        )
        btn_procurar.pack(side='left', padx=5)

        # Botão carregar
        btn_carregar = tk.Button(
            self.frame_carregar,
            text="🚀 Carregar Planilha",
            command=self.carregar_planilha,
            bg='#27ae60',
            fg='white',
            font=("Arial", 12, "bold"),
            padx=30,
            pady=10,
            cursor='hand2'
        )
        btn_carregar.pack(pady=30)

        # Botões de modo de estudo/prova
        frame_modos = tk.Frame(self.frame_carregar, bg='#2c3e50')
        frame_modos.pack(pady=10)

        btn_modo_estudo = tk.Button(
            frame_modos,
            text="📖 Modo Estudo",
            command=lambda: self.definir_modo('estudo'),
            bg='#3498db',
            fg='white',
            font=("Arial", 10, "bold"),
            padx=15,
            cursor='hand2'
        )
        btn_modo_estudo.pack(side='left', padx=5)

        btn_modo_prova = tk.Button(
            frame_modos,
            text="📝 Modo Prova",
            command=lambda: self.definir_modo('prova'),
            bg='#e74c3c',
            fg='white',
            font=("Arial", 10, "bold"),
            padx=15,
            cursor='hand2'
        )
        btn_modo_prova.pack(side='left', padx=5)

        # Botão teorias
        btn_teorias = tk.Button(
            frame_modos,
            text="📚 Diretório de Teorias",
            command=self.definir_diretorio_teorias,
            bg='#9b59b6',
            fg='white',
            font=("Arial", 10, "bold"),
            padx=15,
            cursor='hand2'
        )
        btn_teorias.pack(side='left', padx=5)

        # Botão imagens
        btn_imagens = tk.Button(
            frame_modos,
            text="🖼️ Diretório de Imagens",
            command=self.definir_diretorio_imagens,
            bg='#f39c12',
            fg='white',
            font=("Arial", 10, "bold"),
            padx=15,
            cursor='hand2'
        )
        btn_imagens.pack(side='left', padx=5)

        # Status
        self.lbl_status = tk.Label(
            self.frame_carregar,
            text="",
            font=("Arial", 10),
            bg='#2c3e50',
            fg='#f39c12'
        )
        self.lbl_status.pack(pady=10)

        # Botão sair
        btn_sair = tk.Button(
            self.frame_carregar,
            text="🚪 Sair",
            command=self.sair,
            bg='#e74c3c',
            fg='white',
            font=("Arial", 10),
            cursor='hand2'
        )
        btn_sair.pack(pady=10)

        # Modo atual
        self.modo_atual = 'estudo'

    def definir_diretorio_teorias(self):
        """Define o diretório onde estão os arquivos de teoria"""
        try:
            from tkinter import filedialog
            diretorio = filedialog.askdirectory(
                title="Selecione o diretório com os arquivos de teoria (.docx)"
            )
            if diretorio:
                self.diretorio_teorias = diretorio
                messagebox.showinfo(
                    "Diretório Definido",
                    f"Diretório de teorias definido:\n{diretorio}"
                )
                self.lbl_status.config(text=f"✅ Diretório de teorias: {os.path.basename(diretorio)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar diretório: {e}")

    def definir_diretorio_imagens(self):
        """Define o diretório onde estão as imagens das questões"""
        try:
            from tkinter import filedialog
            diretorio = filedialog.askdirectory(
                title="Selecione o diretório com as imagens das questões"
            )
            if diretorio:
                self.diretorio_imagens = diretorio
                messagebox.showinfo(
                    "Diretório Definido",
                    f"Diretório de imagens definido:\n{diretorio}"
                )
                self.lbl_status.config(text=f"✅ Diretório de imagens: {os.path.basename(diretorio)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar diretório: {e}")

    def sair(self):
        """Salva progresso antes de sair"""
        if self.questoes_agrupadas:
            resposta = messagebox.askyesno(
                "Salvar Progresso",
                "Deseja salvar seu progresso antes de sair?"
            )
            if resposta:
                self.salvar_progresso()

        self.root.quit()

    def definir_modo(self, modo):
        """Define o modo de estudo ou prova"""
        self.modo_atual = modo

        if modo == 'prova':
            # Configurar temporizador para prova
            self.timer_ativo = False
            self.tempo_restante = None

            # Perguntar tempo da prova
            tempo_str = simpledialog.askstring(
                "Tempo da Prova",
                "Digite o tempo da prova em minutos:",
                initialvalue="60"
            )

            if tempo_str:
                try:
                    minutos = int(tempo_str)
                    self.tempo_prova = minutos * 60
                    self.lbl_temporizador.config(text=f"⏱️ Tempo: {minutos:02d}:00")
                    messagebox.showinfo(
                        "Modo Prova",
                        f"Modo prova ativado!\n\n"
                        f"Você terá {minutos} minutos para responder.\n"
                        f"O temporizador começará na primeira questão."
                    )
                except:
                    messagebox.showwarning("Atenção", "Tempo inválido. Usando 60 minutos.")
                    self.tempo_prova = 60 * 60
            else:
                self.modo_atual = 'estudo'
                messagebox.showinfo("Informação", "Modo estudo mantido.")
        else:
            # Modo estudo
            self.timer_ativo = False
            self.tempo_restante = None
            self.lbl_temporizador.config(text="⏱️ Tempo: 00:00")
            messagebox.showinfo("Modo Estudo", "Modo estudo ativado!\n\nTempo livre para estudar.")

    def procurar_arquivo(self):
        """Abre diálogo para selecionar arquivo"""
        try:
            from tkinter import filedialog
            arquivo = filedialog.askopenfilename(
                title="Selecione o arquivo Excel",
                filetypes=[("Arquivos Excel", "*.xlsx *.xls *.xlsm"), ("Todos arquivos", "*.*")]
            )
            if arquivo:
                self.entry_arquivo.delete(0, tk.END)
                self.entry_arquivo.insert(0, arquivo)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar arquivo: {e}")

    def eh_algarismo_romano(self, texto):
        """Verifica se o texto é um algarismo romano (I, II, III, IV, V, etc.)"""
        texto = texto.strip().upper()
        # Padrão para algarismos romanos: I, II, III, IV, V, VI, VII, VIII, IX, X
        padrao_romano = r'^[IVXLCDM]+$'
        return bool(re.match(padrao_romano, texto)) and len(texto) <= 5

    def agrupar_questoes_corretamente(self, df):
        """Agrupa corretamente as linhas da mesma questão (cada alternativa em linha separada)"""
        questoes = []
        questao_atual = None

        for idx, row in df.iterrows():
            # Verificar se é uma nova questão (número de questão não é NaN)
            if pd.notna(row['Questão ']):
                # Salvar questão anterior se existir
                if questao_atual is not None:
                    questoes.append(questao_atual)

                # Criar nova questão
                questao_atual = {
                    'tomo': row['TOMO'] if pd.notna(row['TOMO']) else "N/A",
                    'numero': row['Questão '],
                    'enunciado': row['Enunciado'] if pd.notna(row['Enunciado']) else "",
                    'imagem': row['Imagem'] if pd.notna(row.get('Imagem', '')) else "",
                    # MUDADO DE "Introdução teórica" para "Imagem"
                    'alternativas': [],
                    'usa_algarismos_romanos': False  # Flag para detectar romanos
                }

            # Adicionar alternativa (mesmo que seja continuação da questão anterior)
            if pd.notna(row['Alternativas']):
                alternativa = {
                    'letra': str(row['Alternativas']).strip(),
                    'texto': str(row['Textos das Alternativas']).strip() if pd.notna(
                        row['Textos das Alternativas']) else "",
                    'analise': str(row['Análise das alternativas ']).strip() if pd.notna(
                        row['Análise das alternativas ']) else "",
                    'correta': str(row['Alternativa correta']).strip() == 'X' if pd.notna(
                        row['Alternativa correta']) else False
                }
                questao_atual['alternativas'].append(alternativa)

                # Verificar se é algarismo romano
                if self.eh_algarismo_romano(str(row['Alternativas'])):
                    questao_atual['usa_algarismos_romanos'] = True

        # Não esquecer da última questão
        if questao_atual is not None:
            questoes.append(questao_atual)

        return questoes

    def carregar_planilha(self):
        """Carrega a planilha Excel com tratamento para diferentes formatos"""
        arquivo = self.entry_arquivo.get().strip()

        if not arquivo:
            messagebox.showwarning("Atenção", "Por favor, informe o nome do arquivo!")
            return

        if not os.path.exists(arquivo):
            messagebox.showerror("Erro", f"Arquivo não encontrado: {arquivo}")
            return

        try:
            self.lbl_status.config(text="⏳ Carregando planilha...")
            self.root.update()

            # Tentar carregar com diferentes engines dependendo da extensão
            extensao = os.path.splitext(arquivo)[1].lower()

            if extensao == '.xlsx' or extensao == '.xlsm':
                # Usar openpyxl para arquivos .xlsx
                self.df = pd.read_excel(arquivo, engine='openpyxl')
            elif extensao == '.xls':
                # Usar xlrd para arquivos .xls (antigos)
                self.df = pd.read_excel(arquivo, engine='xlrd')
            else:
                # Deixar pandas escolher o engine
                self.df = pd.read_excel(arquivo)

            # Agrupar questões CORRETAMENTE
            self.questoes_agrupadas = self.agrupar_questoes_corretamente(self.df)

            if len(self.questoes_agrupadas) == 0:
                messagebox.showwarning("Atenção", "Nenhuma questão encontrada na planilha!")
                return

            # Resetar estatísticas para nova planilha
            self.total_acertos = 0
            self.total_erros = 0
            self.questoes_respondidas = set()
            self.atualizar_estatisticas()

            self.lbl_status.config(text=f"✅ {len(self.questoes_agrupadas)} questões carregadas!")

            # Mostrar primeira questão
            self.mostrar_questao()

        except ImportError as e:
            if 'xlrd' in str(e):
                messagebox.showerror(
                    "Erro - Biblioteca ausente",
                    "Biblioteca 'xlrd' não encontrada!\n\n"
                    "Para arquivos .xls (formato antigo), instale com:\n"
                    "pip install xlrd\n\n"
                    "Ou converta o arquivo para .xlsx"
                )
            elif 'openpyxl' in str(e):
                messagebox.showerror(
                    "Erro - Biblioteca ausente",
                    "Biblioteca 'openpyxl' não encontrada!\n\n"
                    "Para arquivos .xlsx, instale com:\n"
                    "pip install openpyxl"
                )
            else:
                messagebox.showerror("Erro de Importação", f"{e}")
            self.lbl_status.config(text="")

        except Exception as e:
            messagebox.showerror(
                "Erro ao carregar planilha",
                f"Erro: {e}\n\n"
                "Tente:\n"
                "1. Verificar se o arquivo não está aberto em outro programa\n"
                "2. Converter o arquivo para formato .xlsx\n"
                "3. Verificar se as colunas estão nomeadas corretamente"
            )
            self.lbl_status.config(text="")

    def mostrar_questao(self):
        """Mostra a questão atual com TODAS as alternativas"""

        # Limpar frame principal
        for widget in self.frame_principal.winfo_children():
            widget.destroy()

        questao = self.questoes_agrupadas[self.questao_atual]

        # Iniciar temporizador no modo prova (primeira questão)
        if self.modo_atual == 'prova' and not self.timer_ativo and self.tempo_restante is None:
            self.iniciar_temporizador_prova()

        # Frame container com scrollbar
        canvas = tk.Canvas(self.frame_principal, bg='#2c3e50')
        scrollbar = ttk.Scrollbar(self.frame_principal, orient="vertical", command=canvas.yview)
        frame_container = tk.Frame(canvas, bg='#2c3e50')

        frame_container.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=frame_container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Cabeçalho
        frame_cabecalho = tk.Frame(frame_container, bg='#34495e')
        frame_cabecalho.pack(fill='x', pady=(0, 20))

        lbl_info = tk.Label(
            frame_cabecalho,
            text=f"📖 TOMO {questao['tomo']} - Questão {int(questao['numero'])}",
            font=("Arial", 14, "bold"),
            bg='#34495e',
            fg='white'
        )
        lbl_info.pack(pady=10)

        # Indicador de algarismos romanos
        '''if questao['usa_algarismos_romanos']:
            lbl_romano = tk.Label(
                frame_cabecalho,
                text="ℹ️ Esta questão usa algarismos romanos (I, II, III...)",
                font=("Arial", 10, "italic"),
                bg='#34495e',
                fg='#f39c12'
            )
            lbl_romano.pack(pady=(0, 5))'''

        # Progresso
        frame_progresso = tk.Frame(frame_cabecalho, bg='#34495e')
        frame_progresso.pack(pady=5)

        progresso = f"Questão {self.questao_atual + 1} de {len(self.questoes_agrupadas)}"
        lbl_progresso = tk.Label(
            frame_progresso,
            text=progresso,
            font=("Arial", 10),
            bg='#34495e',
            fg='#f39c12'
        )
        lbl_progresso.pack()

        # Botão para abrir enunciado completo em janela separada
        btn_enunciado_completo = tk.Button(
            frame_cabecalho,
            text="📄 Ver Enunciado Completo",
            command=self.abrir_enunciado_completo,
            bg='#3498db',
            fg='white',
            font=("Arial", 10, "bold"),
            padx=15,
            cursor='hand2'
        )
        btn_enunciado_completo.pack(pady=(5, 0))

        # Enunciado
        frame_enunciado = tk.Frame(frame_container, bg='#2c3e50')
        frame_enunciado.pack(fill='x', pady=(0, 20))

        lbl_enun_titulo = tk.Label(
            frame_enunciado,
            text="📝 Enunciado:",
            font=("Arial", 12, "bold"),
            bg='#2c3e50',
            fg='white'
        )
        lbl_enun_titulo.pack(anchor='w', pady=(0, 5))

        # Texto do enunciado (versão resumida)
        enunciado_resumido = questao['enunciado']
        #if len(enunciado_resumido) > 800:
        #    enunciado_resumido = enunciado_resumido[:800] + "..."

        txt_enunciado = scrolledtext.ScrolledText(
            frame_enunciado,
            wrap=tk.WORD,
            font=("Arial", 11),
            bg='#ecf0f1',
            fg='#2c3e50',
            height=6,
            padx=15,
            pady=15
        )
        txt_enunciado.insert(tk.END, enunciado_resumido)
        txt_enunciado.config(state='disabled')
        txt_enunciado.pack(fill='x')

        # Imagem da questão (se existir) - APÓS O ENUNCIADO
        if questao['imagem']:
            self.mostrar_imagem_questao(questao['imagem'], frame_container)

        # Alternativas - AGORA MOSTRANDO TODAS!
        frame_alternativas = tk.Frame(frame_container, bg='#2c3e50')
        frame_alternativas.pack(fill='x', pady=(0, 20))

        lbl_alt = tk.Label(
            frame_alternativas,
            text="Escolha uma alternativa:",
            font=("Arial", 12, "bold"),
            bg='#2c3e50',
            fg='white'
        )
        lbl_alt.pack(anchor='w', pady=(10, 15))

        # Variável para armazenar seleção
        self.var_alternativa = tk.StringVar()
        self.var_alternativa.set("")

        # Armazenar referências aos botões de alternativa
        self.botoes_alternativas = {}

        # Criar botões para TODAS as alternativas
        for alt in questao['alternativas']:
            frame_alt = tk.Frame(frame_alternativas, bg='#2c3e50')
            frame_alt.pack(fill='x', pady=8)

            texto_completo = f"{alt['letra']}) {alt['texto']}"

            radio = tk.Radiobutton(
                frame_alt,
                text=texto_completo,
                variable=self.var_alternativa,
                value=alt['letra'],
                font=("Arial", 11),
                bg='#2c3e50',
                fg='white',
                selectcolor='#34495e',
                activebackground='#2c3e50',
                activeforeground='white',
                anchor='w',
                justify='left',
                wraplength=800
            )
            radio.pack(fill='x', padx=10)
            self.botoes_alternativas[alt['letra']] = radio

        # Área de feedback (inicialmente oculta)
        self.frame_feedback = tk.Frame(frame_container, bg='#2c3e50')
        self.frame_feedback.pack(fill='x', pady=(0, 20))

        # Botões de navegação
        frame_botoes = tk.Frame(frame_container, bg='#2c3e50')
        frame_botoes.pack(fill='x', pady=20)

        # Botão anterior
        if self.questao_atual > 0:
            btn_anterior = tk.Button(
                frame_botoes,
                text="⬅️ Anterior",
                command=self.questao_anterior,
                bg='#95a5a6',
                fg='white',
                font=("Arial", 11, "bold"),
                padx=20,
                cursor='hand2'
            )
            btn_anterior.pack(side='left', padx=5)

        # Botão verificar resposta
        btn_verificar = tk.Button(
            frame_botoes,
            text="✅ Verificar Resposta",
            command=self.verificar_resposta,
            bg='#27ae60',
            fg='white',
            font=("Arial", 11, "bold"),
            padx=20,
            cursor='hand2'
        )
        btn_verificar.pack(side='left', padx=5)

        # Botão mostrar resposta (desabilitado inicialmente)
        self.btn_mostrar_resposta = tk.Button(
            frame_botoes,
            text="🔍 Mostrar Resposta",
            command=self.mostrar_resposta_completa,
            bg='#f39c12',
            fg='white',
            font=("Arial", 11, "bold"),
            padx=20,
            state='disabled',
            cursor='hand2'
        )
        self.btn_mostrar_resposta.pack(side='left', padx=5)

        # Botão ver teoria
        btn_ver_teorias = tk.Button(
            frame_botoes,
            text="📖 Ver Teoria",
            command=self.abrir_teorias,
            bg='#9b59b6',
            fg='white',
            font=("Arial", 11, "bold"),
            padx=20,
            cursor='hand2'
        )
        btn_ver_teorias.pack(side='left', padx=5)

        # Botão próxima
        btn_proxima = tk.Button(
            frame_botoes,
            text="Próxima ➡️",
            command=self.proxima_questao,
            bg='#3498db',
            fg='white',
            font=("Arial", 11, "bold"),
            padx=20,
            cursor='hand2'
        )
        btn_proxima.pack(side='right', padx=5)

        # Botão menu principal
        btn_menu = tk.Button(
            frame_botoes,
            text="🏠 Menu Principal",
            command=self.voltar_menu,
            bg='#e74c3c',
            fg='white',
            font=("Arial", 11),
            padx=15,
            cursor='hand2'
        )
        btn_menu.pack(side='right', padx=5)

    def abrir_enunciado_completo(self):
        """Abre janela com o enunciado completo da questão atual"""
        try:
            from PIL import Image, ImageTk, ImageOps
            PIL_DISPONIVEL = True
        except ImportError:
            PIL_DISPONIVEL = False

        questao = self.questoes_agrupadas[self.questao_atual]

        # Criar janela para exibir enunciado completo
        janela_enunciado = tk.Toplevel(self.root)
        janela_enunciado.title(f"Enunciado Completo - TOMO {questao['tomo']} Questão {int(questao['numero'])}")
        janela_enunciado.geometry("900x700")
        janela_enunciado.configure(bg='#2c3e50')

        # Frame container com scrollbar
        canvas = tk.Canvas(janela_enunciado, bg='#2c3e50')
        scrollbar = ttk.Scrollbar(janela_enunciado, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas, bg='#2c3e50')

        frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Título
        lbl_titulo = tk.Label(
            frame,
            text=f"📖 Enunciado Completo - TOMO {questao['tomo']} Questão {int(questao['numero'])}",
            font=("Arial", 16, "bold"),
            bg='#2c3e50',
            fg='white',
            pady=20
        )
        lbl_titulo.pack()

        # Enunciado completo
        frame_enunciado = tk.Frame(frame, bg='#2c3e50')
        frame_enunciado.pack(fill='x', padx=20, pady=10)

        lbl_enun_titulo = tk.Label(
            frame_enunciado,
            text="📝 Enunciado:",
            font=("Arial", 14, "bold"),
            bg='#2c3e50',
            fg='white'
        )
        lbl_enun_titulo.pack(anchor='w', pady=(0, 10))

        txt_enunciado = scrolledtext.ScrolledText(
            frame_enunciado,
            wrap=tk.WORD,
            font=("Arial", 12),
            bg='#ecf0f1',
            fg='#2c3e50',
            padx=20,
            pady=20,
            height=26
        )
        txt_enunciado.insert(tk.END, questao['enunciado'])
        txt_enunciado.config(state='disabled')
        txt_enunciado.pack(fill='both', expand=True)

        # Imagem da questão (se existir)
        if questao['imagem'] and PIL_DISPONIVEL:
            self.mostrar_imagem_completa(questao['imagem'], frame)

        # Botão fechar
        btn_fechar = tk.Button(
            frame,
            text="Fechar",
            command=janela_enunciado.destroy,
            bg='#e74c3c',
            fg='white',
            font=("Arial", 11, "bold"),
            padx=30,
            pady=10,
            cursor='hand2'
        )
        btn_fechar.pack(pady=(20, 20))

    def mostrar_imagem_completa(self, nome_imagem, frame_pai):
        """Mostra a imagem da questão no enunciado completo"""
        try:
            from PIL import Image, ImageTk, ImageOps
            PIL_DISPONIVEL = True
        except ImportError:
            PIL_DISPONIVEL = False
            return

        if not PIL_DISPONIVEL:
            return

        # Verificar se o diretório existe
        if not os.path.exists(self.diretorio_imagens):
            return

        # Procurar imagem com diferentes extensões
        extensoes = ['.png', '.jpg', '.jpeg', '.gif', '.bmp']
        caminho_imagem = None

        for ext in extensoes:
            caminho_teste = os.path.join(self.diretorio_imagens, nome_imagem + ext)
            if os.path.exists(caminho_teste):
                caminho_imagem = caminho_teste
                break

        # Se não encontrou, tentar o nome exato
        if not caminho_imagem:
            caminho_teste = os.path.join(self.diretorio_imagens, nome_imagem)
            if os.path.exists(caminho_teste):
                caminho_imagem = caminho_teste

        # Se ainda não encontrou, retornar
        if not caminho_imagem:
            return

        try:
            # Abrir imagem com PIL
            img_pil = Image.open(caminho_imagem)

            # Redimensionar se necessário (máximo 800px de largura)
            largura_max = 800
            if img_pil.width > largura_max:
                proporcao = largura_max / img_pil.width
                nova_altura = int(img_pil.height * proporcao)
                img_pil = img_pil.resize((largura_max, nova_altura), Image.Resampling.LANCZOS)

            # Converter para PhotoImage
            img_tk = ImageTk.PhotoImage(img_pil)

            # Criar frame para imagem
            frame_imagem = tk.Frame(frame_pai, bg='#2c3e50')
            frame_imagem.pack(fill='x', padx=20, pady=20)

            # Criar label para imagem
            lbl_imagem = tk.Label(frame_imagem, image=img_tk, bg='#2c3e50')
            lbl_imagem.image = img_tk  # Manter referência
            lbl_imagem.pack(pady=10)

            # Adicionar legenda
            lbl_legenda = tk.Label(
                frame_imagem,
                text=f"Figura: {nome_imagem}",
                font=("Arial", 10, "italic"),
                bg='#2c3e50',
                fg='#95a5a6'
            )
            lbl_legenda.pack(pady=(10, 0))

        except Exception as e:
            print(f"Erro ao exibir imagem {nome_imagem}: {e}")

    def mostrar_imagem_questao(self, nome_imagem, frame_pai):
        """Mostra a imagem da questão no enunciado"""
        try:
            from PIL import Image, ImageTk, ImageOps
            PIL_DISPONIVEL = True
        except ImportError:
            PIL_DISPONIVEL = False
            print("Pillow não instalado. Imagens não serão exibidas.")
            return

        if not PIL_DISPONIVEL:
            return

        # Verificar se o diretório existe
        if not os.path.exists(self.diretorio_imagens):
            resposta = messagebox.askyesno(
                "Diretório não encontrado",
                f"Diretório '{self.diretorio_imagens}' não encontrado.\n\n"
                "Deseja selecionar o diretório agora?"
            )
            if resposta:
                self.definir_diretorio_imagens()
                # Tentar novamente mostrar imagem
                self.mostrar_imagem_questao(nome_imagem, frame_pai)
            return

        # Procurar imagem com diferentes extensões
        extensoes = ['.png', '.jpg', '.jpeg', '.gif', '.bmp']
        caminho_imagem = None

        for ext in extensoes:
            caminho_teste = os.path.join(self.diretorio_imagens, nome_imagem + ext)
            if os.path.exists(caminho_teste):
                caminho_imagem = caminho_teste
                break

        # Se não encontrou, tentar o nome exato
        if not caminho_imagem:
            caminho_teste = os.path.join(self.diretorio_imagens, nome_imagem)
            if os.path.exists(caminho_teste):
                caminho_imagem = caminho_teste

        # Se ainda não encontrou, mostrar mensagem
        if not caminho_imagem:
            print(f"Imagem não encontrada: {nome_imagem}")
            return

        try:
            # Abrir imagem com PIL
            img_pil = Image.open(caminho_imagem)

            # Redimensionar se necessário (máximo 800px de largura)
            largura_max = 800
            if img_pil.width > largura_max:
                proporcao = largura_max / img_pil.width
                nova_altura = int(img_pil.height * proporcao)
                img_pil = img_pil.resize((largura_max, nova_altura), Image.Resampling.LANCZOS)

            # Converter para PhotoImage
            img_tk = ImageTk.PhotoImage(img_pil)

            # Criar frame para imagem
            frame_imagem = tk.Frame(frame_pai, bg='#2c3e50')
            frame_imagem.pack(fill='x', pady=(0, 20))

            # Criar label para imagem
            lbl_imagem = tk.Label(frame_imagem, image=img_tk, bg='#2c3e50')
            lbl_imagem.image = img_tk  # Manter referência
            lbl_imagem.pack(padx=20, pady=10)

            # Adicionar legenda
            lbl_legenda = tk.Label(
                frame_imagem,
                text=f"Figura: {nome_imagem}",
                font=("Arial", 9, "italic"),
                bg='#2c3e50',
                fg='#95a5a6'
            )
            lbl_legenda.pack(pady=(0, 10))

        except Exception as e:
            print(f"Erro ao exibir imagem {nome_imagem}: {e}")

    def abrir_teorias(self):
        """Abre janela com a teoria da questão atual (incluindo imagens)"""
        if not DOCX_DISPONIVEL:
            messagebox.showerror(
                "Erro - Biblioteca ausente",
                "Biblioteca 'python-docx' não encontrada!\n\n"
                "Para ler arquivos .docx, instale com:\n"
                "pip install python-docx\n\n"
                "Para exibir imagens, também instale:\n"
                "pip install pillow"
            )
            return

        # Verificar se Pillow está instalado para imagens
        try:
            from PIL import Image, ImageTk, ImageOps
            PIL_DISPONIVEL = True
        except ImportError:
            PIL_DISPONIVEL = False
            messagebox.showwarning(
                "Aviso",
                "Biblioteca 'Pillow' não encontrada.\n"
                "As imagens não serão exibidas.\n\n"
                "Instale com: pip install pillow"
            )

        questao = self.questoes_agrupadas[self.questao_atual]

        # Converter TOMO para inteiro (remover .0)
        tomo = int(float(questao['tomo'])) if pd.notna(questao['tomo']) else 0

        # Converter número da questão para inteiro
        numero = int(float(questao['numero'])) if pd.notna(questao['numero']) else 0

        # Verificar se o diretório existe
        if not os.path.exists(self.diretorio_teorias):
            resposta = messagebox.askyesno(
                "Diretório não encontrado",
                f"Diretório '{self.diretorio_teorias}' não encontrado.\n\n"
                "Deseja selecionar o diretório agora?"
            )
            if resposta:
                self.definir_diretorio_teorias()
                self.abrir_teorias()  # Tentar novamente
            return

        # Gerar possíveis nomes de arquivo
        possiveis_nomes = [
            f"T{tomo}Q{numero}.docx",  # Formato: T1Q2.docx
            f"T{tomo}Q{numero:02d}.docx",  # Formato: T1Q02.docx (com zero à esquerda)
            f"Tomo{tomo}Questao{numero}.docx",  # Formato: Tomo1Questao2.docx
            f"Tomo {tomo} Questão {numero}.docx",  # Formato: Tomo 1 Questão 2.docx
            f"Tomo{tomo}Q{numero}.docx",  # Formato: Tomo1Q2.docx
        ]

        caminho_arquivo = None

        # Procurar por qualquer um dos formatos possíveis
        for nome in possiveis_nomes:
            caminho_teste = os.path.join(self.diretorio_teorias, nome)
            if os.path.exists(caminho_teste):
                caminho_arquivo = caminho_teste
                break

        # Se não encontrou, mostrar mensagem detalhada
        if not caminho_arquivo:
            # Listar arquivos disponíveis no diretório para ajudar
            try:
                arquivos = os.listdir(self.diretorio_teorias)
                arquivos_docx = [f for f in arquivos if f.endswith('.docx')]

                if arquivos_docx:
                    lista_arquivos = "\n".join(arquivos_docx[:10])  # Mostrar até 10 arquivos
                    if len(arquivos_docx) > 10:
                        lista_arquivos += f"\n... e mais {len(arquivos_docx) - 10} arquivos"

                    mensagem = f"Arquivo de teoria não encontrado para Tomo {tomo} Questão {numero}.\n\n"
                    mensagem += "Formatos procurados:\n"
                    for nome in possiveis_nomes:
                        mensagem += f"  - {nome}\n"
                    mensagem += f"\nArquivos .docx encontrados no diretório:\n{lista_arquivos}\n\n"
                    mensagem += "Verifique se o arquivo existe e está nomeado corretamente."

                    messagebox.showwarning("Teoria não encontrada", mensagem)
                else:
                    messagebox.showwarning(
                        "Teoria não encontrada",
                        f"Nenhum arquivo .docx encontrado no diretório:\n{self.diretorio_teorias}\n\n"
                        "Verifique se o diretório está correto."
                    )
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao listar arquivos: {e}")

            return

        try:
            # Criar janela para exibir teoria
            janela_teorias = tk.Toplevel(self.root)
            janela_teorias.title(f"Teoria - TOMO {tomo} Questão {numero}")
            janela_teorias.geometry("900x700")
            janela_teorias.configure(bg='#2c3e50')

            # Frame container com scrollbar
            canvas = tk.Canvas(janela_teorias, bg='#2c3e50')
            scrollbar = ttk.Scrollbar(janela_teorias, orient="vertical", command=canvas.yview)
            frame = tk.Frame(canvas, bg='#2c3e50')

            frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            # Título
            lbl_titulo = tk.Label(
                frame,
                text=f"📖 Teoria - TOMO {tomo} Questão {numero}",
                font=("Arial", 16, "bold"),
                bg='#2c3e50',
                fg='white',
                pady=20
            )
            lbl_titulo.pack()

            # Ler arquivo .docx
            doc = Document(caminho_arquivo)

            # Lista para armazenar referências de imagens (para evitar garbage collection)
            self.imagens_refs = []

            # Processar cada parágrafo do documento
            for paragrafo in doc.paragraphs:
                texto = paragrafo.text.strip()

                # Verificar se o parágrafo tem imagens
                tem_imagem = False

                # Verificar runs para encontrar imagens
                for run in paragrafo.runs:
                    # Verificar se o run contém imagens
                    r = run._element
                    blips = r.xpath('.//a:blip')

                    if blips:
                        tem_imagem = True
                        break

                # Se tiver imagem, processar
                if tem_imagem:
                    try:
                        # Extrair imagens do documento
                        import zipfile
                        from io import BytesIO

                        # Abrir o arquivo .docx como zip
                        with zipfile.ZipFile(caminho_arquivo) as docx_zip:
                            # Listar todas as imagens
                            imagens = [f for f in docx_zip.namelist() if f.startswith('word/media/')]

                            for img_path in imagens:
                                try:
                                    # Ler dados da imagem
                                    img_data = docx_zip.read(img_path)

                                    # Abrir imagem com PIL
                                    img_pil = Image.open(BytesIO(img_data))

                                    # Redimensionar se necessário (máximo 800px de largura)
                                    largura_max = 800
                                    if img_pil.width > largura_max:
                                        proporcao = largura_max / img_pil.width
                                        nova_altura = int(img_pil.height * proporcao)
                                        img_pil = img_pil.resize((largura_max, nova_altura), Image.Resampling.LANCZOS)

                                    # Converter para PhotoImage
                                    img_tk = ImageTk.PhotoImage(img_pil)

                                    # Criar label para imagem
                                    lbl_imagem = tk.Label(frame, image=img_tk, bg='#2c3e50')
                                    lbl_imagem.image = img_tk  # Manter referência
                                    lbl_imagem.pack(padx=20, pady=10)

                                    # Armazenar referência para evitar garbage collection
                                    self.imagens_refs.append(img_tk)

                                    # Adicionar legenda
                                    lbl_legenda = tk.Label(
                                        frame,
                                        text=f"Figura: {os.path.basename(img_path)}",
                                        font=("Arial", 9, "italic"),
                                        bg='#2c3e50',
                                        fg='#95a5a6'
                                    )
                                    lbl_legenda.pack(pady=(0, 15))

                                except Exception as e:
                                    print(f"Erro ao processar imagem {img_path}: {e}")
                                    continue

                    except Exception as e:
                        print(f"Erro ao extrair imagens: {e}")

                # Se tiver texto, exibir
                if texto:
                    # Verificar estilo do parágrafo para formatação
                    if paragrafo.style.name.startswith('Heading'):
                        # Títulos em negrito e maior
                        try:
                            nivel_heading = int(paragrafo.style.name.replace('Heading', ''))
                            fonte_tamanho = 16 - (nivel_heading - 1) * 2
                            if fonte_tamanho < 10:
                                fonte_tamanho = 10
                        except:
                            fonte_tamanho = 14

                        lbl_texto = tk.Label(
                            frame,
                            text=texto,
                            font=("Arial", fonte_tamanho, "bold"),
                            bg='#2c3e50',
                            fg='white',
                            wraplength=850,
                            justify='left'
                        )
                        lbl_texto.pack(anchor='w', padx=20, pady=(15, 5))
                    else:
                        # Parágrafos normais
                        txt_paragrafo = tk.Text(
                            frame,
                            wrap=tk.WORD,
                            font=("Arial", 11),
                            bg='#ecf0f1',
                            fg='#2c3e50',
                            padx=15,
                            pady=10,
                            height=3,
                            relief='flat'
                        )
                        txt_paragrafo.insert(tk.END, texto)
                        txt_paragrafo.config(state='disabled')
                        txt_paragrafo.pack(fill='x', padx=20, pady=5)

            # Botão fechar
            btn_fechar = tk.Button(
                frame,
                text="Fechar",
                command=janela_teorias.destroy,
                bg='#e74c3c',
                fg='white',
                font=("Arial", 11, "bold"),
                padx=30,
                pady=10,
                cursor='hand2'
            )
            btn_fechar.pack(pady=(20, 20))

        except Exception as e:
            messagebox.showerror(
                "Erro ao ler teoria",
                f"Erro ao ler o arquivo {os.path.basename(caminho_arquivo)}:\n\n{e}"
            )

    def iniciar_temporizador_prova(self):
        """Inicia o temporizador para o modo prova"""
        self.timer_ativo = True
        self.tempo_restante = self.tempo_prova
        self.atualizar_temporizador()

    def atualizar_temporizador(self):
        """Atualiza o display do temporizador"""
        if self.timer_ativo and self.tempo_restante is not None:
            minutos = self.tempo_restante // 60
            segundos = self.tempo_restante % 60

            self.lbl_temporizador.config(text=f"⏱️ Tempo: {minutos:02d}:{segundos:02d}")

            # Alerta quando faltar pouco tempo
            if self.tempo_restante <= 300:  # 5 minutos
                self.lbl_temporizador.config(fg='#e74c3c')

            if self.tempo_restante > 0:
                self.tempo_restante -= 1
                self.timer_id = self.root.after(1000, self.atualizar_temporizador)
            else:
                # Tempo esgotado
                self.timer_ativo = False
                messagebox.showwarning(
                    "⏰ Tempo Esgotado!",
                    "Seu tempo acabou!\n\nDeseja continuar respondendo sem limite de tempo?",
                    type='yesno'
                )

    def verificar_resposta(self):
        """Verifica se a resposta está correta e mostra feedback na tela"""
        if not self.var_alternativa.get():
            messagebox.showwarning("Atenção", "Por favor, selecione uma alternativa!")
            return

        self.alternativa_selecionada = self.var_alternativa.get()
        questao = self.questoes_agrupadas[self.questao_atual]

        # Encontrar alternativa correta
        alt_correta = None
        for alt in questao['alternativas']:
            if alt['correta']:
                alt_correta = alt['letra']
                break

        # Verificar se acertou
        acertou = (self.alternativa_selecionada == alt_correta)

        # Atualizar estatísticas
        if self.questao_atual not in self.questoes_respondidas:
            if acertou:
                self.total_acertos += 1
            else:
                self.total_erros += 1
            self.questoes_respondidas.add(self.questao_atual)
            self.atualizar_estatisticas()

        # Limpar frame de feedback anterior
        for widget in self.frame_feedback.winfo_children():
            widget.destroy()

        # Criar feedback visual
        if acertou:
            # Feedback positivo - verde
            frame_resultado = tk.Frame(self.frame_feedback, bg='#27ae60', padx=20, pady=20)
            frame_resultado.pack(fill='x')

            lbl_resultado = tk.Label(
                frame_resultado,
                text="🎉 PARABÉNS! Resposta CORRETA! 🎉",
                font=("Arial", 16, "bold"),
                bg='#27ae60',
                fg='white'
            )
            lbl_resultado.pack(pady=(0, 10))

            lbl_detalhe = tk.Label(
                frame_resultado,
                text=f"Você acertou! A alternativa {self.alternativa_selecionada} está correta.",
                font=("Arial", 12),
                bg='#27ae60',
                fg='white'
            )
            lbl_detalhe.pack()

            # Destacar alternativa selecionada em verde
            if self.alternativa_selecionada in self.botoes_alternativas:
                self.botoes_alternativas[self.alternativa_selecionada].config(
                    bg='#27ae60',
                    fg='white',
                    font=("Arial", 11, "bold")
                )

        else:
            # Feedback negativo - vermelho
            frame_resultado = tk.Frame(self.frame_feedback, bg='#e74c3c', padx=20, pady=20)
            frame_resultado.pack(fill='x')

            lbl_resultado = tk.Label(
                frame_resultado,
                text="❌ Resposta INCORRETA ❌",
                font=("Arial", 16, "bold"),
                bg='#e74c3c',
                fg='white'
            )
            lbl_resultado.pack(pady=(0, 10))

            lbl_detalhe = tk.Label(
                frame_resultado,
                text=f"Você escolheu: {self.alternativa_selecionada} | Resposta correta: {alt_correta}",
                font=("Arial", 12),
                bg='#e74c3c',
                fg='white'
            )
            lbl_detalhe.pack(pady=(5, 10))

            lbl_dica = tk.Label(
                frame_resultado,
                text="Clique em 'Mostrar Resposta' para ver as justificativas!",
                font=("Arial", 11),
                bg='#e74c3c',
                fg='white'
            )
            lbl_dica.pack()

            # Destacar alternativa selecionada em vermelho
            if self.alternativa_selecionada in self.botoes_alternativas:
                self.botoes_alternativas[self.alternativa_selecionada].config(
                    bg='#e74c3c',
                    fg='white',
                    font=("Arial", 11, "bold")
                )

            # Destacar alternativa correta em verde
            if alt_correta in self.botoes_alternativas:
                self.botoes_alternativas[alt_correta].config(
                    bg='#27ae60',
                    fg='white',
                    font=("Arial", 11, "bold")
                )

        # Desabilitar seleção
        for btn in self.botoes_alternativas.values():
            btn.config(state='disabled')

        # Habilitar botão mostrar resposta
        self.btn_mostrar_resposta.config(state='normal')

        self.resposta_verificada = True

        # Salvar progresso automaticamente
        self.salvar_progresso()

    def atualizar_estatisticas(self):
        """Atualiza o display das estatísticas"""
        total = self.total_acertos + self.total_erros
        self.lbl_estatisticas.config(
            text=f"📊 Acertos: {self.total_acertos} | Erros: {self.total_erros} | Total: {total}"
        )

    def mostrar_resposta_completa(self):
        """Mostra a resposta correta e justificativas com botões para cada alternativa"""

        questao = self.questoes_agrupadas[self.questao_atual]

        # Encontrar alternativa correta
        alt_correta = None
        for alt in questao['alternativas']:
            if alt['correta']:
                alt_correta = alt['letra']
                break

        # Criar nova janela
        janela_resposta = tk.Toplevel(self.root)
        janela_resposta.title("Resposta e Justificativas")
        janela_resposta.geometry("900x700")
        janela_resposta.configure(bg='#2c3e50')

        # Frame container com scrollbar
        canvas = tk.Canvas(janela_resposta, bg='#2c3e50')
        scrollbar = ttk.Scrollbar(janela_resposta, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas, bg='#2c3e50')

        frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Título
        lbl_titulo = tk.Label(
            frame,
            text="📝 Resposta Correta e Justificativas",
            font=("Arial", 16, "bold"),
            bg='#2c3e50',
            fg='white'
        )
        lbl_titulo.pack(pady=(20, 20))

        # Resposta correta
        frame_resposta = tk.Frame(frame, bg='#34495e', padx=20, pady=20)
        frame_resposta.pack(fill='x', pady=(0, 20), padx=20)

        if alt_correta:
            lbl_resposta = tk.Label(
                frame_resposta,
                text=f"✅ Alternativa Correta: {alt_correta}",
                font=("Arial", 14, "bold"),
                bg='#34495e',
                fg='#27ae60'
            )
            lbl_resposta.pack()

        # Sua resposta
        if self.alternativa_selecionada:
            cor = '#27ae60' if self.alternativa_selecionada == alt_correta else '#e74c3c'
            simbolo = "✅" if self.alternativa_selecionada == alt_correta else "❌"

            lbl_sua_resposta = tk.Label(
                frame_resposta,
                text=f"{simbolo} Sua resposta: {self.alternativa_selecionada}",
                font=("Arial", 12, "bold"),
                bg='#34495e',
                fg=cor
            )
            lbl_sua_resposta.pack(pady=10)

        # Verificar se usa algarismos romanos
        if questao['usa_algarismos_romanos']:
            # Mostrar TODAS as justificativas juntas
            lbl_just = tk.Label(
                frame,
                text="📚 Justificativas de todas as alternativas (algarismos romanos):",
                font=("Arial", 14, "bold"),
                bg='#2c3e50',
                fg='white'
            )
            lbl_just.pack(anchor='w', pady=(10, 15), padx=20)

            # Frame para todas as justificativas
            frame_todas_just = tk.Frame(frame, bg='#ecf0f1', padx=20, pady=20)
            frame_todas_just.pack(fill='both', expand=True, padx=20, pady=(0, 20))

            # Mostrar justificativa de cada alternativa
            for alt in questao['alternativas']:
                self.mostrar_justificativa_no_frame(alt, alt_correta, frame_todas_just)

        else:
            # Justificativas individuais com botões (comportamento normal)
            lbl_just = tk.Label(
                frame,
                text="📚 Clique em uma alternativa para ver sua justificativa:",
                font=("Arial", 14, "bold"),
                bg='#2c3e50',
                fg='white'
            )
            lbl_just.pack(anchor='w', pady=(10, 15), padx=20)

            # Frame para botões de alternativas
            frame_btns = tk.Frame(frame, bg='#2c3e50')
            frame_btns.pack(fill='x', pady=(0, 20), padx=20)

            # Criar botão para cada alternativa
            for alt in questao['alternativas']:
                cor_btn = '#27ae60' if alt['correta'] else '#3498db'

                btn_alt = tk.Button(
                    frame_btns,
                    text=f"Alternativa {alt['letra']}",
                    command=lambda a=alt, c=alt_correta: self.mostrar_justificativa_individual(a, c, frame),
                    bg=cor_btn,
                    fg='white',
                    font=("Arial", 10, "bold"),
                    padx=15,
                    pady=8,
                    cursor='hand2'
                )
                btn_alt.pack(side='left', padx=5)

            # Área de texto para justificativa atual
            self.frame_justificativa_atual = tk.Frame(frame, bg='#2c3e50')
            self.frame_justificativa_atual.pack(fill='both', expand=True, padx=20, pady=(0, 20))

            # Mostrar justificativa da alternativa correta por padrão
            if alt_correta:
                for alt in questao['alternativas']:
                    if alt['letra'] == alt_correta:
                        self.mostrar_justificativa_individual(alt, alt_correta, frame)
                        break
            elif questao['alternativas']:
                self.mostrar_justificativa_individual(questao['alternativas'][0], alt_correta, frame)

        # Botão fechar
        btn_fechar = tk.Button(
            frame,
            text="Fechar",
            command=janela_resposta.destroy,
            bg='#e74c3c',
            fg='white',
            font=("Arial", 11, "bold"),
            padx=30,
            pady=10,
            cursor='hand2'
        )
        btn_fechar.pack(pady=(0, 20))

    def mostrar_justificativa_no_frame(self, alternativa, alt_correta, frame_pai):
        """Mostra a justificativa de uma alternativa dentro de um frame existente"""

        # Frame para justificativa individual
        frame_just = tk.Frame(frame_pai, bg='#ffffff', padx=15, pady=15, bd=1, relief='solid')
        frame_just.pack(fill='x', pady=(0, 15))

        # Cabeçalho
        eh_correta = alternativa['correta']
        cor_titulo = '#27ae60' if eh_correta else '#e74c3c'
        simbolo = "✅" if eh_correta else "❌"

        lbl_titulo = tk.Label(
            frame_just,
            text=f"{simbolo} Alternativa {alternativa['letra']}",
            font=("Arial", 12, "bold"),
            bg='#ffffff',
            fg=cor_titulo
        )
        lbl_titulo.pack(anchor='w', pady=(0, 8))

        # Texto da alternativa
        if alternativa['texto']:
            lbl_texto = tk.Label(
                frame_just,
                text=f"Texto: {alternativa['texto']}",
                font=("Arial", 10, "italic"),
                bg='#ffffff',
                fg='#2c3e50',
                wraplength=800,
                justify='left'
            )
            lbl_texto.pack(anchor='w', pady=(0, 10))

        # Justificativa
        lbl_just_titulo = tk.Label(
            frame_just,
            text="Justificativa:",
            font=("Arial", 11, "bold"),
            bg='#ffffff',
            fg='#2c3e50'
        )
        lbl_just_titulo.pack(anchor='w', pady=(5, 3))

        # Texto da justificativa
        analise = alternativa['analise']
        if analise.startswith("Alternativa correta."):
            analise = analise.replace("Alternativa correta.", "", 1).strip()
        elif analise.startswith("Alternativa incorreta."):
            analise = analise.replace("Alternativa incorreta.", "", 1).strip()

        lbl_analise = tk.Label(
            frame_just,
            text=analise if analise else "Justificativa não disponível.",
            font=("Arial", 10),
            bg='#ffffff',
            fg='#2c3e50',
            wraplength=800,
            justify='left'
        )
        lbl_analise.pack(anchor='w', pady=(0, 5))

    def mostrar_justificativa_individual(self, alternativa, alt_correta, frame_pai):
        """Mostra a justificativa de uma alternativa específica"""

        # Limpar frame atual
        for widget in self.frame_justificativa_atual.winfo_children():
            widget.destroy()

        # Frame para justificativa
        frame_just = tk.Frame(self.frame_justificativa_atual, bg='#ecf0f1', padx=20, pady=20)
        frame_just.pack(fill='both', expand=True)

        # Cabeçalho
        eh_correta = alternativa['correta']
        cor_titulo = '#27ae60' if eh_correta else '#e74c3c'
        simbolo = "✅" if eh_correta else "❌"

        lbl_titulo = tk.Label(
            frame_just,
            text=f"{simbolo} Alternativa {alternativa['letra']}",
            font=("Arial", 13, "bold"),
            bg='#ecf0f1',
            fg=cor_titulo
        )
        lbl_titulo.pack(anchor='w', pady=(0, 10))

        # Texto da alternativa
        if alternativa['texto']:
            lbl_texto = tk.Label(
                frame_just,
                text=f"Texto: {alternativa['texto']}",
                font=("Arial", 11, "italic"),
                bg='#ecf0f1',
                fg='#2c3e50',
                wraplength=800,
                justify='left'
            )
            lbl_texto.pack(anchor='w', pady=(0, 15))

        # Justificativa
        lbl_just_titulo = tk.Label(
            frame_just,
            text="Justificativa:",
            font=("Arial", 12, "bold"),
            bg='#ecf0f1',
            fg='#2c3e50'
        )
        lbl_just_titulo.pack(anchor='w', pady=(10, 5))

        # Área de texto para justificativa
        txt_justificativa = scrolledtext.ScrolledText(
            frame_just,
            wrap=tk.WORD,
            font=("Arial", 11),
            bg='#ffffff',
            fg='#2c3e50',
            padx=15,
            pady=15,
            height=12
        )

        # Limpar "Alternativa correta." ou "Alternativa incorreta." do início
        analise = alternativa['analise']
        if analise.startswith("Alternativa correta."):
            analise = analise.replace("Alternativa correta.", "", 1).strip()
        elif analise.startswith("Alternativa incorreta."):
            analise = analise.replace("Alternativa incorreta.", "", 1).strip()

        txt_justificativa.insert(tk.END, analise if analise else "Justificativa não disponível.")
        txt_justificativa.config(state='disabled')
        txt_justificativa.pack(fill='both', expand=True)

    def questao_anterior(self):
        """Volta para questão anterior"""
        if self.questao_atual > 0:
            self.questao_atual -= 1
            self.alternativa_selecionada = None
            self.resposta_verificada = False
            self.mostrar_questao()

    def proxima_questao(self):
        """Vai para próxima questão"""
        if self.questao_atual < len(self.questoes_agrupadas) - 1:
            self.questao_atual += 1
            self.alternativa_selecionada = None
            self.resposta_verificada = False
            self.mostrar_questao()
        else:
            # Última questão - finalizar prova se estiver no modo prova
            if self.modo_atual == 'prova':
                self.finalizar_prova()
            else:
                resposta = messagebox.askyesno(
                    "Fim das questões",
                    "Você chegou ao fim das questões!\n\nDeseja recomeçar do início?"
                )
                if resposta:
                    self.questao_atual = 0
                    self.alternativa_selecionada = None
                    self.resposta_verificada = False
                    self.mostrar_questao()

    def finalizar_prova(self):
        """Finaliza o modo prova e mostra resumo"""
        # Parar temporizador
        self.timer_ativo = False
        if self.timer_id:
            self.root.after_cancel(self.timer_id)

        # Calcular estatísticas
        total_respondidas = len(self.questoes_respondidas)
        total_questoes = len(self.questoes_agrupadas)
        aproveitamento = (self.total_acertos / total_respondidas * 100) if total_respondidas > 0 else 0

        # Mostrar resumo em janela
        janela_resumo = tk.Toplevel(self.root)
        janela_resumo.title("Resumo da Prova")
        janela_resumo.geometry("600x500")
        janela_resumo.configure(bg='#2c3e50')

        frame_resumo = tk.Frame(janela_resumo, bg='#2c3e50', padx=30, pady=30)
        frame_resumo.pack(fill='both', expand=True)

        # Título
        lbl_titulo = tk.Label(
            frame_resumo,
            text="🎉 FIM DA PROVA! 🎉",
            font=("Arial", 20, "bold"),
            bg='#2c3e50',
            fg='#f39c12'
        )
        lbl_titulo.pack(pady=(0, 20))

        # Estatísticas
        frame_stats = tk.Frame(frame_resumo, bg='#34495e', padx=20, pady=20)
        frame_stats.pack(fill='x', pady=10)

        stats_texto = f"""
📊 RESUMO DA PROVA

Total de questões: {total_questoes}
Questões respondidas: {total_respondidas}

✅ Acertos: {self.total_acertos}
❌ Erros: {self.total_erros}

📈 Aproveitamento: {aproveitamento:.1f}%

⏱️ Tempo utilizado: {self.calcular_tempo_utilizado()}
        """

        lbl_stats = tk.Label(
            frame_stats,
            text=stats_texto,
            font=("Arial", 12, "bold"),
            bg='#34495e',
            fg='white',
            justify='left'
        )
        lbl_stats.pack()

        # Mensagem de incentivo
        if aproveitamento >= 80:
            mensagem = "🌟 Excelente! Você foi muito bem!"
        elif aproveitamento >= 60:
            mensagem = "👍 Bom desempenho! Continue estudando!"
        elif aproveitamento >= 40:
            mensagem = "📚 Precisa estudar mais. Não desista!"
        else:
            mensagem = "💪 Persistência é a chave! Continue tentando!"

        lbl_mensagem = tk.Label(
            frame_resumo,
            text=mensagem,
            font=("Arial", 14, "bold"),
            bg='#2c3e50',
            fg='#f39c12',
            pady=20
        )
        lbl_mensagem.pack()

        # Botões
        frame_botoes = tk.Frame(frame_resumo, bg='#2c3e50')
        frame_botoes.pack(pady=20)

        btn_reiniciar = tk.Button(
            frame_botoes,
            text="🔄 Refazer Prova",
            command=lambda: self.reiniciar_prova(janela_resumo),
            bg='#3498db',
            fg='white',
            font=("Arial", 11, "bold"),
            padx=20,
            pady=10,
            cursor='hand2'
        )
        btn_reiniciar.pack(side='left', padx=10)

        btn_estudo = tk.Button(
            frame_botoes,
            text="📖 Modo Estudo",
            command=lambda: self.mudar_para_modo_estudo(janela_resumo),
            bg='#27ae60',
            fg='white',
            font=("Arial", 11, "bold"),
            padx=20,
            pady=10,
            cursor='hand2'
        )
        btn_estudo.pack(side='left', padx=10)

        btn_menu = tk.Button(
            frame_botoes,
            text="🏠 Menu Principal",
            command=lambda: self.voltar_menu_com_confirmacao(janela_resumo),
            bg='#e74c3c',
            fg='white',
            font=("Arial", 11, "bold"),
            padx=20,
            pady=10,
            cursor='hand2'
        )
        btn_menu.pack(side='left', padx=10)

    def calcular_tempo_utilizado(self):
        """Calcula o tempo utilizado na prova"""
        tempo_usado = self.tempo_prova - self.tempo_restante if self.tempo_restante else self.tempo_prova
        minutos = tempo_usado // 60
        segundos = tempo_usado % 60
        return f"{minutos:02d}:{segundos:02d}"

    def reiniciar_prova(self, janela_resumo):
        """Reinicia a prova do zero"""
        janela_resumo.destroy()

        # Resetar estatísticas
        self.total_acertos = 0
        self.total_erros = 0
        self.questoes_respondidas = set()
        self.questao_atual = 0
        self.alternativa_selecionada = None
        self.resposta_verificada = False

        # Reiniciar temporizador
        self.timer_ativo = False
        self.tempo_restante = None

        # Atualizar estatísticas
        self.atualizar_estatisticas()

        # Mostrar primeira questão
        self.mostrar_questao()

    def mudar_para_modo_estudo(self, janela_resumo):
        """Muda para modo estudo após a prova"""
        janela_resumo.destroy()

        self.modo_atual = 'estudo'
        self.timer_ativo = False
        self.tempo_restante = None
        self.lbl_temporizador.config(text="⏱️ Tempo: 00:00", fg='#f39c12')

        messagebox.showinfo("Modo Estudo", "Agora você está no modo estudo!\n\nContinue revisando as questões.")

    def voltar_menu_com_confirmacao(self, janela_resumo):
        """Volta para o menu principal com confirmação"""
        janela_resumo.destroy()

        resposta = messagebox.askyesno(
            "Voltar ao Menu",
            "Deseja voltar ao menu principal?\n\nO progresso será perdido."
        )
        if resposta:
            # Resetar variáveis
            self.questao_atual = 0
            self.alternativa_selecionada = None
            self.resposta_verificada = False
            self.df = None
            self.questoes_agrupadas = []

            # Recriar TODOS os widgets do menu principal
            self.criar_widgets()

    def voltar_menu(self):
        """Volta para o menu principal LIMPANDO tudo corretamente"""
        resposta = messagebox.askyesno(
            "Voltar ao Menu",
            "Deseja voltar ao menu principal?\n\nO progresso será perdido."
        )
        if resposta:
            # Resetar variáveis
            self.questao_atual = 0
            self.alternativa_selecionada = None
            self.resposta_verificada = False
            self.df = None
            self.questoes_agrupadas = []

            # Recriar TODOS os widgets do menu principal
            self.criar_widgets()

    def salvar_progresso(self):
        """Salva o progresso atual em um arquivo JSON"""
        if not self.questoes_agrupadas:
            return

        progresso = {
            'data_salvamento': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'total_acertos': self.total_acertos,
            'total_erros': self.total_erros,
            'questoes_respondidas': list(self.questoes_respondidas),
            'questao_atual': self.questao_atual,
            'modo_atual': self.modo_atual,
            'tempo_restante': self.tempo_restante if self.tempo_restante else None,
            'diretorio_teorias': self.diretorio_teorias,
            'diretorio_imagens': self.diretorio_imagens
        }

        try:
            with open(self.arquivo_progresso, 'w', encoding='utf-8') as f:
                json.dump(progresso, f, indent=4, ensure_ascii=False)

            # Verificar se o widget ainda existe antes de atualizar
            if hasattr(self, 'lbl_status') and self.lbl_status.winfo_exists():
                self.lbl_status.config(text=f"💾 Progresso salvo em {datetime.now().strftime('%H:%M:%S')}")

        except Exception as e:
            # Verificar se o widget ainda existe antes de mostrar erro
            if hasattr(self, 'lbl_status') and self.lbl_status.winfo_exists():
                self.lbl_status.config(text=f"Erro ao salvar: {e}")
            else:
                print(f"Erro ao salvar progresso: {e}")

    def carregar_progresso(self):
        """Carrega o progresso salvo anteriormente"""
        if not os.path.exists(self.arquivo_progresso):
            return

        try:
            with open(self.arquivo_progresso, 'r', encoding='utf-8') as f:
                progresso = json.load(f)

            self.total_acertos = progresso.get('total_acertos', 0)
            self.total_erros = progresso.get('total_erros', 0)
            self.questoes_respondidas = set(progresso.get('questoes_respondidas', []))
            self.questao_atual = progresso.get('questao_atual', 0)
            self.modo_atual = progresso.get('modo_atual', 'estudo')
            self.tempo_restante = progresso.get('tempo_restante', None)
            self.diretorio_teorias = progresso.get('diretorio_teorias', "teorias")
            self.diretorio_imagens = progresso.get('diretorio_imagens', "imagens")

            self.atualizar_estatisticas()

            # Informar que progresso foi carregado
            data = progresso.get('data_salvamento', 'desconhecida')
            print(f"Progresso carregado de: {data}")

        except Exception as e:
            print(f"Erro ao carregar progresso: {e}")


def main():
    """Função principal"""
    root = tk.Tk()

    # Verificar dependências
    try:
        import pandas
    except ImportError:
        messagebox.showerror(
            "Erro",
            "Biblioteca pandas não encontrada!\n\n"
            "Instale com: pip install pandas"
        )
        return

    app = QuestoesApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
