import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageDraw, ImageFont 
from pdf2image import convert_from_path

# #############################################################################
# --- SEÇÃO 1: CONFIGURAÇÕES E UTILITÁRIOS ---
# #############################################################################

def resource_path(relative_path):
    """ Retorna o caminho absoluto para recursos, funcionando em dev e no PyInstaller """
    try:
        # Caminho temporário criado pelo PyInstaller
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- CONFIGURAÇÕES DE LAYOUT (Ajuste conforme sua V19) ---
DPI_CONVERSAO = 150 
POS_Y_TITULO = 260 
TAMANHO_FONTE = 44 
PASTA_FONTE = "Roboto"
ARQUIVO_FONTE_BOLD = "Roboto-Bold.ttf" 

# Caminhos dinâmicos usando resource_path
CAMINHO_FONTE_FINAL = resource_path(os.path.join(PASTA_FONTE, ARQUIVO_FONTE_BOLD))
# O Poppler deve estar na pasta do projeto para ser empacotado
POPPLER_BIN = resource_path("poppler/Library/bin") 

FOTOS_POR_PAGINA = 6 
GRID_COLUNAS = 2
GRID_LINHAS = 3 
EXTENSOES_PERMITIDAS = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff')

MARGEM_SUPERIOR = 320 
MARGEM_LATERAL = 120 
MARGEM_INFERIOR = 120 
PADDING = 30 
PAGINA_TEMPLATE_PARA_USAR = 1 

# #############################################################################
# --- SEÇÃO 2: O MOTOR DO PDF ---
# #############################################################################

def criar_pdf_com_fundo(pasta_fotos, arquivo_template, arquivo_saida, status_callback=None):
    def log(mensagem):
        if status_callback:
            status_callback(mensagem)

    # 1. Carregar o template
    log("Carregando modelo PDF...")
    template_pil = convert_from_path(
        arquivo_template, dpi=DPI_CONVERSAO, poppler_path=POPPLER_BIN,
        first_page=PAGINA_TEMPLATE_PARA_USAR + 1, last_page=PAGINA_TEMPLATE_PARA_USAR + 1
    )[0]
    
    LARGURA_A4, ALTURA_A4 = template_pil.size

    # 2. Calcular slots
    LARGURA_AREA_UTIL = LARGURA_A4 - (MARGEM_LATERAL * 2)
    ALTURA_AREA_UTIL = ALTURA_A4 - MARGEM_SUPERIOR - MARGEM_INFERIOR
    LARGURA_SLOT = int((LARGURA_AREA_UTIL - (PADDING * (GRID_COLUNAS - 1))) / GRID_COLUNAS)
    ALTURA_SLOT = int((ALTURA_AREA_UTIL - (PADDING * (GRID_LINHAS - 1))) / GRID_LINHAS)

    # 3. Listar imagens
    arquivos = [os.path.join(pasta_fotos, f) for f in sorted(os.listdir(pasta_fotos)) 
                if f.lower().endswith(EXTENSOES_PERMITIDAS)]
    
    if not arquivos:
        raise ValueError("Nenhuma imagem válida encontrada na pasta.")

    # 4. Carregar Fonte
    nome_pasta_titulo = os.path.basename(os.path.normpath(pasta_fotos)).upper()
    try:
        font = ImageFont.truetype(CAMINHO_FONTE_FINAL, size=TAMANHO_FONTE)
    except:
        font = ImageFont.load_default()

    lista_paginas_final = []
    
    # 5. Processamento das Páginas
    for i in range(0, len(arquivos), FOTOS_POR_PAGINA):
        log(f"Processando página {(i//FOTOS_POR_PAGINA) + 1}...")
        paths_da_pagina = arquivos[i : i + FOTOS_POR_PAGINA]
        pagina_atual = template_pil.copy()
        draw = ImageDraw.Draw(pagina_atual)

        # Desenhar Título
        bbox = draw.textbbox((0, 0), nome_pasta_titulo, font=font)
        text_width = bbox[2] - bbox[0]
        draw.text(((LARGURA_A4 - text_width) / 2, POS_Y_TITULO), nome_pasta_titulo, fill="black", font=font)

        for index, caminho_img in enumerate(paths_da_pagina):
            img_original = Image.open(caminho_img).convert('RGB')
            # Redimensionamento Proporcional (Crop Center)
            img_original.thumbnail((LARGURA_SLOT * 2, ALTURA_SLOT * 2), Image.Resampling.LANCZOS)
            
            # Cálculo de posição e colagem (simplificado para o exemplo)
            coluna, linha = index % GRID_COLUNAS, index // GRID_COLUNAS
            pos_x = MARGEM_LATERAL + (coluna * (LARGURA_SLOT + PADDING))
            pos_y = MARGEM_SUPERIOR + (linha * (ALTURA_SLOT + PADDING))
            
            # Redimensiona para o slot exato
            img_final = img_original.resize((LARGURA_SLOT, ALTURA_SLOT), Image.Resampling.LANCZOS)
            pagina_atual.paste(img_final, (pos_x, pos_y))
        
        lista_paginas_final.append(pagina_atual)

    # 6. Salvar
    lista_paginas_final[0].save(arquivo_saida, save_all=True, append_images=lista_paginas_final[1:], resolution=DPI_CONVERSAO)

# #############################################################################
# --- SEÇÃO 3: INTERFACE GRÁFICA ---
# #############################################################################

class AppGeradorPDF:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatizador de Relatórios")
        self.root.geometry("600x450")

        self.path_fotos = tk.StringVar(value="Selecione a pasta...")
        self.path_template = tk.StringVar(value="Selecione o modelo PDF...")

        self.setup_ui()

    def setup_ui(self):
        frame = ttk.Frame(self.root, padding="20")
        frame.pack(expand=True, fill=tk.BOTH)

        ttk.Label(frame, text="Relatório Fotográfico Automático", font=("Helvetica", 14, "bold")).pack(pady=(0, 20))

        # Pasta de Fotos
        ttk.Button(frame, text="1. Escolher Pasta de Fotos", command=self.selecionar_pasta).pack(fill=tk.X)
        ttk.Label(frame, textvariable=self.path_fotos, wraplength=500, foreground="gray").pack(pady=5)

        # Template
        ttk.Button(frame, text="2. Escolher Modelo PDF", command=self.selecionar_template).pack(fill=tk.X)
        ttk.Label(frame, textvariable=self.path_template, wraplength=500, foreground="gray").pack(pady=5)

        # Botão Ação
        self.btn_gerar = ttk.Button(frame, text="GERAR RELATÓRIO AGORA", command=self.start_thread)
        self.btn_gerar.pack(pady=20, ipady=10, fill=tk.X)

        # Barra de Status
        self.status_var = tk.StringVar(value="Pronto")
        self.lbl_status = ttk.Label(frame, textvariable=self.status_var, font=("Helvetica", 9, "italic"))
        self.lbl_status.pack()

    def selecionar_pasta(self):
        p = filedialog.askdirectory()
        if p: self.path_fotos.set(p)

    def selecionar_template(self):
        p = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if p: self.path_template.set(p)

    def update_status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def start_thread(self):
        # Validação básica
        if "Selecione" in self.path_fotos.get() or "Selecione" in self.path_template.get():
            messagebox.showwarning("Atenção", "Selecione os arquivos antes de continuar.")
            return

        arquivo_saida = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not arquivo_saida: return

        # Desativa o botão e inicia a thread
        self.btn_gerar.config(state=tk.DISABLED)
        threading.Thread(target=self.worker, args=(arquivo_saida,), daemon=True).start()

    def worker(self, arquivo_saida):
        try:
            criar_pdf_com_fundo(
                self.path_fotos.get(),
                self.path_template.get(),
                arquivo_saida,
                status_callback=self.update_status
            )
            messagebox.showinfo("Sucesso", "Relatório gerado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro Crítico", f"Erro durante o processamento:\n{str(e)}")
        finally:
            self.btn_gerar.config(state=tk.NORMAL)
            self.update_status("Pronto")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppGeradorPDF(root)
    root.mainloop()