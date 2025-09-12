import pandas as pd
import math
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QTextEdit, 
                             QFileDialog, QProgressBar, QMessageBox, QGroupBox,
                             QFormLayout, QLineEdit, QComboBox, QTableWidget, QTableWidgetItem)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import os


PAGE_WIDTH, PAGE_HEIGHT = A4
MARGEM_GERAL = 15 * mm

# Define áreas reservadas para cabeçalho e rodapé
HEADER_AREA_ALTURA = 25 * mm
FOOTER_AREA_ALTURA = 25 * mm

def desenhar_cabecalho(c, nome_arquivo):
    # Usa as constantes globais para se posicionar
    c.setFont("Helvetica-Bold", 14)
    y_pos_texto = PAGE_HEIGHT - HEADER_AREA_ALTURA + (10 * mm)
    c.drawCentredString(PAGE_WIDTH / 2, y_pos_texto, f"Desenho da Peça: {nome_arquivo}")
    
    y_pos_linha = PAGE_HEIGHT - HEADER_AREA_ALTURA
    c.line(MARGEM_GERAL, y_pos_linha, PAGE_WIDTH - MARGEM_GERAL, y_pos_linha)

def desenhar_rodape_aprimorado(c, row):
    """Desenha um rodapé que respeita a área definida pelas constantes globais."""
    largura_total = PAGE_WIDTH - 2 * MARGEM_GERAL
    
    # Posiciona o bloco do rodapé dentro da área reservada
    altura_bloco = 12 * mm
    y_rodape = FOOTER_AREA_ALTURA - altura_bloco - (5 * mm) # Centraliza na área com uma margem

    c.setStrokeColorRGB(0, 0, 0)
    c.rect(MARGEM_GERAL, y_rodape, largura_total, altura_bloco)
    
    coluna1_x = MARGEM_GERAL
    coluna2_x = MARGEM_GERAL + largura_total * 0.60
    coluna3_x = MARGEM_GERAL + largura_total * 0.80
    
    c.line(coluna2_x, y_rodape, coluna2_x, y_rodape + altura_bloco)
    c.line(coluna3_x, y_rodape, coluna3_x, y_rodape + altura_bloco)
    
    y_titulo, y_valor = y_rodape + altura_bloco - 4*mm, y_rodape + 4*mm

    c.setFont("Helvetica", 7)
    c.drawCentredString(coluna1_x + (coluna2_x - coluna1_x)/2, y_titulo, "NOME DA PEÇA / IDENTIFICADOR")
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(coluna1_x + (coluna2_x - coluna1_x)/2, y_valor, str(row.get('nome_arquivo', 'N/A')))
    
    c.setFont("Helvetica", 7)
    c.drawCentredString(coluna2_x + (coluna3_x - coluna2_x)/2, y_titulo, "ESPESSURA")
    c.setFont("Helvetica-Bold", 10)
    espessura_formatada = formatar_numero(row.get('espessura', 0))
    c.drawCentredString(coluna2_x + (coluna3_x - coluna2_x)/2, y_valor, f"{espessura_formatada} mm")
    
    c.setFont("Helvetica", 7)
    c.drawCentredString(coluna3_x + (PAGE_WIDTH - MARGEM_GERAL - coluna3_x)/2, y_titulo, "QUANTIDADE")
    c.setFont("Helvetica-Bold", 10)
    quantidade_formatada = formatar_numero(row.get('qtd', 0))
    c.drawCentredString(coluna3_x + (PAGE_WIDTH - MARGEM_GERAL - coluna3_x)/2, y_valor, quantidade_formatada)

def desenhar_erro_dados(c, forma):
    c.setFont("Helvetica-Bold", 14); c.drawCentredString(A4[0]/2, A4[1]/2, f"Dados inválidos ou ausentes para a forma: '{forma}'"); c.setFont("Helvetica", 10); c.drawCentredString(A4[0]/2, A4[1]/2 - 10*mm, "Verifique os valores na planilha Excel.")

def formatar_numero(valor):
    if valor == int(valor): return str(int(valor))
    return str(valor)

def desenhar_cota_horizontal(c, x1, x2, y, valor):
    """Desenha uma cota horizontal com a linha principal 'quebrada' para o texto."""
    FONT_NAME, FONT_SIZE = "Helvetica", 12
    c.setFont(FONT_NAME, FONT_SIZE)
    
    texto = formatar_numero(valor)
    largura_texto = c.stringWidth(texto, FONT_NAME, FONT_SIZE)
    
    centro_x = (x1 + x2) / 2
    y_texto = y + 1.5*mm  
    
    padding = 2 * mm
    gap_inicio = centro_x - (largura_texto / 2) - padding
    gap_fim = centro_x + (largura_texto / 2) + padding
    
  
    c.line(x1, y, gap_inicio, y)
    c.line(gap_fim, y, x2, y)
    
    
    c.drawCentredString(centro_x, y_texto, texto)


def desenhar_cota_horizontal(c, x1, x2, y, texto): 
    """Desenha uma cota horizontal com 'ticks' de limite."""
    FONT_NAME, FONT_SIZE = "Helvetica", 10 # Tamanho de fonte reduzido para clareza
    c.setFont(FONT_NAME, FONT_SIZE)
    
    # Linha principal da cota
    c.line(x1, y, x2, y)

    # Ticks de limite (traços a 45 graus)
    tick_len = 2 * mm
    c.line(x1, y - tick_len/2, x1 + tick_len/2, y + tick_len/2)
    c.line(x2, y - tick_len/2, x2 - tick_len/2, y + tick_len/2)
    
    # Posiciona o texto acima da linha, quebrado no meio
    largura_texto = c.stringWidth(texto, FONT_NAME, FONT_SIZE)
    centro_x = (x1 + x2) / 2
    y_texto = y + 1*mm
    
    padding = 1.5 * mm
    gap_inicio = centro_x - (largura_texto / 2) - padding
    gap_fim = centro_x + (largura_texto / 2) + padding
    
    # Redesenha a linha com o espaço para o texto
    c.line(x1, y, gap_inicio, y)
    c.line(gap_fim, y, x2, y)
    c.drawCentredString(centro_x, y_texto, texto)

def desenhar_cota_vertical(c, y1, y2, x, texto):
    """Desenha uma cota vertical com 'ticks' de limite."""
    FONT_NAME, FONT_SIZE = "Helvetica", 10
    c.setFont(FONT_NAME, FONT_SIZE)

    # Linha principal da cota
    c.line(x, y1, x, y2)

    # Ticks de limite
    tick_len = 2 * mm
    c.line(x - tick_len/2, y1, x + tick_len/2, y1 + tick_len/2)
    c.line(x - tick_len/2, y2, x + tick_len/2, y2 - tick_len/2)

    # Posiciona o texto rotacionado e quebra a linha
    largura_texto = c.stringWidth(texto, FONT_NAME, FONT_SIZE)
    altura_texto = FONT_SIZE
    centro_y = (y1 + y2) / 2
    
    c.saveState()
    c.translate(x, centro_y)
    c.rotate(90)
    # Fundo branco para quebrar a linha de cota
    c.setFillColorRGB(1, 1, 1) 
    c.rect(-largura_texto/2 - 1*mm, -altura_texto/2, largura_texto + 2*mm, altura_texto, stroke=0, fill=1)
    # Texto
    c.setFillColorRGB(0, 0, 0)
    c.drawCentredString(0, -1.5*mm, texto)
    c.restoreState()



def desenhar_retangulo(c, row):
    largura, altura = row.get('largura', 0), row.get('altura', 0)
    if largura <= 0 or altura <= 0: desenhar_erro_dados(c, "Retângulo"); return

    max_w = PAGE_WIDTH - 2 * MARGEM_GERAL
    max_h = PAGE_HEIGHT - HEADER_AREA_ALTURA - FOOTER_AREA_ALTURA
    
    dist_cota_furo = 8 * mm
    dist_cota_total = 16 * mm
    overshoot = 2 * mm # O quanto a linha de extensão passa da linha de cota

    # --- LÓGICA DE CENTRALIZAÇÃO FINAL ---
    # O espaço total ocupado pelas cotas é a posição da cota mais externa
    espaco_cota_x = dist_cota_total 
    espaco_cota_y = dist_cota_total

    # A escala deve garantir que a PEÇA + ESPAÇO DAS COTAS caibam na área útil
    escala = min(
        (max_w - espaco_cota_x) / largura, 
        (max_h - espaco_cota_y) / altura
    ) * 0.95 # Fator de respiro
    
    dw, dh = largura * escala, altura * escala

    # O bloco visual é a peça desenhada (dw, dh) mais o espaço reservado para as cotas
    bloco_visual_width = dw + espaco_cota_x
    bloco_visual_height = dh + espaco_cota_y
    
    base_desenho_y = FOOTER_AREA_ALTURA
    
    # Centraliza o BLOCO VISUAL na área de desenho
    inicio_bloco_x = MARGEM_GERAL + (max_w - bloco_visual_width) / 2
    inicio_bloco_y = base_desenho_y + (max_h - bloco_visual_height) / 2

    # A origem da PEÇA (x0, y0) é o canto do bloco deslocado pelo espaço das cotas
    x0 = inicio_bloco_x + espaco_cota_x
    y0 = inicio_bloco_y + espaco_cota_y
    # --- FIM DA LÓGICA DE CENTRALIZAÇÃO ---
    
    c.rect(x0, y0, dw, dh)
    
    furos = row.get('furos')
    if isinstance(furos, list) and furos:
        for furo in furos:
            c.circle(x0 + (furo['x'] * escala), y0 + (furo['y'] * escala), (furo['diam'] / 2) * escala, stroke=1, fill=0)

        # Cotas em cadeia
        y_pos_cota = y0 - dist_cota_furo
        unique_x_coords = sorted(list(set(f['x'] for f in furos)))
        dim_points_x = [0] + unique_x_coords + [largura]
        for x_real in dim_points_x:
            c.line(x0 + x_real * escala, y0, x0 + x_real * escala, y_pos_cota - overshoot)
        for i in range(1, len(dim_points_x)):
            start_x, end_x = dim_points_x[i-1], dim_points_x[i]
            if (end_x - start_x) > 0.01:
                desenhar_cota_horizontal(c, x0 + start_x * escala, x0 + end_x * escala, y_pos_cota, formatar_numero(end_x - start_x))

        x_pos_cota = x0 - dist_cota_furo
        unique_y_coords = sorted(list(set(f['y'] for f in furos)))
        dim_points_y = [0] + unique_y_coords + [altura]
        for y_real in dim_points_y:
            c.line(x0, y0 + y_real * escala, x_pos_cota - overshoot, y0 + y_real * escala)
        for i in range(1, len(dim_points_y)):
            start_y, end_y = dim_points_y[i-1], dim_points_y[i]
            if (end_y - start_y) > 0.01:
                desenhar_cota_vertical(c, y0 + start_y * escala, y0 + end_y * escala, x_pos_cota, formatar_numero(end_y - start_y))

        primeiro_furo = furos[0]
        desenhar_cota_diametro_furo(c, x0 + primeiro_furo['x'] * escala, y0 + primeiro_furo['y'] * escala, (primeiro_furo['diam'] / 2) * escala, primeiro_furo['diam'])

    # Cotas Totais
    y_cota_total = y0 - dist_cota_total
    c.line(x0, y0, x0, y_cota_total - overshoot); c.line(x0 + dw, y0, x0 + dw, y_cota_total - overshoot)
    desenhar_cota_horizontal(c, x0, x0 + dw, y_cota_total, formatar_numero(largura))

    x_cota_total = x0 - dist_cota_total
    c.line(x0, y0, x_cota_total - overshoot, y0); c.line(x0, y0 + dh, x_cota_total - overshoot, y0 + dh)
    desenhar_cota_vertical(c, y0, y0 + dh, x_cota_total, formatar_numero(altura))

def desenhar_circulo(c, row):
    diametro = row.get('diametro', 0)
    if diametro <= 0: desenhar_erro_dados(c, "Círculo"); return

    max_w = PAGE_WIDTH - 2 * MARGEM_GERAL
    max_h = PAGE_HEIGHT - HEADER_AREA_ALTURA - FOOTER_AREA_ALTURA
    
    dist_cota = 8*mm
    overshoot = 2*mm

    escala = min(
        (max_w - dist_cota - overshoot) / diametro, 
        (max_h - dist_cota - overshoot) / diametro
    ) * 0.95
    
    raio_desenhado = (diametro * escala) / 2
    dw_dh_desenhado = raio_desenhado * 2
    
    bloco_visual_width = dw_dh_desenhado + dist_cota + overshoot
    bloco_visual_height = dw_dh_desenhado + dist_cota + overshoot

    base_desenho_y = FOOTER_AREA_ALTURA
    
    inicio_bloco_x = MARGEM_GERAL + (max_w - bloco_visual_width) / 2
    inicio_bloco_y = base_desenho_y + (max_h - bloco_visual_height) / 2

    cx = inicio_bloco_x + dist_cota + overshoot + raio_desenhado
    cy = inicio_bloco_y + dist_cota + overshoot + raio_desenhado

    c.circle(cx, cy, raio_desenhado, stroke=1, fill=0)

    furos = row.get('furos')
    if isinstance(furos, list):
        x0_peca = cx - raio_desenhado
        y0_peca = cy - raio_desenhado
        for furo in furos:
            c.circle(x0_peca + (furo['x'] * escala), y0_peca + (furo['y'] * escala), (furo['diam'] / 2) * escala, stroke=1, fill=0)

    y_cota_h = cy - raio_desenhado - dist_cota
    c.line(cx - raio_desenhado, cy, cx - raio_desenhado, y_cota_h - overshoot)
    c.line(cx + raio_desenhado, cy, cx + raio_desenhado, y_cota_h - overshoot)
    desenhar_cota_horizontal(c, cx - raio_desenhado, cx + raio_desenhado, y_cota_h, f"Ø {formatar_numero(diametro)}")

def desenhar_triangulo_retangulo(c, row):
    base, altura = row.get('rt_base', 0), row.get('rt_height', 0)
    if base <= 0 or altura <= 0: desenhar_erro_dados(c, "Triângulo Retângulo"); return

    max_w = PAGE_WIDTH - 2 * MARGEM_GERAL
    max_h = PAGE_HEIGHT - HEADER_AREA_ALTURA - FOOTER_AREA_ALTURA
    
    dist_cota = 8 * mm
    overshoot = 2 * mm

    escala = min(
        (max_w - dist_cota - overshoot) / base, 
        (max_h - dist_cota - overshoot) / altura
    ) * 0.95
    
    db, dh = base * escala, altura * escala

    bloco_visual_width = db + dist_cota + overshoot
    bloco_visual_height = dh + dist_cota + overshoot

    base_desenho_y = FOOTER_AREA_ALTURA
    
    inicio_bloco_x = MARGEM_GERAL + (max_w - bloco_visual_width) / 2
    inicio_bloco_y = base_desenho_y + (max_h - bloco_visual_height) / 2

    x0 = inicio_bloco_x + dist_cota + overshoot
    y0 = inicio_bloco_y + dist_cota + overshoot
    
    path = c.beginPath(); path.moveTo(x0, y0); path.lineTo(x0 + db, y0); path.lineTo(x0, y0 + dh); path.close(); c.drawPath(path)

    furos = row.get('furos')
    if isinstance(furos, list) and furos:
        for furo in furos:
            c.circle(x0 + (furo['x'] * escala), y0 + (furo['y'] * escala), (furo['diam'] / 2) * escala, stroke=1, fill=0)
    
    y_cota_h = y0 - dist_cota
    c.line(x0, y0, x0, y_cota_h - overshoot); c.line(x0 + db, y0, x0 + db, y_cota_h - overshoot)
    desenhar_cota_horizontal(c, x0, x0 + db, y_cota_h, formatar_numero(base))

    x_cota_v = x0 - dist_cota
    c.line(x0, y0, x_cota_v - overshoot, y0); c.line(x0, y0 + dh, x_cota_v - overshoot, y0 + dh)
    desenhar_cota_vertical(c, y0, y0 + dh, x_cota_v, formatar_numero(altura))



def desenhar_forma(c, row):
    desenhar_cabecalho(c, row.get('nome_arquivo', 'SEM NOME'))
    desenhar_rodape_aprimorado(c, row)
    forma = str(row.get('forma', '')).strip().lower()
    if forma == 'rectangle':
        desenhar_retangulo(c, row)   # <<<<----- Chamada de Função das formas "Implementar Novas Formas" 
    elif forma == 'circle':
        desenhar_circulo(c, row)
    elif forma == 'right_triangle':
        desenhar_triangulo_retangulo(c, row)
    else:
        c.setFont("Helvetica", 12); c.drawCentredString(A4[0]/2, A4[1]/2, f"Forma '{forma}' desconhecida ou não implementada.")

def desenhar_cota_diametro_furo(c, x, y, raio, diametro_real):
    """
    Desenha uma linha de chamada aprimorada para cotar o diâmetro de um furo.
    O texto agora fica posicionado de forma limpa acima da linha.
    """
    c.saveState()
    c.setFont("Helvetica", 10)
    texto = f"Ø {formatar_numero(diametro_real)}"
    largura_texto = c.stringWidth(texto, "Helvetica", 10)

    # Ponto na borda do círculo (a 45 graus, no quadrante superior direito)
    ponto_borda_x = x + raio * 0.7071 # cos(45)
    ponto_borda_y = y + raio * 0.7071 # sin(45)

    # Ponto intermediário da linha de chamada
    ponto_meio_x = ponto_borda_x + 4 * mm
    ponto_meio_y = ponto_borda_y + 4 * mm

    # Ponto final da "prateleira" horizontal
    ponto_final_x = ponto_meio_x + largura_texto + 2*mm

    # Desenha a linha de chamada
    path = c.beginPath()
    path.moveTo(ponto_borda_x, ponto_borda_y)
    path.lineTo(ponto_meio_x, ponto_meio_y)
    path.lineTo(ponto_final_x, ponto_meio_y)
    c.drawPath(path)

    # Desenha o texto acima da linha/prateleira
    c.drawString(ponto_meio_x + 1*mm, ponto_meio_y + 1*mm, texto)
    
    c.restoreState()

def gerar_relatorio_pdf(df, nome_arquivo):
    doc = SimpleDocTemplate(nome_arquivo, pagesize=A4); story = []
    styles = getSampleStyleSheet(); story.append(Paragraph("Relatório Geral de Peças a Produzir", styles['h1']))
    colunas_relatorio = ['nome_arquivo', 'forma', 'espessura', 'qtd', 'largura', 'altura', 'diametro', 'rt_base', 'rt_height']
    df_relatorio = df[[col for col in colunas_relatorio if col in df.columns]].copy()
    dados_tabela = [df_relatorio.columns.tolist()] + df_relatorio.fillna('-').values.tolist()
    tabela = Table(dados_tabela)
    tabela.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.darkslategray),('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),('ALIGN', (0, 0), (-1, -1), 'CENTER'),('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),('GRID', (0, 0), (-1, -1), 1, colors.black)]))
    story.append(tabela); doc.build(story)
    print(f"\nArquivo de relatório '{nome_arquivo}' gerado com sucesso!")



class ProcessThread(QThread):
    update_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, dataframe_to_process):
        super().__init__()
        self.df = dataframe_to_process
        
    def run(self):
        try:
            self.update_signal.emit("Dados recebidos para processamento.")
            
            # A leitura do Excel foi removida daqui, pois o DataFrame já é passado pronto.
            df = self.df
            
            self.update_signal.emit("Processando dados...")
            
            # Normalização e validação dos dados
            df.columns = df.columns.str.strip().str.lower()
            df = df.loc[:, ~df.columns.duplicated()]
            
            todas_dimensoes_possiveis = [
                'largura', 'altura', 'diametro', 'rt_base', 'rt_height', 
                'trapezoid_large_base', 'trapezoid_height'
            ]
            colunas_para_converter = [col for col in df.columns if col in todas_dimensoes_possiveis]
            
            for col in colunas_para_converter:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            self.update_signal.emit("Gerando relatório geral...")
            gerar_relatorio_pdf(df, "Relatorio_Geral_de_Pecas.pdf")
            
            if 'espessura' not in df.columns:
                raise KeyError("'espessura' -> Esta coluna é obrigatória para agrupar os desenhos.")
            
            df['espessura'] = df['espessura'].fillna('Sem_Espessura')
            grouped = df.groupby('espessura')
            total_groups = len(grouped)
            
            if total_groups == 0:
                self.update_signal.emit("Nenhuma peça para gerar PDFs.")
                self.finished_signal.emit(True, "Nenhuma peça encontrada para processar.")
                return

            self.update_signal.emit(f"Gerando {total_groups} arquivos PDF agrupados por espessura...")
            
            for i, (espessura, group) in enumerate(grouped):
                nome_espessura = str(espessura).replace('.', '_')
                nome_arquivo_pdf = f"Desenhos_Espessura_{nome_espessura}mm.pdf"
                c = canvas.Canvas(nome_arquivo_pdf, pagesize=A4)
                
                for index, row in group.iterrows():
                    desenhar_forma(c, row)
                    c.showPage()
                
                c.save()
                progress = int((i + 1) / total_groups * 100)
                self.progress_signal.emit(progress)
                self.update_signal.emit(f"Arquivo '{nome_arquivo_pdf}' gerado com {len(group)} peças.")
            
            self.finished_signal.emit(True, "Processamento concluído com sucesso!")
            
        except Exception as e:
            self.finished_signal.emit(False, f"Erro durante o processamento: {str(e)}")



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerador de Desenhos Técnicos")
        self.setGeometry(100, 100, 1100, 850)
        
        self.colunas_df = ['nome_arquivo', 'forma', 'espessura', 'qtd', 'largura', 'altura', 'diametro', 'rt_base', 'rt_height', 'furos']
        self.manual_df = pd.DataFrame(columns=self.colunas_df)
        self.excel_df = pd.DataFrame(columns=self.colunas_df)
        self.furos_atuais = []

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # ... (O layout da interface permanece o mesmo da versão anterior) ...
        top_h_layout = QHBoxLayout()
        left_v_layout = QVBoxLayout()
        
        file_group = QGroupBox("1. Carregar Planilha Excel (Opcional)")
        file_layout = QVBoxLayout()
        self.file_label = QLabel("Nenhuma planilha selecionada")
        file_button_layout = QHBoxLayout()
        self.select_file_btn = QPushButton("Selecionar Planilha")
        self.select_file_btn.clicked.connect(self.select_file)
        self.clear_excel_btn = QPushButton("Limpar Planilha")
        self.clear_excel_btn.clicked.connect(self.clear_excel_data)
        self.clear_excel_btn.setEnabled(False)
        file_button_layout.addWidget(self.select_file_btn)
        file_button_layout.addWidget(self.clear_excel_btn)
        file_layout.addWidget(self.file_label)
        file_layout.addLayout(file_button_layout)
        file_group.setLayout(file_layout)
        left_v_layout.addWidget(file_group)

        manual_group = QGroupBox("2. Informações da Peça")
        manual_layout = QFormLayout()
        self.nome_input = QLineEdit()
        self.espessura_input = QLineEdit()
        self.qtd_input = QLineEdit()
        self.forma_combo = QComboBox()
        self.forma_combo.addItems(['rectangle', 'circle', 'right_triangle'])
        self.largura_input = QLineEdit()
        self.altura_input = QLineEdit()
        self.diametro_input = QLineEdit()
        self.rt_base_input = QLineEdit()
        self.rt_height_input = QLineEdit()
        manual_layout.addRow("Nome/ID da Peça:", self.nome_input)
        manual_layout.addRow("Forma:", self.forma_combo)
        manual_layout.addRow("Espessura (mm):", self.espessura_input)
        manual_layout.addRow("Quantidade:", self.qtd_input)
        self.largura_row = [QLabel("Largura:"), self.largura_input]; manual_layout.addRow(self.largura_row[0], self.largura_row[1])
        self.altura_row = [QLabel("Altura:"), self.altura_input]; manual_layout.addRow(self.altura_row[0], self.altura_row[1])
        self.diametro_row = [QLabel("Diâmetro:"), self.diametro_input]; manual_layout.addRow(self.diametro_row[0], self.diametro_row[1])
        self.rt_base_row = [QLabel("Base Triângulo:"), self.rt_base_input]; manual_layout.addRow(self.rt_base_row[0], self.rt_base_row[1])
        self.rt_height_row = [QLabel("Altura Triângulo:"), self.rt_height_input]; manual_layout.addRow(self.rt_height_row[0], self.rt_height_row[1])
        manual_group.setLayout(manual_layout)
        left_v_layout.addWidget(manual_group)
        left_v_layout.addStretch()
        top_h_layout.addLayout(left_v_layout)

        furos_main_group = QGroupBox("3. Adicionar Furos à Peça")
        furos_main_layout = QVBoxLayout()
        self.rep_group = QGroupBox("Furação Rápida (Replicar Furos nos Cantos)")
        rep_layout = QFormLayout()
        self.rep_diam_input = QLineEdit()
        self.rep_offset_input = QLineEdit()
        rep_layout.addRow("Diâmetro dos Furos:", self.rep_diam_input)
        rep_layout.addRow("Distância da Borda (Offset):", self.rep_offset_input)
        self.replicate_btn = QPushButton("Replicar Furos")
        self.replicate_btn.clicked.connect(self.replicate_holes)
        rep_layout.addRow(self.replicate_btn)
        self.rep_group.setLayout(rep_layout)
        furos_main_layout.addWidget(self.rep_group)
        man_group = QGroupBox("Furos Manuais")
        man_layout = QVBoxLayout()
        man_form_layout = QFormLayout()
        self.diametro_furo_input = QLineEdit()
        self.pos_x_input = QLineEdit()
        self.pos_y_input = QLineEdit()
        man_form_layout.addRow("Diâmetro do Furo:", self.diametro_furo_input)
        man_form_layout.addRow("Posição X (da origem):", self.pos_x_input)
        man_form_layout.addRow("Posição Y (da origem):", self.pos_y_input)
        self.add_furo_btn = QPushButton("Adicionar Furo Manual")
        self.add_furo_btn.clicked.connect(self.add_furo_temp)
        man_layout.addLayout(man_form_layout)
        man_layout.addWidget(self.add_furo_btn)
        self.furos_table = QTableWidget()
        self.furos_table.setColumnCount(4)
        self.furos_table.setHorizontalHeaderLabels(["Diâmetro", "Pos X", "Pos Y", "Ação"])
        man_layout.addWidget(self.furos_table)
        man_group.setLayout(man_layout)
        furos_main_layout.addWidget(man_group)
        furos_main_group.setLayout(furos_main_layout)
        top_h_layout.addWidget(furos_main_group, stretch=1)
        main_layout.addLayout(top_h_layout)
        self.add_piece_btn = QPushButton("Adicionar Peça à Lista de Produção")
        self.add_piece_btn.clicked.connect(self.add_manual_piece)
        main_layout.addWidget(self.add_piece_btn)

        list_group = QGroupBox("4. Lista de Peças para Produção")
        list_layout = QVBoxLayout()
        self.pieces_table = QTableWidget()
        self.table_headers = [col.replace('_', ' ').title() for col in self.colunas_df] + ["Ação"]
        self.pieces_table.setColumnCount(len(self.table_headers))
        self.pieces_table.setHorizontalHeaderLabels(self.table_headers)
        list_layout.addWidget(self.pieces_table)
        self.process_btn = QPushButton("Gerar PDFs da Lista Acima")
        self.process_btn.clicked.connect(self.process_data)
        self.process_btn.setEnabled(False)
        list_layout.addWidget(self.process_btn)
        list_group.setLayout(list_layout)
        main_layout.addWidget(list_group, stretch=1)
        
        self.progress_bar = QProgressBar(); self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        log_group = QGroupBox("Log de Execução")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit(); self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)
        
        self.statusBar().showMessage("Pronto")
        self.forma_combo.currentTextChanged.connect(self.update_dimension_fields)
        self.update_dimension_fields(self.forma_combo.currentText())


    # --- FUNÇÃO ATUALIZADA PARA ADICIONAR O BOTÃO DE EDITAR ---
    def update_table_display(self):
        combined_df = pd.concat([self.excel_df, self.manual_df], ignore_index=True)
        self.pieces_table.setRowCount(0)
        
        if combined_df.empty:
            self.process_btn.setEnabled(False)
            return

        self.pieces_table.setRowCount(len(combined_df))
        for i, row in combined_df.iterrows():
            # Preenche as colunas de dados
            for j, col in enumerate(self.colunas_df):
                value = row[col]
                if col == 'furos':
                    display_value = f"{len(value)} Furo(s)" if isinstance(value, list) and value else '-'
                else:
                    display_value = '-' if pd.isna(value) or value == 0 else str(value)
                self.pieces_table.setItem(i, j, QTableWidgetItem(display_value))
            
            # --- NOVO: Cria um widget com layout para os botões de ação ---
            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(5, 0, 5, 0) # Deixa os botões mais juntos

            edit_btn = QPushButton("Editar")
            edit_btn.clicked.connect(lambda _, row=i: self.edit_row(row))
            
            delete_btn = QPushButton("Excluir")
            delete_btn.clicked.connect(lambda _, row=i: self.delete_row(row))

            action_layout.addWidget(edit_btn)
            action_layout.addWidget(delete_btn)
            
            self.pieces_table.setCellWidget(i, len(self.colunas_df), action_widget)
        
        self.pieces_table.resizeColumnsToContents()
        self.process_btn.setEnabled(True)

    # --- NOVA FUNÇÃO PARA CARREGAR OS DADOS PARA EDIÇÃO ---
    def edit_row(self, row_index):
        # 1. Identificar a peça correta e obter seus dados
        len_excel = len(self.excel_df)
        if row_index < len_excel:
            piece_data = self.excel_df.iloc[row_index]
            piece_name = piece_data['nome_arquivo']
        else:
            manual_index = row_index - len_excel
            piece_data = self.manual_df.iloc[manual_index]
            piece_name = piece_data['nome_arquivo']

        # 2. Popular os campos do formulário com os dados da peça
        self.nome_input.setText(str(piece_data.get('nome_arquivo', '')))
        self.espessura_input.setText(str(piece_data.get('espessura', '')))
        self.qtd_input.setText(str(piece_data.get('qtd', '')))
        
        shape = piece_data.get('forma', 'rectangle')
        index = self.forma_combo.findText(shape, Qt.MatchFixedString)
        if index >= 0:
            self.forma_combo.setCurrentIndex(index)

        self.largura_input.setText(str(piece_data.get('largura', '')))
        self.altura_input.setText(str(piece_data.get('altura', '')))
        self.diametro_input.setText(str(piece_data.get('diametro', '')))
        self.rt_base_input.setText(str(piece_data.get('rt_base', '')))
        self.rt_height_input.setText(str(piece_data.get('rt_height', '')))
        
        # 3. Popular a lista de furos temporária e atualizar a tabela de furos
        holes = piece_data.get('furos', [])
        self.furos_atuais = holes.copy() if isinstance(holes, list) else []
        self.update_furos_table()

        # 4. Remover a peça original da lista (ela será readicionada ao salvar)
        if row_index < len_excel:
            self.excel_df = self.excel_df.drop(self.excel_df.index[row_index]).reset_index(drop=True)
        else:
            self.manual_df = self.manual_df.drop(self.manual_df.index[manual_index]).reset_index(drop=True)
        
        self.log_text.append(f"Peça '{piece_name}' carregada para edição. Adicione-a novamente após modificar.")
        self.update_table_display()


    # O resto das funções (add_manual_piece, delete_row, etc.) permanece
    # o mesmo. Para garantir, estou incluindo todas abaixo.
    def replicate_holes(self):
        try:
            forma = self.forma_combo.currentText()
            if forma != 'rectangle':
                QMessageBox.warning(self, "Função Indisponível", "A replicação de furos só está disponível para a forma 'Retângulo'.")
                return
            largura = float(self.largura_input.text().replace(',', '.')); altura = float(self.altura_input.text().replace(',', '.'))
            diam = float(self.rep_diam_input.text().replace(',', '.')); offset = float(self.rep_offset_input.text().replace(',', '.'))
            if diam <= 0 or offset <= 0: QMessageBox.warning(self, "Valores Inválidos", "Diâmetro e offset devem ser maiores que zero."); return
            if (offset * 2) >= largura or (offset * 2) >= altura: QMessageBox.warning(self, "Offset Inválido", "O dobro do offset não pode ser maior que a largura ou altura da peça."); return
            furos_replicados = [{'diam': diam, 'x': offset, 'y': offset}, {'diam': diam, 'x': largura - offset, 'y': offset}, {'diam': diam, 'x': largura - offset, 'y': altura - offset}, {'diam': diam, 'x': offset, 'y': altura - offset}]
            self.furos_atuais.extend(furos_replicados); self.update_furos_table(); self.log_text.append(f"4 furos de Ø{diam} com offset de {offset} replicados.")
        except ValueError: QMessageBox.critical(self, "Erro de Valor", "Verifique se a Largura, Altura, Diâmetro e Offset estão preenchidos com números válidos.")
    def update_dimension_fields(self, shape):
        shape = shape.lower()
        for widget in self.largura_row + self.altura_row + self.diametro_row + self.rt_base_row + self.rt_height_row: widget.setVisible(False)
        is_rectangle = (shape == 'rectangle'); self.rep_group.setEnabled(is_rectangle)
        if is_rectangle:
            for widget in self.largura_row + self.altura_row: widget.setVisible(True)
        elif shape == 'circle':
            for widget in self.diametro_row: widget.setVisible(True)
        elif shape == 'right_triangle':
            for widget in self.rt_base_row + self.rt_height_row: widget.setVisible(True)
    def add_furo_temp(self):
        try:
            diam = float(self.diametro_furo_input.text().replace(',', '.')); pos_x = float(self.pos_x_input.text().replace(',', '.')); pos_y = float(self.pos_y_input.text().replace(',', '.'))
            if diam <= 0: QMessageBox.warning(self, "Valor Inválido", "O diâmetro do furo deve ser maior que zero."); return
            self.furos_atuais.append({'diam': diam, 'x': pos_x, 'y': pos_y}); self.update_furos_table()
            self.diametro_furo_input.clear(); self.pos_x_input.clear(); self.pos_y_input.clear()
        except ValueError: QMessageBox.critical(self, "Erro de Valor", "Os campos de furo (diâmetro, x, y) devem ser números válidos.")
    def update_furos_table(self):
        self.furos_table.setRowCount(0)
        if not self.furos_atuais: return
        self.furos_table.setRowCount(len(self.furos_atuais))
        for i, furo in enumerate(self.furos_atuais):
            self.furos_table.setItem(i, 0, QTableWidgetItem(str(furo['diam']))); self.furos_table.setItem(i, 1, QTableWidgetItem(str(furo['x']))); self.furos_table.setItem(i, 2, QTableWidgetItem(str(furo['y'])))
            delete_btn = QPushButton("Excluir"); delete_btn.clicked.connect(lambda _, row=i: self.delete_furo_temp(row)); self.furos_table.setCellWidget(i, 3, delete_btn)
        self.furos_table.resizeColumnsToContents()
    def delete_furo_temp(self, row_index):
        del self.furos_atuais[row_index]; self.update_furos_table()
    def add_manual_piece(self):
        try:
            nome = self.nome_input.text().strip()
            if not nome: QMessageBox.warning(self, "Campo Obrigatório", "O 'Nome/ID da Peça' é obrigatório."); return
            new_piece = {col: 0 if col != 'furos' else [] for col in self.colunas_df}
            new_piece['nome_arquivo'] = nome; new_piece['forma'] = self.forma_combo.currentText()
            new_piece['espessura'] = float(self.espessura_input.text().replace(',', '.')); new_piece['qtd'] = int(self.qtd_input.text())
            new_piece['furos'] = self.furos_atuais.copy()
            shape = new_piece['forma']
            if shape == 'rectangle': new_piece['largura'] = float(self.largura_input.text().replace(',', '.')); new_piece['altura'] = float(self.altura_input.text().replace(',', '.'))
            elif shape == 'circle': new_piece['diametro'] = float(self.diametro_input.text().replace(',', '.'))
            elif shape == 'right_triangle': new_piece['rt_base'] = float(self.rt_base_input.text().replace(',', '.')); new_piece['rt_height'] = float(self.rt_height_input.text().replace(',', '.'))
            new_row_df = pd.DataFrame([new_piece])
            self.manual_df = pd.concat([self.manual_df, new_row_df], ignore_index=True)
            self.log_text.append(f"Peça '{nome}' com {len(self.furos_atuais)} furos adicionada/atualizada.")
            for field in [self.nome_input, self.espessura_input, self.qtd_input, self.largura_input, self.altura_input, self.diametro_input, self.rt_base_input, self.rt_height_input]: field.clear()
            self.furos_atuais = []; self.update_furos_table(); self.update_table_display()
        except ValueError: QMessageBox.critical(self, "Erro de Valor", "Verifique se os campos da peça estão preenchidos com números válidos.")
        except Exception as e: QMessageBox.critical(self, "Erro Inesperado", f"Ocorreu um erro: {e}")
    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar Planilha Excel", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            try:
                self.log_text.append(f"Lendo arquivo: {file_path}")
                df = pd.read_excel(file_path, header=0, decimal=','); df.columns = df.columns.str.strip().str.lower(); df = df.loc[:, ~df.columns.duplicated()]
                for col in self.colunas_df:
                    if col not in df.columns: df[col] = pd.NA
                if 'furos' in df.columns: df['furos'] = df['furos'].apply(lambda x: x if isinstance(x, list) else [])
                self.excel_df = df[self.colunas_df]
                self.file_label.setText(f"Planilha carregada: {os.path.basename(file_path)}"); self.log_text.append("Planilha lida e processada com sucesso.")
                self.clear_excel_btn.setEnabled(True); self.update_table_display()
            except Exception as e: QMessageBox.critical(self, "Erro de Leitura", f"Não foi possível ler o arquivo Excel.\nErro: {e}"); self.log_text.append(f"Falha ao ler o arquivo: {e}")
    def clear_excel_data(self):
        self.excel_df = pd.DataFrame(columns=self.colunas_df); self.file_label.setText("Nenhuma planilha selecionada")
        self.clear_excel_btn.setEnabled(False); self.log_text.append("Dados da planilha removidos da lista."); self.update_table_display()
    def delete_row(self, row_index):
        len_excel = len(self.excel_df)
        if row_index < len_excel: piece_name = self.excel_df.iloc[row_index]['nome_arquivo']; self.excel_df = self.excel_df.drop(self.excel_df.index[row_index]).reset_index(drop=True)
        else: manual_index = row_index - len_excel; piece_name = self.manual_df.iloc[manual_index]['nome_arquivo']; self.manual_df = self.manual_df.drop(self.manual_df.index[manual_index]).reset_index(drop=True)
        self.log_text.append(f"Peça '{piece_name}' removida da lista."); self.update_table_display()
    def process_data(self):
        combined_df = pd.concat([self.excel_df, self.manual_df], ignore_index=True)
        if combined_df.empty: QMessageBox.warning(self, "Aviso", "A lista de peças está vazia."); return
        self.process_btn.setEnabled(False); self.select_file_btn.setEnabled(False); self.add_piece_btn.setEnabled(False); self.clear_excel_btn.setEnabled(False); self.add_furo_btn.setEnabled(False); self.replicate_btn.setEnabled(False)
        self.progress_bar.setVisible(True); self.progress_bar.setValue(0); self.log_text.clear()
        self.process_thread = ProcessThread(combined_df.copy()); self.process_thread.update_signal.connect(self.update_log)
        self.process_thread.progress_signal.connect(self.update_progress); self.process_thread.finished_signal.connect(self.processing_finished)
        self.process_thread.start()
    def update_log(self, message): self.log_text.append(message); self.statusBar().showMessage(message)
    def update_progress(self, value): self.progress_bar.setValue(value)
    def processing_finished(self, success, message):
        self.process_btn.setEnabled(True); self.select_file_btn.setEnabled(True); self.add_piece_btn.setEnabled(True); self.add_furo_btn.setEnabled(True); self.replicate_btn.setEnabled(True)
        if not self.excel_df.empty: self.clear_excel_btn.setEnabled(True)
        if success: QMessageBox.information(self, "Sucesso", message)
        else: QMessageBox.critical(self, "Erro", message)
        self.statusBar().showMessage("Pronto")

def main():
    # Criar aplicação Qt
    app = QApplication(sys.argv)
    
    # Criar e mostrar janela principal
    window = MainWindow()
    window.show()
    
    # Executar loop de eventos
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()