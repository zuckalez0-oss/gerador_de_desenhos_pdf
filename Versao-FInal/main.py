# main.py

import sys
import os
import json
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QTextEdit, 
                             QFileDialog, QProgressBar, QMessageBox, QGroupBox,
                             QFormLayout, QLineEdit, QComboBox, QTableWidget, 
                             QTableWidgetItem, QDialog, QInputDialog)
from PyQt5.QtCore import Qt

# <<< IMPORTAÇÕES DAS CLASSES ENCAPSULADAS >>>
from code_manager import CodeGenerator
from history_manager import HistoryManager
from history_dialog import HistoryDialog
from processing import ProcessThread

# =============================================================================
# CLASSE PRINCIPAL DA INTERFACE GRÁFICA
# =============================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerador de Desenhos Técnicos e DXF")
        self.setGeometry(100, 100, 1100, 850)
        
        # Instancia as classes dos módulos importados
        self.code_generator = CodeGenerator()
        self.history_manager = HistoryManager()
        
        # Variáveis de estado da aplicação
        self.colunas_df = ['nome_arquivo', 'forma', 'espessura', 'qtd', 'largura', 'altura', 'diametro', 'rt_base', 'rt_height', 'trapezoid_large_base', 'trapezoid_small_base', 'trapezoid_height', 'furos']
        self.manual_df = pd.DataFrame(columns=self.colunas_df)
        self.excel_df = pd.DataFrame(columns=self.colunas_df)
        self.furos_atuais = []
        self.project_directory = None

        # =====================================================================
        # <<< INÍCIO DA CONSTRUÇÃO DA INTERFACE GRÁFICA (UI) >>>
        # Esta é a parte que provavelmente estava faltando.
        # =====================================================================
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        top_h_layout = QHBoxLayout()
        left_v_layout = QVBoxLayout()
        
        # --- Grupo 1: Projeto ---
        project_group = QGroupBox("1. Projeto")
        project_layout = QVBoxLayout()
        self.start_project_btn = QPushButton("Iniciar Novo Projeto...")
        self.start_project_btn.setStyleSheet("background-color: #007bff; color: white; font-weight: bold;")
        self.history_btn = QPushButton("Ver Histórico de Projetos")
        project_layout.addWidget(self.start_project_btn)
        project_layout.addWidget(self.history_btn)
        project_group.setLayout(project_layout)
        left_v_layout.addWidget(project_group)
        
        # --- Grupo 2: Carregar Planilha ---
        file_group = QGroupBox("2. Carregar Planilha (Opcional)")
        file_layout = QVBoxLayout()
        self.file_label = QLabel("Nenhum projeto ativo.")
        file_button_layout = QHBoxLayout()
        self.select_file_btn = QPushButton("Selecionar Planilha")
        self.clear_excel_btn = QPushButton("Limpar Planilha")
        file_button_layout.addWidget(self.select_file_btn)
        file_button_layout.addWidget(self.clear_excel_btn)
        file_layout.addWidget(self.file_label)
        file_layout.addLayout(file_button_layout)
        file_group.setLayout(file_layout)
        left_v_layout.addWidget(file_group)

        # --- Grupo 3: Informações da Peça ---
        manual_group = QGroupBox("3. Informações da Peça")
        manual_layout = QFormLayout()
        self.projeto_input = QLineEdit()
        self.projeto_input.setReadOnly(True)
        manual_layout.addRow("Nº do Projeto Ativo:", self.projeto_input)
        self.nome_input = QLineEdit()
        self.generate_code_btn = QPushButton("Gerar Código")
        name_layout = QHBoxLayout()
        name_layout.addWidget(self.nome_input)
        name_layout.addWidget(self.generate_code_btn)
        manual_layout.addRow("Nome/ID da Peça:", name_layout)
        self.forma_combo = QComboBox()
        self.forma_combo.addItems(['rectangle', 'circle', 'right_triangle', 'trapezoid'])
        self.espessura_input, self.qtd_input = QLineEdit(), QLineEdit()
        manual_layout.addRow("Forma:", self.forma_combo)
        manual_layout.addRow("Espessura (mm):", self.espessura_input)
        manual_layout.addRow("Quantidade:", self.qtd_input)
        self.largura_input, self.altura_input = QLineEdit(), QLineEdit()
        self.diametro_input, self.rt_base_input, self.rt_height_input = QLineEdit(), QLineEdit(), QLineEdit()
        self.trapezoid_large_base_input, self.trapezoid_small_base_input, self.trapezoid_height_input = QLineEdit(), QLineEdit(), QLineEdit()
        self.largura_row = [QLabel("Largura:"), self.largura_input]; manual_layout.addRow(*self.largura_row)
        self.altura_row = [QLabel("Altura:"), self.altura_input]; manual_layout.addRow(*self.altura_row)
        self.diametro_row = [QLabel("Diâmetro:"), self.diametro_input]; manual_layout.addRow(*self.diametro_row)
        self.rt_base_row = [QLabel("Base Triângulo:"), self.rt_base_input]; manual_layout.addRow(*self.rt_base_row)
        self.rt_height_row = [QLabel("Altura Triângulo:"), self.rt_height_input]; manual_layout.addRow(*self.rt_height_row)
        self.trap_large_base_row = [QLabel("Base Maior:"), self.trapezoid_large_base_input]; manual_layout.addRow(*self.trap_large_base_row)
        self.trap_small_base_row = [QLabel("Base Menor:"), self.trapezoid_small_base_input]; manual_layout.addRow(*self.trap_small_base_row)
        self.trap_height_row = [QLabel("Altura:"), self.trapezoid_height_input]; manual_layout.addRow(*self.trap_height_row)
        manual_group.setLayout(manual_layout)
        left_v_layout.addWidget(manual_group)
        left_v_layout.addStretch()
        top_h_layout.addLayout(left_v_layout)
        
        # --- Grupo 4: Furos ---
        furos_main_group = QGroupBox("4. Adicionar Furos")
        furos_main_layout = QVBoxLayout()
        self.rep_group = QGroupBox("Furação Rápida")
        rep_layout = QFormLayout()
        self.rep_diam_input, self.rep_offset_input = QLineEdit(), QLineEdit()
        rep_layout.addRow("Diâmetro Furos:", self.rep_diam_input)
        rep_layout.addRow("Offset Borda:", self.rep_offset_input)
        self.replicate_btn = QPushButton("Replicar Furos")
        rep_layout.addRow(self.replicate_btn)
        self.rep_group.setLayout(rep_layout)
        furos_main_layout.addWidget(self.rep_group)
        man_group = QGroupBox("Furos Manuais")
        man_layout = QVBoxLayout()
        man_form_layout = QFormLayout()
        self.diametro_furo_input, self.pos_x_input, self.pos_y_input = QLineEdit(), QLineEdit(), QLineEdit()
        man_form_layout.addRow("Diâmetro:", self.diametro_furo_input)
        man_form_layout.addRow("Posição X:", self.pos_x_input)
        man_form_layout.addRow("Posição Y:", self.pos_y_input)
        self.add_furo_btn = QPushButton("Adicionar Furo Manual")
        man_layout.addLayout(man_form_layout)
        man_layout.addWidget(self.add_furo_btn)
        self.furos_table = QTableWidget(0, 4)
        self.furos_table.setHorizontalHeaderLabels(["Diâmetro", "Pos X", "Pos Y", "Ação"])
        man_layout.addWidget(self.furos_table)
        man_group.setLayout(man_layout)
        furos_main_layout.addWidget(man_group)
        furos_main_group.setLayout(furos_main_layout)
        top_h_layout.addWidget(furos_main_group, stretch=1)
        main_layout.addLayout(top_h_layout)

        self.add_piece_btn = QPushButton("Adicionar Peça à Lista")
        main_layout.addWidget(self.add_piece_btn)

        # --- Grupo 5: Lista de Peças ---
        list_group = QGroupBox("5. Lista de Peças para Produção")
        list_layout = QVBoxLayout()
        self.pieces_table = QTableWidget()
        self.table_headers = [col.replace('_', ' ').title() for col in self.colunas_df] + ["Ações"]
        self.pieces_table.setColumnCount(len(self.table_headers))
        self.pieces_table.setHorizontalHeaderLabels(self.table_headers)
        list_layout.addWidget(self.pieces_table)
        self.dir_label = QLabel("Nenhum projeto ativo. Inicie um novo projeto.")
        self.dir_label.setStyleSheet("font-style: italic; color: grey;")
        list_layout.addWidget(self.dir_label)
        process_buttons_layout = QHBoxLayout()
        self.conclude_project_btn = QPushButton("Projeto Concluído")
        self.conclude_project_btn.setStyleSheet("background-color: #28a745; color: white; font-weight: bold;")
        self.export_excel_btn = QPushButton("Exportar para Excel")
        self.process_pdf_btn, self.process_dxf_btn, self.process_all_btn = QPushButton("Gerar PDFs"), QPushButton("Gerar DXFs"), QPushButton("Gerar PDFs e DXFs")
        process_buttons_layout.addWidget(self.export_excel_btn)
        process_buttons_layout.addWidget(self.conclude_project_btn)
        process_buttons_layout.addStretch()
        process_buttons_layout.addWidget(self.process_pdf_btn)
        process_buttons_layout.addWidget(self.process_dxf_btn)
        process_buttons_layout.addWidget(self.process_all_btn)
        list_layout.addLayout(process_buttons_layout)
        list_group.setLayout(list_layout)
        main_layout.addWidget(list_group, stretch=1)
        
        # --- Barra de Progresso e Log ---
        self.progress_bar = QProgressBar()
        main_layout.addWidget(self.progress_bar)
        log_group = QGroupBox("Log de Execução")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)
        self.statusBar().showMessage("Pronto")
        
        # --- Conexões de Sinais e Slots (Eventos) ---
        self.start_project_btn.clicked.connect(self.start_new_project)
        self.history_btn.clicked.connect(self.show_history_dialog)
        self.select_file_btn.clicked.connect(self.select_file)
        self.clear_excel_btn.clicked.connect(self.clear_excel_data)
        self.generate_code_btn.clicked.connect(self.generate_piece_code)
        self.add_piece_btn.clicked.connect(self.add_manual_piece)
        self.forma_combo.currentTextChanged.connect(self.update_dimension_fields)
        self.replicate_btn.clicked.connect(self.replicate_holes)
        self.add_furo_btn.clicked.connect(self.add_furo_temp)
        self.process_pdf_btn.clicked.connect(self.start_pdf_generation)
        self.process_dxf_btn.clicked.connect(self.start_dxf_generation)
        self.process_all_btn.clicked.connect(self.start_all_generation)
        self.conclude_project_btn.clicked.connect(self.conclude_project)
        self.export_excel_btn.clicked.connect(self.export_project_to_excel)
        
        # --- Estado Inicial da UI ---
        self.set_initial_button_state()
        self.update_dimension_fields(self.forma_combo.currentText())

    # =====================================================================
    # <<< INÍCIO DAS FUNÇÕES (MÉTODOS) DA CLASSE MainWindow >>>
    # Todos os métodos que definem o comportamento da aplicação
    # =====================================================================

    def start_new_project(self):
        parent_dir = QFileDialog.getExistingDirectory(self, "Selecione a Pasta Principal para o Novo Projeto")
        if not parent_dir: return
        project_name, ok = QInputDialog.getText(self, "Novo Projeto", "Digite o nome ou número do novo projeto:")
        if ok and project_name:
            project_path = os.path.join(parent_dir, project_name)
            if os.path.exists(project_path):
                reply = QMessageBox.question(self, 'Diretório Existente', f"A pasta '{project_name}' já existe.\nDeseja usá-la como o diretório do projeto ativo?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.No: return
            else:
                try: os.makedirs(project_path)
                except OSError as e: QMessageBox.critical(self, "Erro ao Criar Pasta", f"Não foi possível criar o diretório do projeto:\n{e}"); return
            self._clear_session(clear_project_number=True)
            self.project_directory = project_path
            self.projeto_input.setText(project_name)
            self.dir_label.setText(f"Projeto Ativo: {self.project_directory}")
            self.dir_label.setStyleSheet("font-style: normal; color: black;")
            self.log_text.append(f"\n--- NOVO PROJETO INICIADO: {project_name} ---")
            self.log_text.append(f"Arquivos serão salvos em: {self.project_directory}")
            self.set_initial_button_state()

    def set_initial_button_state(self):
        is_project_active = self.project_directory is not None
        self.start_project_btn.setEnabled(True); self.history_btn.setEnabled(True)
        self.select_file_btn.setEnabled(is_project_active); self.clear_excel_btn.setEnabled(is_project_active and not self.excel_df.empty)
        self.generate_code_btn.setEnabled(is_project_active); self.add_piece_btn.setEnabled(is_project_active)
        self.replicate_btn.setEnabled(is_project_active); self.add_furo_btn.setEnabled(is_project_active)
        has_items = not (self.excel_df.empty and self.manual_df.empty)
        self.process_pdf_btn.setEnabled(is_project_active and has_items); self.process_dxf_btn.setEnabled(is_project_active and has_items)
        self.process_all_btn.setEnabled(is_project_active and has_items); self.conclude_project_btn.setEnabled(is_project_active and has_items)
        self.export_excel_btn.setEnabled(is_project_active and has_items)
        self.progress_bar.setVisible(False)

    def show_history_dialog(self):
        dialog = HistoryDialog(self.history_manager, self)
        if dialog.exec_() == QDialog.Accepted:
            loaded_pieces = dialog.loaded_project_data
            if loaded_pieces:
                # Tenta obter o número do projeto a partir do primeiro item da lista
                project_number_loaded = loaded_pieces[0].get('project_number') if loaded_pieces and 'project_number' in loaded_pieces[0] else dialog.project_list_widget.currentItem().text()
                self.start_new_project_from_history(project_number_loaded, loaded_pieces)
    
    def start_new_project_from_history(self, project_name, pieces_data):
        parent_dir = QFileDialog.getExistingDirectory(self, f"Selecione uma pasta para o projeto '{project_name}'")
        if not parent_dir: return
        project_path = os.path.join(parent_dir, project_name)
        os.makedirs(project_path, exist_ok=True)
        self._clear_session(clear_project_number=True)
        self.project_directory = project_path
        self.projeto_input.setText(project_name)
        self.excel_df = pd.DataFrame(columns=self.colunas_df)
        self.manual_df = pd.DataFrame(pieces_data)
        self.dir_label.setText(f"Projeto Ativo: {self.project_directory}"); self.dir_label.setStyleSheet("font-style: normal; color: black;")
        self.log_text.append(f"\n--- PROJETO DO HISTÓRICO CARREGADO: {project_name} ---")
        self.update_table_display()
        self.set_initial_button_state()

    def start_pdf_generation(self): self.start_processing(generate_pdf=True, generate_dxf=False)
    def start_dxf_generation(self): self.start_processing(generate_pdf=False, generate_dxf=True)
    def start_all_generation(self): self.start_processing(generate_pdf=True, generate_dxf=True)

    def start_processing(self, generate_pdf, generate_dxf):
        if not self.project_directory:
            QMessageBox.warning(self, "Nenhum Projeto Ativo", "Inicie um novo projeto antes de gerar arquivos."); return
        combined_df = pd.concat([self.excel_df, self.manual_df], ignore_index=True)
        if combined_df.empty: QMessageBox.warning(self, "Aviso", "A lista de peças está vazia."); return
        self.set_buttons_enabled_on_process(False)
        self.progress_bar.setVisible(True); self.progress_bar.setValue(0); self.log_text.clear()
        self.process_thread = ProcessThread(combined_df.copy(), generate_pdf, generate_dxf, self.project_directory)
        self.process_thread.update_signal.connect(self.log_text.append)
        self.process_thread.progress_signal.connect(self.progress_bar.setValue)
        self.process_thread.finished_signal.connect(self.processing_finished)
        self.process_thread.start()

    def processing_finished(self, success, message):
        self.set_buttons_enabled_on_process(True); self.progress_bar.setVisible(False)
        msgBox = QMessageBox.information if success else QMessageBox.critical
        msgBox(self, "Concluído" if success else "Erro", message); self.statusBar().showMessage("Pronto")
    
    def conclude_project(self):
        project_number = self.projeto_input.text().strip()
        if not project_number:
            QMessageBox.warning(self, "Projeto sem Número", "O projeto ativo não tem um número definido.")
            return
        reply = QMessageBox.question(self, 'Concluir Projeto', f"Deseja salvar e concluir o projeto '{project_number}'?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            combined_df = pd.concat([self.excel_df, self.manual_df], ignore_index=True)
            if not combined_df.empty:
                combined_df['project_number'] = project_number
                self.history_manager.save_project(project_number, combined_df)
                self.log_text.append(f"Projeto '{project_number}' salvo no histórico.")
            self._clear_session(clear_project_number=True)
            self.project_directory = None
            self.dir_label.setText("Nenhum projeto ativo. Inicie um novo projeto."); self.dir_label.setStyleSheet("font-style: italic; color: grey;")
            self.set_initial_button_state()
            self.log_text.append(f"\n--- PROJETO '{project_number}' CONCLUÍDO ---")
    
    def export_project_to_excel(self):
        project_number = self.projeto_input.text().strip()
        if not project_number:
            QMessageBox.warning(self, "Nenhum Projeto Ativo", "Inicie um novo projeto para poder exportá-lo.")
            return

        combined_df = pd.concat([self.excel_df, self.manual_df], ignore_index=True)
        if combined_df.empty:
            QMessageBox.warning(self, "Lista Vazia", "Não há peças na lista para exportar.")
            return

        default_filename = os.path.join(self.project_directory, f"Resumo_Projeto_{project_number}.xlsx")
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Resumo do Projeto", default_filename, "Excel Files (*.xlsx)")

        if not save_path:
            return

        try:
            export_df = pd.DataFrame()
            export_df['Numero do Projeto'] = [project_number] * len(combined_df)
            export_df['Nome da Peça'] = combined_df['nome_arquivo']
            export_df['Forma'] = combined_df['forma']
            export_df['Espessura (mm)'] = combined_df['espessura']
            export_df['Quantidade'] = combined_df['qtd']
            export_df['Largura (mm)'] = combined_df['largura']
            export_df['Altura (mm)'] = combined_df['altura']
            export_df['Diâmetro (mm)'] = combined_df['diametro']
            export_df['Base Triângulo (mm)'] = combined_df['rt_base']
            export_df['Altura Triângulo (mm)'] = combined_df['rt_height']
            export_df['Base Maior Trapézio (mm)'] = combined_df['trapezoid_large_base']
            export_df['Base Menor Trapézio (mm)'] = combined_df['trapezoid_small_base']
            export_df['Altura Trapézio (mm)'] = combined_df['trapezoid_height']
            
            furos_data = combined_df['furos'].fillna('[]').apply(
                lambda f: f if isinstance(f, list) else json.loads(f.replace("'", "\""))
            )
            export_df['Quantidade de Furos'] = furos_data.apply(len)
            
            def get_unique_diameters(furos_list):
                if not furos_list: return ""
                diameters = sorted(list(set([f.get('diam', 0) for f in furos_list])))
                return ", ".join(map(str, diameters))

            export_df['Diâmetros dos Furos (mm)'] = furos_data.apply(get_unique_diameters)
            export_df['Detalhes dos Furos (JSON)'] = furos_data.apply(json.dumps)
            
            export_df.replace(0, pd.NA, inplace=True)
            export_df.dropna(axis=1, how='all', inplace=True)

            export_df.to_excel(save_path, index=False, engine='openpyxl')
            
            self.log_text.append(f"Resumo do projeto salvo com sucesso em: {save_path}")
            QMessageBox.information(self, "Sucesso", f"O arquivo Excel foi salvo com sucesso em:\n{save_path}")

        except Exception as e:
            self.log_text.append(f"ERRO ao exportar para Excel: {e}")
            QMessageBox.critical(self, "Erro na Exportação", f"Ocorreu um erro ao salvar o arquivo:\n{e}")

    def _clear_session(self, clear_project_number=False):
        fields_to_clear = [self.nome_input, self.espessura_input, self.qtd_input, self.largura_input, self.altura_input, self.diametro_input, self.rt_base_input, self.rt_height_input, self.trapezoid_large_base_input, self.trapezoid_small_base_input, self.trapezoid_height_input, self.rep_diam_input, self.rep_offset_input, self.diametro_furo_input, self.pos_x_input, self.pos_y_input]
        if clear_project_number:
            fields_to_clear.append(self.projeto_input)
        for field in fields_to_clear:
            field.clear()
        self.furos_atuais = []
        self.update_furos_table()
        self.file_label.setText("Nenhum projeto ativo.")
        if clear_project_number: 
            self.excel_df = pd.DataFrame(columns=self.colunas_df)
            self.manual_df = pd.DataFrame(columns=self.colunas_df)
            self.update_table_display()

    def set_buttons_enabled_on_process(self, enabled):
        is_project_active = self.project_directory is not None
        self.start_project_btn.setEnabled(enabled)
        self.history_btn.setEnabled(enabled)
        self.select_file_btn.setEnabled(enabled and is_project_active)
        self.clear_excel_btn.setEnabled(enabled and is_project_active and not self.excel_df.empty)
        self.generate_code_btn.setEnabled(enabled and is_project_active)
        self.add_piece_btn.setEnabled(enabled and is_project_active)
        self.replicate_btn.setEnabled(enabled and is_project_active)
        self.add_furo_btn.setEnabled(enabled and is_project_active)
        has_items = not (self.excel_df.empty and self.manual_df.empty)
        self.process_pdf_btn.setEnabled(enabled and is_project_active and has_items)
        self.process_dxf_btn.setEnabled(enabled and is_project_active and has_items)
        self.process_all_btn.setEnabled(enabled and is_project_active and has_items)
        self.conclude_project_btn.setEnabled(enabled and is_project_active and has_items)
        self.export_excel_btn.setEnabled(enabled and is_project_active and has_items)
    
    def update_table_display(self):
        self.set_initial_button_state() # Garante que os botões sempre reflitam o estado atual
        combined_df = pd.concat([self.excel_df, self.manual_df], ignore_index=True)
        self.pieces_table.setRowCount(0)
        if combined_df.empty: return
        self.pieces_table.setRowCount(len(combined_df))
        for i, row in combined_df.iterrows():
            for j, col in enumerate(self.colunas_df):
                value = row[col]
                display_value = f"{len(value)} Furo(s)" if col == 'furos' and isinstance(value, list) else ('-' if pd.isna(value) or value == 0 else str(value))
                self.pieces_table.setItem(i, j, QTableWidgetItem(display_value))
            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(5, 0, 5, 0)
            edit_btn, delete_btn = QPushButton("Editar"), QPushButton("Excluir")
            edit_btn.clicked.connect(lambda _, r=i: self.edit_row(r))
            delete_btn.clicked.connect(lambda _, r=i: self.delete_row(r))
            action_layout.addWidget(edit_btn)
            action_layout.addWidget(delete_btn)
            self.pieces_table.setCellWidget(i, len(self.colunas_df), action_widget)
        self.pieces_table.resizeColumnsToContents()
    
    def edit_row(self, row_index):
        len_excel = len(self.excel_df)
        is_from_excel = row_index < len_excel
        df_source = self.excel_df if is_from_excel else self.manual_df
        local_index = row_index if is_from_excel else row_index - len_excel
        
        piece_data = df_source.iloc[local_index]
        self.nome_input.setText(str(piece_data.get('nome_arquivo', '')))
        self.espessura_input.setText(str(piece_data.get('espessura', '')))
        self.qtd_input.setText(str(piece_data.get('qtd', '')))
        
        shape = piece_data.get('forma', '')
        index = self.forma_combo.findText(shape, Qt.MatchFixedString)
        if index >= 0: self.forma_combo.setCurrentIndex(index)
        
        self.largura_input.setText(str(piece_data.get('largura', '')))
        self.altura_input.setText(str(piece_data.get('altura', '')))
        self.diametro_input.setText(str(piece_data.get('diametro', '')))
        self.rt_base_input.setText(str(piece_data.get('rt_base', '')))
        self.rt_height_input.setText(str(piece_data.get('rt_height', '')))
        self.trapezoid_large_base_input.setText(str(piece_data.get('trapezoid_large_base', '')))
        self.trapezoid_small_base_input.setText(str(piece_data.get('trapezoid_small_base', '')))
        self.trapezoid_height_input.setText(str(piece_data.get('trapezoid_height', '')))
        
        self.furos_atuais = piece_data.get('furos', []).copy() if isinstance(piece_data.get('furos'), list) else []
        self.update_furos_table()
        
        # Remove a linha do dataframe para que ela possa ser adicionada novamente
        df_source.drop(df_source.index[local_index], inplace=True)
        df_source.reset_index(drop=True, inplace=True)
        self.log_text.append(f"Peça '{piece_data['nome_arquivo']}' carregada para edição.")
        self.update_table_display()
    
    def delete_row(self, row_index):
        len_excel = len(self.excel_df)
        is_from_excel = row_index < len_excel
        df_source = self.excel_df if is_from_excel else self.manual_df
        local_index = row_index if is_from_excel else row_index - len_excel
        
        piece_name = df_source.iloc[local_index]['nome_arquivo']
        df_source.drop(df_source.index[local_index], inplace=True)
        df_source.reset_index(drop=True, inplace=True)
        self.log_text.append(f"Peça '{piece_name}' removida.")
        self.update_table_display()
    
    def generate_piece_code(self):
        project_number = self.projeto_input.text().strip()
        if not project_number: QMessageBox.warning(self, "Campo Obrigatório", "Inicie um projeto para definir o 'Nº do Projeto'."); return
        new_code = self.code_generator.generate_new_code(project_number, prefix='DES')
        if new_code: self.nome_input.setText(new_code); self.log_text.append(f"Código '{new_code}' gerado para o projeto '{project_number}'.")
    
    def add_manual_piece(self):
        try:
            nome = self.nome_input.text().strip()
            if not nome: QMessageBox.warning(self, "Campo Obrigatório", "'Nome/ID da Peça' é obrigatório."); return
            new_piece = {'furos': self.furos_atuais.copy()}
            for col in self.colunas_df:
                if col != 'furos': new_piece[col] = 0.0
            new_piece.update({'nome_arquivo': nome, 'forma': self.forma_combo.currentText()})
            fields_map = { 'espessura': self.espessura_input, 'qtd': self.qtd_input, 'largura': self.largura_input, 'altura': self.altura_input, 'diametro': self.diametro_input, 'rt_base': self.rt_base_input, 'rt_height': self.rt_height_input, 'trapezoid_large_base': self.trapezoid_large_base_input, 'trapezoid_small_base': self.trapezoid_small_base_input, 'trapezoid_height': self.trapezoid_height_input }
            for key, field in fields_map.items():
                new_piece[key] = float(field.text().replace(',', '.')) if field.text() else 0.0
            self.manual_df = pd.concat([self.manual_df, pd.DataFrame([new_piece])], ignore_index=True)
            self.log_text.append(f"Peça '{nome}' adicionada/atualizada.")
            self._clear_session(clear_project_number=False) # Não limpa o número do projeto
            self.update_table_display()
        except ValueError: QMessageBox.critical(self, "Erro de Valor", "Campos numéricos devem conter números válidos.")
    
    def select_file(self):
        if not self.project_directory:
            QMessageBox.warning(self, "Nenhum Projeto Ativo", "Inicie um projeto antes de carregar uma planilha.")
            return
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar Planilha", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            try:
                df = pd.read_excel(file_path, header=0, decimal=','); df.columns = df.columns.str.strip().str.lower()
                df = df.loc[:, ~df.columns.duplicated()]
                for col in self.colunas_df:
                    if col not in df.columns: df[col] = pd.NA
                if 'furos' in df.columns:
                    def parse_furos(x):
                        if isinstance(x, list): return x
                        if isinstance(x, str) and x.startswith('['):
                            try: return json.loads(x.replace("'", "\""))
                            except json.JSONDecodeError: return []
                        return []
                    df['furos'] = df['furos'].apply(parse_furos)
                else:
                    df['furos'] = [[] for _ in range(len(df))]
                self.excel_df = df[self.colunas_df]
                self.file_label.setText(f"Planilha: {os.path.basename(file_path)}"); self.update_table_display()
            except Exception as e: QMessageBox.critical(self, "Erro de Leitura", f"Falha ao ler o arquivo: {e}")
    
    def clear_excel_data(self):
        self.excel_df = pd.DataFrame(columns=self.colunas_df); self.file_label.setText("Nenhuma planilha selecionada"); self.update_table_display()
    
    def replicate_holes(self):
        try:
            if self.forma_combo.currentText() != 'rectangle': QMessageBox.warning(self, "Função Indisponível", "Replicação disponível apenas para Retângulos."); return
            largura, altura = float(self.largura_input.text().replace(',', '.')), float(self.altura_input.text().replace(',', '.'))
            diam, offset = float(self.rep_diam_input.text().replace(',', '.')), float(self.rep_offset_input.text().replace(',', '.'))
            if (offset * 2) >= largura or (offset * 2) >= altura: QMessageBox.warning(self, "Offset Inválido", "Offset excede as dimensões da peça."); return
            furos = [{'diam': diam, 'x': offset, 'y': offset}, {'diam': diam, 'x': largura - offset, 'y': offset}, {'diam': diam, 'x': largura - offset, 'y': altura - offset}, {'diam': diam, 'x': offset, 'y': altura - offset}]
            self.furos_atuais.extend(furos); self.update_furos_table()
        except ValueError: QMessageBox.critical(self, "Erro de Valor", "Largura, Altura, Diâmetro e Offset devem ser números válidos.")
    
    def update_dimension_fields(self, shape):
        shape = shape.lower()
        is_rect, is_circ, is_tri, is_trap = shape == 'rectangle', shape == 'circle', shape == 'right_triangle', shape == 'trapezoid'
        for w in self.largura_row + self.altura_row: w.setVisible(is_rect)
        for w in self.diametro_row: w.setVisible(is_circ)
        for w in self.rt_base_row + self.rt_height_row: w.setVisible(is_tri)
        for w in self.trap_large_base_row + self.trap_small_base_row + self.trap_height_row: w.setVisible(is_trap)
        self.rep_group.setEnabled(is_rect)
    
    def add_furo_temp(self):
        try:
            diam, pos_x, pos_y = float(self.diametro_furo_input.text().replace(',', '.')), float(self.pos_x_input.text().replace(',', '.')), float(self.pos_y_input.text().replace(',', '.'))
            if diam <= 0: QMessageBox.warning(self, "Valor Inválido", "Diâmetro do furo deve ser maior que zero."); return
            self.furos_atuais.append({'diam': diam, 'x': pos_x, 'y': pos_y}); self.update_furos_table()
            for field in [self.diametro_furo_input, self.pos_x_input, self.pos_y_input]: field.clear()
        except ValueError: QMessageBox.critical(self, "Erro de Valor", "Campos de furo devem ser números válidos.")
    
    def update_furos_table(self):
        self.furos_table.setRowCount(0); self.furos_table.setRowCount(len(self.furos_atuais))
        for i, furo in enumerate(self.furos_atuais):
            self.furos_table.setItem(i, 0, QTableWidgetItem(str(furo['diam'])))
            self.furos_table.setItem(i, 1, QTableWidgetItem(str(furo['x'])))
            self.furos_table.setItem(i, 2, QTableWidgetItem(str(furo['y'])))
            delete_btn = QPushButton("Excluir")
            delete_btn.clicked.connect(lambda _, r=i: self.delete_furo_temp(r))
            self.furos_table.setCellWidget(i, 3, delete_btn)
        self.furos_table.resizeColumnsToContents()
    
    def delete_furo_temp(self, row_index):
        del self.furos_atuais[row_index]
        self.update_furos_table()

# =============================================================================
# PONTO DE ENTRADA DA APLICAÇÃO
# =============================================================================
def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()