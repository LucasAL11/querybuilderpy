import pandas as pd
from PyQt5 import QtWidgets
import sys
import math

class ExcelWizard(QtWidgets.QWizard):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Database Update Wizard")

        # Adiciona páginas com IDs únicos
        self.file_selection_page = FileSelectionPage()
        self.column_selection_page = ColumnSelectionPage()
        self.table_name_page = TableNamePage()

        self.addPage(self.file_selection_page)
        self.addPage(self.column_selection_page)
        self.addPage(self.table_name_page)

class FileSelectionPage(QtWidgets.QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle("Select Excel File")
        self.file_path = None

        layout = QtWidgets.QVBoxLayout(self)

        self.button = QtWidgets.QPushButton("Choose Excel File")
        self.button.clicked.connect(self.select_file)
        layout.addWidget(self.button)

    def select_file(self):
        options = QtWidgets.QFileDialog.Options()
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select Excel File", "", 
            "Excel Files (*.xlsx *.xls);;All Files (*)", options=options
        )
        if file_path:
            self.file_path = file_path
            self.completeChanged.emit()

    def isComplete(self):
        return self.file_path is not None

class ColumnSelectionPage(QtWidgets.QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle("Select Columns for Update and Reference")
        self.columns = None
        self.update_checkboxes = []
        self.reference_checkboxes = []
        self.layout = QtWidgets.QVBoxLayout(self)

    def initializePage(self):
        file_path = self.wizard().file_selection_page.file_path
        df = pd.read_excel(file_path)
        self.columns = df.columns

        self.create_checkboxes("Update Columns", self.update_checkboxes)
        self.create_checkboxes("Reference Column", self.reference_checkboxes)

    def create_checkboxes(self, title, container):
        group_box = QtWidgets.QGroupBox(title)
        layout = QtWidgets.QVBoxLayout(group_box)

        for col in self.columns:
            checkbox = QtWidgets.QCheckBox(col)
            checkbox.stateChanged.connect(self.completeChanged)  # Conecta evento de mudança
            layout.addWidget(checkbox)
            container.append(checkbox)

        self.layout.addWidget(group_box)

    def selected_columns(self, checkboxes):
        return [cb.text() for cb in checkboxes if cb.isChecked()]

    def isComplete(self):
        # Verifica se há ao menos uma coluna para atualização e uma de referência
        return bool(self.selected_columns(self.update_checkboxes)) and \
               bool(self.selected_columns(self.reference_checkboxes))


class TableNamePage(QtWidgets.QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle("Finalize - Enter Table Name")

        # Layout da página
        layout = QtWidgets.QVBoxLayout(self)

        # Campo de entrada para nome da tabela
        self.table_name_input = QtWidgets.QLineEdit()
        self.table_name_input.setPlaceholderText("Enter the table name")
        self.table_name_input.textChanged.connect(self.on_text_changed)  # Detecta mudanças no campo
        layout.addWidget(self.table_name_input)

        # Mensagem adicional
        self.info_label = QtWidgets.QLabel("Click 'Finish' after entering a valid table name.")
        layout.addWidget(self.info_label)

        self.setLayout(layout)

    def on_text_changed(self):
        """Chama a função para emitir o sinal de mudança de estado."""
        self.completeChanged.emit()  # Garante que o wizard atualize o estado do botão Finish

    def isComplete(self):
        """Verifica se o nome da tabela foi preenchido corretamente."""
        return bool(self.table_name_input.text().strip())

def dividir_dataframe(df, max_linhas):
    num_partes = math.ceil(len(df) / max_linhas)
    return [df.iloc[i * max_linhas : (i + 1) * max_linhas] for i in range(num_partes)]

def gerar_queries(df, update_cols, ref_col, table_name):
    MAX_LINHAS_POR_UPDATE = 1000
    update_parts = dividir_dataframe(df, MAX_LINHAS_POR_UPDATE)
    queries = []

    for parte in update_parts:
        set_statements = [
            f"{col} = CASE {ref_col} " +
            " ".join([f"WHEN '{row[ref_col]}' THEN '{row[col]}'" for _, row in parte.iterrows()]) +
            f" ELSE {col} END"
            for col in update_cols
        ]

        where_clause = ", ".join([f"'{row[ref_col]}'" for _, row in parte.iterrows()])
        query = f"""
        UPDATE {table_name}
        SET {', '.join(set_statements)}
        WHERE {ref_col} IN ({where_clause});
        """
        queries.append(query)

    return queries

def main():
    app = QtWidgets.QApplication(sys.argv)
    wizard = ExcelWizard()

    if wizard.exec_() == QtWidgets.QWizard.Accepted:
        file_path = wizard.file_selection_page.file_path
        df = pd.read_excel(file_path)

        update_cols = wizard.column_selection_page.selected_columns(
            wizard.column_selection_page.update_checkboxes
        )
        ref_col = wizard.column_selection_page.selected_columns(
            wizard.column_selection_page.reference_checkboxes
        )[0]
        table_name = wizard.table_name_page.table_name_input.text()

        queries = gerar_queries(df, update_cols, ref_col, table_name)

        # Salva as queries em arquivos
        for idx, query in enumerate(queries):
            file_name = f"update_query_part_{idx + 1}.sql"
            with open(file_name, "w") as file:
                file.write(query.strip())
            print(f"Query {idx + 1} saved as {file_name}.")

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
