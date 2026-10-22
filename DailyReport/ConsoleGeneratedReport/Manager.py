import shutil
import subprocess
import sys

import pandas as pd
from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QTableWidget, QTableWidgetItem,
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QMessageBox, QStyledItemDelegate, QComboBox,
)

FICHEIRO = "Data/Data.xlsx"
FOLHAS = ["Tecnic", "Machines", "Errors"]   # exatamente como no Excel
DAILY_CLEAN = "daily_clean.py"

# ---- Configuração: ajusta aos nomes reais das colunas ----
FOLHA_TECNIC = "Tecnic"
FOLHA_MAQUINAS = "Machines"
COL_ID_MAQUINA = "Name"          # coluna da folha Maquinas com a lista de máquinas
COL_ID = "ID"                  # coluna do ID na folha Tecnic
COL_MACH_ID = "Code"             # coluna do ID na folha Machines
COL_MACH_NOME = "Name"      # coluna do nome na folha Machines
COL_MAQUINA = "Machine"             # coluna da folha Tecnic onde se escolhe a máquina
COL_OWNER = "Owner"
COL_ANOMALIA = "Anomaly"

OWNERS = ["NOVOBANCO", "PROSONIC"]
ANOMALIAS = [
    "Erro no módulo de notas",
    "Erro no módulo de moedas",
    "Erro no módulo de cheques",
    "Erro no sistema de tintagem",
    "Outro",
]

class ComboDelegate(QStyledItemDelegate):
    """Mostra uma lista de opções ao editar a célula."""

    def __init__(self, opcoes_fn, editavel=False, parent=None):
        super().__init__(parent)
        self.opcoes_fn = opcoes_fn      # função que devolve a lista (fica sempre atualizada)
        self.editavel = editavel        # True = também permite escrever texto livre

    def createEditor(self, parent, option, index):
        cb = QComboBox(parent)
        cb.setEditable(self.editavel)
        cb.addItems(self.opcoes_fn())
        # grava assim que se escolhe uma opção
        cb.activated.connect(lambda _: (self.commitData.emit(cb),
                                        self.closeEditor.emit(cb)))
        if not self.editavel:
            QTimer.singleShot(0, cb.showPopup)   # abre a lista logo ao clicar
        return cb

    def setEditorData(self, editor, index):
        valor = index.data(Qt.EditRole) or ""
        i = editor.findText(valor)
        if i >= 0:
            editor.setCurrentIndex(i)
        elif editor.isEditable():
            editor.setCurrentText(valor)
        else:
            editor.setCurrentIndex(-1)

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText().strip(), Qt.EditRole)


class Tabela(QWidget):
    def __init__(self, folha, usar_id=False):
        super().__init__()
        self.folha = folha
        self.colunas_combo = set()
        self.usar_id = usar_id
        self.delegates = {}
        self.tabela = QTableWidget()
        self.tabela.cellClicked.connect(self.clique_celula)
        self.tabela.itemChanged.connect(lambda _: self.ajustar())


        btn_add = QPushButton("Add Line")
        btn_del = QPushButton("Erase Selected")
        btn_add.clicked.connect(lambda: self.tabela.insertRow(self.tabela.rowCount()))
        btn_del.clicked.connect(self.apagar)

        botoes = QHBoxLayout()
        botoes.addWidget(btn_add)
        botoes.addWidget(btn_del)
        botoes.addStretch()

        layout = QVBoxLayout(self)
        layout.addLayout(botoes)
        layout.addWidget(self.tabela)
        self.carregar()

    def clique_celula(self, linha, coluna):
        if coluna in self.colunas_combo:
            self.tabela.edit(self.tabela.model().index(linha, coluna))

    def novo_item(self, valor, coluna):
        item = QTableWidgetItem(valor)
        if self.usar_id and self.tabela.horizontalHeaderItem(coluna).text() == COL_ID:
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # ID só de leitura
        return item



    def carregar(self):
        try:
            df = pd.read_excel(FICHEIRO, sheet_name=self.folha, dtype=str).fillna("")
        except ValueError:
            df = pd.DataFrame(columns=["Coluna1", "Coluna2"])

        if self.usar_id and COL_ID not in df.columns:
            df.insert(0, COL_ID, "")  # cria a coluna, é preenchida pela Janela

        self.tabela.setRowCount(len(df))
        self.tabela.setColumnCount(len(df.columns))
        self.tabela.setHorizontalHeaderLabels(list(df.columns))
        for r, linha in enumerate(df.values):
            for c, valor in enumerate(linha):
                self.tabela.setItem(r, c, self.novo_item(valor, c))
                self.tabela.resizeColumnsToContents()

    def coluna(self, nome):
        """Devolve o índice da coluna com esse cabeçalho, ou -1."""
        for c in range(self.tabela.columnCount()):
            if self.tabela.horizontalHeaderItem(c).text().strip() == nome:
                return c
        return -1

    def valores_coluna(self, nome):
        c = self.coluna(nome)
        if c < 0:
            return []
        vistos = []
        for r in range(self.tabela.rowCount()):
            item = self.tabela.item(r, c)
            v = item.text().strip() if item else ""
            if v and v not in vistos:
                vistos.append(v)
        return vistos

    def definir_combo(self, nome_coluna, delegate):
        c = self.coluna(nome_coluna)
        if c >= 0:
            self.tabela.setItemDelegateForColumn(c, delegate)
            self.colunas_combo.add(c)
            self.delegates[c] = delegate
        else:
            print(f"Aviso: coluna '{nome_coluna}' não encontrada na folha '{self.folha}'")

    def apagar(self):
        linhas = sorted({i.row() for i in self.tabela.selectedIndexes()}, reverse=True)
        for r in linhas:
            self.tabela.removeRow(r)

    def ajustar(self):
        """Ajusta ao conteúdo, mas nunca abaixo da largura da maior opção do dropdown."""
        self.tabela.resizeColumnsToContents()
        fm = self.tabela.fontMetrics()
        for c, delegate in self.delegates.items():
            opcoes = delegate.opcoes_fn()
            maior = max((fm.horizontalAdvance(o) for o in opcoes), default=0)
            minimo = maior + 45  # espaço para a seta e margens
            if self.tabela.columnWidth(c) < minimo:
                self.tabela.setColumnWidth(c, minimo)

    def para_dataframe(self):
        colunas = [self.tabela.horizontalHeaderItem(c).text()
                   for c in range(self.tabela.columnCount())]
        dados = []
        for r in range(self.tabela.rowCount()):
            dados.append([
                (self.tabela.item(r, c).text().strip() if self.tabela.item(r, c) else "")
                for c in range(len(colunas))
            ])
        return pd.DataFrame(dados, columns=colunas)


class Janela(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Manager (Beta)")
        self.resize(1000, 600)


        self.tabs = QTabWidget()
        self.paginas = {f: Tabela(f) for f in FOLHAS}
        for nome, p in self.paginas.items():
            self.tabs.addTab(p, nome)

        # Listas pendentes no Tecnic
        tecnic = self.paginas[FOLHA_TECNIC]
        tecnic.definir_combo(COL_MAQUINA, ComboDelegate(self.lista_maquinas, editavel=False, parent=tecnic))
        tecnic.definir_combo(COL_OWNER, ComboDelegate(lambda: OWNERS, editavel=False, parent=tecnic))
        tecnic.definir_combo(COL_ANOMALIA, ComboDelegate(lambda: ANOMALIAS, editavel=True, parent=tecnic))
        tecnic = self.paginas[FOLHA_TECNIC]
        col_maq = tecnic.coluna(COL_MAQUINA)

        # ao mudar a máquina numa linha, atualiza o ID dessa linha
        tecnic.tabela.itemChanged.connect(
            lambda it: self.atualizar_ids(it.row()) if it.column() == col_maq else None)
        # se editares a folha Machines, refresca todos os IDs
        self.paginas[FOLHA_MAQUINAS].tabela.itemChanged.connect(lambda _: self.atualizar_ids())

        self.atualizar_ids()  # preenche as linhas já existentes

        btn_guardar = QPushButton("Guardar")
        btn_guardar.clicked.connect(self.guardar)

        rodape = QHBoxLayout()
        rodape.addStretch()
        rodape.addWidget(btn_guardar)


        central = QWidget()
        layout = QVBoxLayout(central)
        layout.addWidget(self.tabs)
        layout.addLayout(rodape)
        self.setCentralWidget(central)

    def mapa_maquinas(self):
        """Devolve {nome da máquina: ID} a partir da folha Machines (em memória)."""
        m = self.paginas[FOLHA_MAQUINAS]
        cn, ci = m.coluna(COL_MACH_NOME), m.coluna(COL_MACH_ID)
        mapa = {}
        if cn < 0 or ci < 0:
            print("Aviso: colunas de nome/ID não encontradas na folha das máquinas")
            return mapa
        for r in range(m.tabela.rowCount()):
            nome, idv = m.tabela.item(r, cn), m.tabela.item(r, ci)
            if nome and nome.text().strip():
                mapa[nome.text().strip()] = idv.text().strip() if idv else ""
        return mapa

    def atualizar_ids(self, linha=None):
        """Preenche o ID da máquina em todas as linhas (ou só numa)."""
        t = self.paginas[FOLHA_TECNIC]
        cm, ci = t.coluna(COL_MAQUINA), t.coluna(COL_ID)
        if cm < 0 or ci < 0:
            return
        mapa = self.mapa_maquinas()
        linhas = range(t.tabela.rowCount()) if linha is None else [linha]
        t.tabela.blockSignals(True)
        for r in linhas:
            item = t.tabela.item(r, cm)
            nome = item.text().strip() if item else ""
            t.tabela.setItem(r, ci, t.novo_item(mapa.get(nome, ""), ci))
        t.tabela.blockSignals(False)
        t.ajustar()

    def lista_maquinas(self):
        # lê da tabela em memória, por isso já inclui máquinas novas ainda não guardadas
        return self.paginas[FOLHA_MAQUINAS].valores_coluna(COL_ID_MAQUINA)

    def guardar(self):
        try:
            shutil.copy(FICHEIRO, FICHEIRO.replace(".xlsx", ".bak.xlsx"))
            with pd.ExcelWriter(FICHEIRO, engine="openpyxl", mode="a",
                                if_sheet_exists="replace") as w:
                for nome, p in self.paginas.items():
                    p.para_dataframe().to_excel(w, sheet_name=nome, index=False)
            return True
        except PermissionError:
            QMessageBox.critical(self, "Erro", "O Excel está aberto. Fecha o ficheiro e tenta outra vez.")
            return False

    def guardar_e_limpar(self):
        if not self.guardar():
            return
        r = subprocess.run([sys.executable, DAILY_CLEAN], capture_output=True, text=True)
        msg = r.stdout or "Concluído."
        if r.returncode != 0:
            msg = r.stderr
        QMessageBox.information(self, "Daily Clean", msg[-2000:])


app = QApplication(sys.argv)
j = Janela()
j.show()
sys.exit(app.exec())