import math
import os
import sys
import pickle
from operator import concat

import numpy as np
import pandas as pd
import logging

from functools import partial
from datetime import datetime
from PyQt5 import QtCore, QtGui, QtWidgets, Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QTableWidgetItem, QLabel, QWidget, \
    QScrollArea, QListWidget
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtGui import QPainter, QColor, QPixmap, QRegion, QGuiApplication, QBrush

from Tolerances_Metric_Radial_Tapered_Roller_Bearings_UI import Ui_MainWindow  # Import the converted .ui file
from file_context import Context
from dotenv import load_dotenv
from PyQt5.QtCore import Qt, QTimer



def configure_logging():
    log_file_path = os.path.join(os.getenv("RKB_SOFTWARE_HOME2").replace("curr_user", os.getlogin()),
                                 "rkb_software.log")

    log_level_env = os.getenv("LOG_LEVEL", 'INFO')
    if log_level_env == 'DEBUG':
        log_level = logging.DEBUG
    elif log_level_env == 'INFO':
        log_level = logging.INFO
    elif log_level_env == 'WARNING':
        log_level = logging.WARNING
    elif log_level_env == 'ERROR':
        log_level = logging.ERROR
    else:
        log_level = logging.NOTSET

    logging.basicConfig(
        filename=log_file_path,  # Log file path
        level=log_level,  # Log level (INFO, DEBUG, ERROR, etc.)
        format="%(asctime)s - %(levelname)s - chamfer_dimensions: %(message)s",  # Log format
        force=True
    )


def is_dev():
    return os.getenv("ENV") == "dev"



def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

    # print(f'loading file from base_path={base_path} and relative_path={relative_path}')

    return os.path.join(base_path, relative_path)


# load resources
env_file = ".env.dev" if is_dev() else ".env.prod"
env_file_path = resource_path(env_file)
load_dotenv(env_file_path)

LOGO_FILE_PATH = os.getenv("LOGO_FILE_PATH")
LOGO_FILE_PATH2 = os.getenv("LOGO_FILE_PATH2")
WINDOW_ICO_FILE_PATH = os.getenv("WINDOW_ICO_FILE_PATH")

DRW1_FILE_PATH = os.getenv("DRW1_FILE_PATH")
DRW2_FILE_PATH = os.getenv("DRW2_FILE_PATH")

RKB_SOFTWARE_DATA_FOLDER = os.getenv("RKB_SOFTWARE_DATA").replace("curr_user", os.getlogin())

ISO15_TABLES_PATH = os.getenv("ISO15_TABLES_PATH")
ISO492_2023_PATH = os.getenv("ISO492_2023_PATH")


class RichTextDelegate(QtWidgets.QStyledItemDelegate):
    # (optional) keep read-only by disabling editors
    # def createEditor(self, *a, **k): return None

    def paint(self, painter, option, index):
        opt = QtWidgets.QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)

        text = index.data(QtCore.Qt.DisplayRole) or ""
        is_html = isinstance(text, str) and (
                "<sub>" in text or "<sup>" in text or text.lstrip().startswith("<")
        )

        if not is_html:
            super().paint(painter, opt, index)
            return

        style = opt.widget.style() if opt.widget else QtWidgets.QApplication.style()

        # Panel (selection bg, grid, etc.)
        style.drawPrimitive(QtWidgets.QStyle.PE_PanelItemViewItem, opt, painter, opt.widget)

        # Per-cell background when not selected
        if not (opt.state & QtWidgets.QStyle.State_Selected):
            bg = index.data(QtCore.Qt.BackgroundRole)
            if isinstance(bg, (QtGui.QBrush, QtGui.QColor)):
                painter.save()
                painter.fillRect(opt.rect, QtGui.QBrush(bg))
                painter.restore()

        # Text rect computed like Qt would
        text_rect = style.subElementRect(QtWidgets.QStyle.SE_ItemViewItemText, opt, opt.widget)

        # Build doc with the resolved (bold) font
        doc = QtGui.QTextDocument()
        doc.setDefaultFont(opt.font)

        # keep sub/sup same size and just shift a bit
        doc.setDefaultStyleSheet("""
            /* keep your font; only control sub/sup behavior */
            sub { font-size: 100%; vertical-align: sub; }
            sup { font-size: 100%; vertical-align: super; }
        """)

        # Horizontal alignment from item flags
        if opt.displayAlignment & QtCore.Qt.AlignHCenter:
            align_h = "center"
        elif opt.displayAlignment & QtCore.Qt.AlignRight:
            align_h = "right"
        else:
            align_h = "left"

        doc.setHtml(f"<div style='text-align:{align_h};'>{text}</div>")
        doc.setTextWidth(text_rect.width())

        # Compute vertical offset to honor VCenter / Bottom
        doc_h = doc.size().height()
        y = text_rect.top()
        if opt.displayAlignment & QtCore.Qt.AlignVCenter:
            y = text_rect.top() + (text_rect.height() - doc_h) / 2.0
        elif opt.displayAlignment & QtCore.Qt.AlignBottom:
            y = text_rect.bottom() - doc_h + 1  # +1 to avoid clipping

        # Selection text color
        ctx = QtGui.QAbstractTextDocumentLayout.PaintContext()
        if opt.state & QtWidgets.QStyle.State_Selected:
            ctx.palette.setColor(QtGui.QPalette.Text, opt.palette.highlightedText().color())

        painter.save()
        painter.translate(text_rect.left(), y)
        doc.documentLayout().draw(painter, ctx)
        painter.restore()

        # (optional) focus rect
        if opt.state & QtWidgets.QStyle.State_HasFocus:
            focus_opt = QtWidgets.QStyleOptionFocusRect()
            focus_opt.QStyleOption = opt
            focus_opt.rect = style.subElementRect(QtWidgets.QStyle.SE_ItemViewItemFocusRect, opt, opt.widget)
            focus_opt.backgroundColor = opt.palette.highlight().color()
            style.drawPrimitive(QtWidgets.QStyle.PE_FrameFocusRect, focus_opt, painter, opt.widget)


class Window(QMainWindow):
    c = Context("[Untitled]", "", "")
    programName = "Tolerances_Metric_Radial_Tapered_Roller_Bearings"
    programVersion = "Version 1.1.0-beta"
    lastUpdated = "15.10.2025"
    plot_in_progress = 0
    tables = []

    def __init__(self):
        super().__init__()

        # Load the UI from the converted .py file
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.MainWindow.setWindowIcon(QtGui.QIcon(resource_path(WINDOW_ICO_FILE_PATH)))
        self.ui.RKBLogo.setPixmap(QtGui.QPixmap(resource_path(LOGO_FILE_PATH)))

        self.ui.drawing1.setPixmap(QtGui.QPixmap(resource_path(DRW1_FILE_PATH)))
        self.ui.drawing2.setPixmap(QtGui.QPixmap(resource_path(DRW2_FILE_PATH)))

        # load lists on GUI
        self.loadLists()

        # change color for error message to red
        self.ui.errorLabel.setStyleSheet('color:red')
        self.ui.errorLabel.setText("")

        # update Window title including file name
        self.updateWindowTitle()

        # set fixed size for Window
        self.ui.MainWindow.setWindowFlags(QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowMinimizeButtonHint)

        # Monitor DPI and screen changes
        # screen = QGuiApplication.primaryScreen()
        screen = QGuiApplication.screenAt(self.ui.MainWindow.frameGeometry().center())
        screen.logicalDotsPerInchChanged.connect(self.ui.handle_screen_change)
        QApplication.instance().screenAdded.connect(self.ui.handle_screen_change)
        QApplication.instance().screenRemoved.connect(self.ui.handle_screen_change)


        self.T14_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Taper Normal - inner', 3)
        self.T15_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Taper Normal - outer', 3)
        self.T16_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Taper Normal - width', 3)
        self.T17_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Taper P6X - width', 3)
        self.T18_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Taper P5 - inner', 3)
        self.T19_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Taper P5 - outer', 3)
        self.T20_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Taper P5 - width', 3)
        self.T21_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Taper P4 - inner', 3)
        self.T22_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Taper P4 - outer', 3)
        self.T23_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Taper P4 - width', 3)
        self.T24_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Taper P2 - inner', 3)
        self.T25_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Taper P2 - outer', 3)
        self.T26_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Taper P2 - width', 3)
        self.T27_ISO492_2023 = self.readExcelFileISO492(self.resource_path(ISO492_2023_PATH), 'Flange outer', 3)
        self.T233_NSK = [
            [80, 120, 300, -300, 400, -400],
            [120, 180, 400, -400, 500, -500],
            [180, 250, 450, -450, 600, -600],
            [250, 315, 550, -550, 700, -700],
            [315, 400, 650, -650, 800, -800],
            [400, 500, 700, -700, 900, -900],
            [500, 630, 900, -900, 1000, -1000],
            [630, 800, 1200, -1200, 1500, -1500],
            [800, 1000, 1500, -1500, 1500, -1500],
            [1000, 1250, None, None, 1500, -1500],
            [1200, 1500, None, None, 1500, -1500],
            [1250, 1600, None, None, 1500, -1500],
            [1600, 2000, None, None, 1500, -1500],
        ]

        # Track changes in input data

        # input tab
        self.ui.partNoTextEdit.textChanged.connect(self.inputChanged)
        self.ui.bearingTypeListWidget.currentRowChanged.connect(self.inputChanged)
        self.ui.flangePresenceListWidget.currentRowChanged.connect(self.flangePresenceChanged)
        self.ui.boreDiameterLineEdit.textChanged.connect(self.stiffnessChanged)
        self.ui.outerDiameterLineEdit.textChanged.connect(self.stiffnessChanged)
        self.ui.orFlangeDiameterLineEdit.textChanged.connect(self.inputChanged)
        self.ui.precisionListWidget.currentRowChanged.connect(self.inputChanged)

        # Connect the buttons to actions
        self.ui.calculateButton.clicked.connect(self.calculateClicked)
        self.ui.resetButton.clicked.connect(self.resetClicked)
        self.ui.printButton.clicked.connect(self.print_tabs_to_pdf)

        # # Connect menu buttons to actions
        _translate = QtCore.QCoreApplication.translate
        # File menu
        self.ui.actionNew.triggered.connect(self.resetClicked)
        self.ui.actionNew.setShortcut(_translate("MainWindow", "Ctrl+N"))
        open_triggered_with_params = partial(self.openTriggered, None)
        self.ui.actionOpen.triggered.connect(open_triggered_with_params)
        self.ui.actionOpen.setShortcut(_translate("MainWindow", "Ctrl+O"))
        self.ui.actionSave.triggered.connect(self.saveTriggered)
        self.ui.actionSave.setShortcut(_translate("MainWindow", "Ctrl+S"))
        self.ui.actionSave_as.triggered.connect(self.saveAsTriggered)
        self.ui.actionPrint.triggered.connect(self.print_tabs_to_pdf)
        self.ui.actionPrint.setShortcut(_translate("MainWindow", "Ctrl+P"))
        self.ui.actionClose.triggered.connect(self.close)

        # Tools menu
        self.ui.actionCalculate.triggered.connect(self.calculateClicked)
        self.ui.actionReset.triggered.connect(self.resetClicked)

        # Help menu
        self.ui.actionAbout.triggered.connect(self.aboutTriggered)
        self.ui.actionHelp.triggered.connect(self.helpTriggered)

        # customize tables
        self.customizeInnerRingTableWidget1()
        self.customizeInnerRingTableWidget2()
        self.customizeOuterRingTableWidget1()
        self.customizeOuterRingTableWidget2()
        self.customizeFlangeTableWidget1()
        self.customizeFlangeTableWidget2()

        self.adjust_tables()

        # open file if filename transmitted as argument
        if len(sys.argv) > 1:
            initial_file = sys.argv[1]
            self.openTriggered(initial_file)

    def readExcelFiles(self, file, sheet1, sheet2, sheet3, sheet4, sheet5, sheet6, sheet7, sheet8, nCol):
        table1 = self.readExcelFileISO15(file, sheet1, nCol)
        table2 = self.readExcelFileISO15(file, sheet2, nCol)
        table3 = self.readExcelFileISO15(file, sheet3, nCol)
        table4 = self.readExcelFileISO15(file, sheet4, nCol)
        table5 = self.readExcelFileISO15(file, sheet5, nCol)
        table6 = self.readExcelFileISO15(file, sheet6, nCol)
        table7 = self.readExcelFileISO15(file, sheet7, nCol)
        table8 = self.readExcelFileISO15(file, sheet8, nCol)
        return table1, table2, table3, table4, table5, table6, table7, table8

    def customizeInnerRingTableWidget1(self):
        t = self.ui.innerRingTableWidget1
        __sortingEnabled = t.isSortingEnabled()
        t.setSortingEnabled(False)
        t.horizontalHeader().setVisible(False)
        t.verticalHeader().setVisible(False)
        t.setRowCount(12)

        # spans in first column (pairs of rows)
        t.setSpan(0, 0, 2, 1)
        t.setSpan(2, 0, 2, 1)
        t.setSpan(4, 0, 2, 1)
        t.setSpan(6, 0, 2, 1)
        t.setSpan(8, 0, 2, 1)
        t.setSpan(10, 0, 2, 1)

        font = QtGui.QFont("MS Shell Dlg 2", 11, QtGui.QFont.Bold)
        bg = QtGui.QColor(240, 240, 240)

        def add_item(r, c, text, use_bg=False):
            it = QtWidgets.QTableWidgetItem(text)
            it.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            it.setFont(font)
            it.setFlags(it.flags() & ~QtCore.Qt.ItemIsEditable)
            if use_bg:
                it.setBackground(bg)
            t.setItem(r, c, it)

        # --- Block 1: tΔdmp ---
        add_item(0, 0, "t<sub>Δdmp</sub>")
        add_item(0, 1, "U")
        add_item(1, 1, "L")
        add_item(0, 2, "")
        add_item(1, 2, "")

        # --- Block 2: tΔBs ---
        add_item(2, 0, "t<sub>ΔBs</sub>")
        add_item(2, 1, "U")
        add_item(3, 1, "L")
        add_item(2, 2, "")
        add_item(3, 2, "")

        # --- Block 3: tΔBgp ---
        add_item(4, 0, "t<sub>ΔBgp</sub>")
        add_item(4, 1, "U")
        add_item(5, 1, "L")
        add_item(4, 2, "")
        add_item(5, 2, "")

        # --- Block 4: tΔTg ---
        add_item(6, 0, "t<sub>ΔTg</sub>")
        add_item(6, 1, "U")
        add_item(7, 1, "L")
        add_item(6, 2, "")
        add_item(7, 2, "")

        # --- Block 5 (grey): tΔT1g ---
        add_item(8, 0, "t<sub>ΔT1g</sub>", use_bg=True)
        add_item(8, 1, "U", use_bg=True)
        add_item(9, 1, "L", use_bg=True)
        add_item(8, 2, "", use_bg=True)
        add_item(9, 2, "", use_bg=True)

        # --- Block 6 (grey): tΔds ---
        add_item(10, 0, "t<sub>Δds</sub>", use_bg=True)
        add_item(10, 1, "U", use_bg=True)
        add_item(11, 1, "L", use_bg=True)
        add_item(10, 2, "", use_bg=True)
        add_item(11, 2, "", use_bg=True)

        # Column widths
        cols = t.columnCount()
        for col in range(cols):
            t.setColumnWidth(col, 60)

        # Ensure every cell exists + centered (defensive)
        rows = t.rowCount()
        for r in range(rows):
            for c in range(cols):
                if not t.item(r, c):
                    t.setItem(r, c, QtWidgets.QTableWidgetItem(""))
                t.item(r, c).setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

        # Rich text for <sub>/<sup>
        t.setItemDelegate(RichTextDelegate(t))
        # Read-only table
        t.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        t.setSortingEnabled(__sortingEnabled)



    def customizeInnerRingTableWidget2(self):
        t = self.ui.innerRingTableWidget2
        __sortingEnabled = t.isSortingEnabled()
        t.setSortingEnabled(False)
        t.horizontalHeader().setVisible(False)
        t.verticalHeader().setVisible(False)
        t.setRowCount(5)

        font = QtGui.QFont("MS Shell Dlg 2", 11, QtGui.QFont.Bold)
        bg = QtGui.QColor(240, 240, 240)

        def add_item(r, c, text, use_bg=False):
            it = QtWidgets.QTableWidgetItem(text)
            it.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            it.setFont(font)
            it.setFlags(it.flags() & ~QtCore.Qt.ItemIsEditable)
            if use_bg:
                it.setBackground(bg)
            t.setItem(r, c, it)

        # Rows
        add_item(0, 0, "t<sub>Vdsp</sub>")
        add_item(0, 1, "")

        add_item(1, 0, "t<sub>Vdmp</sub>")
        add_item(1, 1, "")

        add_item(2, 0, "t<sub>Kia</sub>")
        add_item(2, 1, "")

        # Grey section
        add_item(3, 0, "t<sub>Sd</sub>", use_bg=True)
        add_item(3, 1, "", use_bg=True)

        add_item(4, 0, "t<sub>Sia</sub>", use_bg=True)
        add_item(4, 1, "", use_bg=True)

        # Column widths
        cols = t.columnCount()
        for c in range(cols):
            t.setColumnWidth(c, 60)

        # Ensure every cell exists + centered (defensive)
        rows = t.rowCount()
        for r in range(rows):
            for c in range(cols):
                if not t.item(r, c):
                    t.setItem(r, c, QtWidgets.QTableWidgetItem(""))
                t.item(r, c).setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

        # Render HTML (<sub>/<sup>) consistently
        t.setItemDelegate(RichTextDelegate(t))
        # Make table read-only
        t.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        t.setSortingEnabled(__sortingEnabled)

    def customizeOuterRingTableWidget1(self):
        t = self.ui.outerRingTableWidget1
        __sortingEnabled = t.isSortingEnabled()
        t.setSortingEnabled(False)
        t.horizontalHeader().setVisible(False)
        t.verticalHeader().setVisible(False)
        t.setRowCount(10)

        # spans in first column
        t.setSpan(0, 0, 2, 1)
        t.setSpan(2, 0, 2, 1)
        t.setSpan(4, 0, 2, 1)
        t.setSpan(6, 0, 2, 1)
        t.setSpan(8, 0, 2, 1)

        font = QtGui.QFont("MS Shell Dlg 2", 11, QtGui.QFont.Bold)
        bg = QtGui.QColor(240, 240, 240)

        def add_item(r, c, text, use_bg=False):
            it = QtWidgets.QTableWidgetItem(text)
            it.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            it.setFont(font)
            it.setFlags(it.flags() & ~QtCore.Qt.ItemIsEditable)
            if use_bg:
                it.setBackground(bg)
            t.setItem(r, c, it)

        # --- Block 1: tΔDmp ---
        add_item(0, 0, "t<sub>ΔDmp</sub>")
        add_item(0, 1, "U")
        add_item(1, 1, "L")
        add_item(0, 2, "")
        add_item(1, 2, "")

        # --- Block 2: tΔCs ---
        add_item(2, 0, "t<sub>ΔCs</sub>")
        add_item(2, 1, "U")
        add_item(3, 1, "L")
        add_item(2, 2, "")
        add_item(3, 2, "")

        # --- Block 3: tΔCgp ---
        add_item(4, 0, "t<sub>ΔCgp</sub>")
        add_item(4, 1, "U")
        add_item(5, 1, "L")
        add_item(4, 2, "")
        add_item(5, 2, "")

        # --- Block 4 (grey): tΔT2g ---
        add_item(6, 0, "t<sub>ΔT2g</sub>", use_bg=True)
        add_item(6, 1, "U", use_bg=True)
        add_item(7, 1, "L", use_bg=True)
        add_item(6, 2, "", use_bg=True)
        add_item(7, 2, "", use_bg=True)

        # --- Block 5 (grey): tΔDs ---
        add_item(8, 0, "t<sub>ΔDs</sub>", use_bg=True)
        add_item(8, 1, "U", use_bg=True)
        add_item(9, 1, "L", use_bg=True)
        add_item(8, 2, "", use_bg=True)
        add_item(9, 2, "", use_bg=True)

        # Column widths
        cols = t.columnCount()
        for c in range(cols):
            t.setColumnWidth(c, 60)

        # Ensure every cell exists + centered (defensive)
        rows = t.rowCount()
        for r in range(rows):
            for c in range(cols):
                if not t.item(r, c):
                    t.setItem(r, c, QtWidgets.QTableWidgetItem(""))
                t.item(r, c).setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

        # Render HTML (<sub>/<sup>) consistently (and vertically center via your adapted delegate)
        t.setItemDelegate(RichTextDelegate(t))
        # Make table read-only
        t.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        t.setSortingEnabled(__sortingEnabled)

    def customizeOuterRingTableWidget2(self):
        t = self.ui.outerRingTableWidget2
        __sortingEnabled = t.isSortingEnabled()
        t.setSortingEnabled(False)
        t.horizontalHeader().setVisible(False)
        t.verticalHeader().setVisible(False)
        t.setRowCount(5)

        font = QtGui.QFont("MS Shell Dlg 2", 11, QtGui.QFont.Bold)
        bg = QtGui.QColor(240, 240, 240)

        def add_item(r, c, text, use_bg=False):
            it = QtWidgets.QTableWidgetItem(text)
            it.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            it.setFont(font)
            it.setFlags(it.flags() & ~QtCore.Qt.ItemIsEditable)
            if use_bg:
                it.setBackground(bg)
            t.setItem(r, c, it)

        # Rows
        add_item(0, 0, "t<sub>VDsp</sub>")
        add_item(0, 1, "")

        add_item(1, 0, "t<sub>VDmp</sub>")
        add_item(1, 1, "")

        add_item(2, 0, "t<sub>Kea</sub>")
        add_item(2, 1, "")

        # Grey section
        add_item(3, 0, "t<sub>SD</sub>", use_bg=True)
        add_item(3, 1, "", use_bg=True)

        add_item(4, 0, "t<sub>Sea</sub>", use_bg=True)
        add_item(4, 1, "", use_bg=True)

        # Column widths
        cols = t.columnCount()
        for c in range(cols):
            t.setColumnWidth(c, 60)

        # Ensure every cell exists + centered (defensive)
        rows = t.rowCount()
        for r in range(rows):
            for c in range(cols):
                if not t.item(r, c):
                    t.setItem(r, c, QtWidgets.QTableWidgetItem(""))
                t.item(r, c).setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

        # Render HTML (<sub>/<sup>) consistently (and vertical centering via your delegate)
        t.setItemDelegate(RichTextDelegate(t))
        # Make table read-only
        t.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        t.setSortingEnabled(__sortingEnabled)

    def customizeFlangeTableWidget1(self):
        t = self.ui.flangeTableWidget1
        __sortingEnabled = t.isSortingEnabled()
        t.setSortingEnabled(False)
        t.horizontalHeader().setVisible(False)
        t.verticalHeader().setVisible(False)
        t.setRowCount(6)

        # spans in first column
        t.setSpan(0, 0, 2, 1)
        t.setSpan(2, 0, 2, 1)
        t.setSpan(4, 0, 2, 1)

        font = QtGui.QFont("MS Shell Dlg 2", 11, QtGui.QFont.Bold)
        bg = QtGui.QColor(240, 240, 240)

        def add_item(r, c, text, use_bg=False):
            it = QtWidgets.QTableWidgetItem(text)
            it.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            it.setFont(font)
            it.setFlags(it.flags() & ~QtCore.Qt.ItemIsEditable)
            if use_bg:
                it.setBackground(bg)
            t.setItem(r, c, it)

        # --- Block 1: tΔD1s ---
        add_item(0, 0, "t<sub>ΔD1s</sub>")
        add_item(0, 1, "U")
        add_item(1, 1, "L")
        add_item(0, 2, "")
        add_item(1, 2, "")

        # --- Block 2: tΔTFg ---
        add_item(2, 0, "t<sub>ΔTFg</sub>")
        add_item(2, 1, "U")
        add_item(3, 1, "L")
        add_item(2, 2, "")
        add_item(3, 2, "")

        # --- Block 3 (grey): tΔTF2g ---
        add_item(4, 0, "t<sub>ΔTF2g</sub>", use_bg=True)
        add_item(4, 1, "U", use_bg=True)
        add_item(5, 1, "L", use_bg=True)
        add_item(4, 2, "", use_bg=True)
        add_item(5, 2, "", use_bg=True)

        # Column widths
        cols = t.columnCount()
        for c in range(cols):
            t.setColumnWidth(c, 60)

        # Ensure every cell exists + centered (defensive)
        rows = t.rowCount()
        for r in range(rows):
            for c in range(cols):
                if not t.item(r, c):
                    t.setItem(r, c, QtWidgets.QTableWidgetItem(""))
                t.item(r, c).setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

        # Rich text for <sub>/<sup> and vertical centering via your delegate
        t.setItemDelegate(RichTextDelegate(t))
        # Read-only table
        t.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        t.setSortingEnabled(__sortingEnabled)

    def customizeFlangeTableWidget2(self):
        t = self.ui.flangeTableWidget2
        __sortingEnabled = t.isSortingEnabled()
        t.setSortingEnabled(False)
        t.horizontalHeader().setVisible(False)
        t.verticalHeader().setVisible(False)
        t.setRowCount(2)

        font = QtGui.QFont("MS Shell Dlg 2", 11, QtGui.QFont.Bold)
        bg = QtGui.QColor(240, 240, 240)

        def add_item(r, c, text, use_bg=False):
            it = QtWidgets.QTableWidgetItem(text)
            it.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            it.setFont(font)
            it.setFlags(it.flags() & ~QtCore.Qt.ItemIsEditable)
            if use_bg:
                it.setBackground(bg)
            t.setItem(r, c, it)

        # Grey rows
        add_item(0, 0, "t<sub>SD1</sub>", use_bg=True)
        add_item(0, 1, "", use_bg=True)

        add_item(1, 0, "t<sub>Sea1</sub>", use_bg=True)
        add_item(1, 1, "", use_bg=True)

        # Column widths
        cols = t.columnCount()
        for c in range(cols):
            t.setColumnWidth(c, 60)

        # Ensure every cell exists + centered (defensive)
        rows = t.rowCount()
        for r in range(rows):
            for c in range(cols):
                if not t.item(r, c):
                    t.setItem(r, c, QtWidgets.QTableWidgetItem(""))
                t.item(r, c).setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

        # Rich text for <sub>/<sup> and vertical centering via your delegate
        t.setItemDelegate(RichTextDelegate(t))
        # Read-only table
        t.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        t.setSortingEnabled(__sortingEnabled)

    def aboutTriggered(self):
        '''
        displays message box with details about the program
        :return:
        '''
        msg = QMessageBox()
        # for local run
        msg.setIconPixmap(QPixmap(resource_path(LOGO_FILE_PATH2)))
        msg.setWindowIcon(QtGui.QIcon(resource_path(WINDOW_ICO_FILE_PATH)))

        # for exe with --onefile option
        # msg.setIconPixmap(QPixmap(self.resource_path(".\\RKBlogo3.png")))
        # msg.setWindowIcon(QtGui.QIcon(self.resource_path('.\\icon.ico')))

        msg.setStandardButtons(QMessageBox.Cancel)
        msg.setInformativeText("<p align='center'>This program comes with absolutely no warranty.</p>")
        msg.setWindowTitle("About " + self.programName)
        msg.setText("<p align='center'>{}<br>"
                    "\u00A9 RKB Bearings Industries Switzerland <br><br>"
                    "Last updated at {}</p>".format(self.programVersion, self.lastUpdated))
        x = msg.exec_()

    def helpTriggered(self):
        msg = QMessageBox()
        msg.setWindowTitle("Help")
        msg.setText("<p align='center'> No documentation available for now. </p>")

        # for local run
        msg.setWindowIcon(QtGui.QIcon(resource_path(WINDOW_ICO_FILE_PATH)))
        msg.setIconPixmap(QPixmap(resource_path(WINDOW_ICO_FILE_PATH)))

        # for exe with --onefile option
        # msg.setWindowIcon(QtGui.QIcon(self.resource_path('.\\icon.ico')))
        # msg.setIconPixmap(QPixmap(self.resource_path('.\\icon.ico')))

        x = msg.exec_()

    def openTriggered(self, initial_file=None):
        '''
        Loads all the data from the file to the form
        :return:
        '''
        file_to_open = None
        if initial_file is not None:
            file_to_open = self.resource_path(initial_file)

        else:
            path = self.resource_path(
                os.path.join('C:/', 'Documents and Settings', os.getlogin(), 'Documents/RKB App Data'))
            fname = QFileDialog.getOpenFileName(self, 'Open file', path, 'Data files (*.rkb)')
            file_to_open = fname[0]

        if file_to_open != "":
            with open(file_to_open, 'rb') as f:
                try:
                    data = pickle.load(f)
                    self.resetClicked()
                    if data["programName"] == self.programName:
                        # set input data
                        # input Tab
                        self.ui.partNoTextEdit.setText(data["partNo"])
                        self.ui.innerRingSymmetryListWidget.setCurrentRow(data["irSymmetry"])
                        self.ui.outerRingSymmetryListWidget.setCurrentRow(data["orSymmetry"])
                        self.ui.bearingTypeListWidget.setCurrentRow(data["bearingType"])
                        self.ui.flangePresenceListWidget.setCurrentRow(data["flangePresence"])
                        self.ui.flangeTypeListWidget.setCurrentRow(data["flangeType"])
                        self.ui.boreDiameterLineEdit.setText(data["d"])
                        self.ui.outerDiameterLineEdit.setText(data["D"])
                        self.ui.orFlangeDiameterLineEdit.setText(data["D1"])
                        self.ui.precisionListWidget.setCurrentRow(data["precision"])
                        self.ui.errorLabel.setText(data["error"])
                        self.ui.stiffnessLineEdit.setText(data["stiffness"])
                        try:
                            self.ui.fsLineEdit.setText(data["fs"])
                        except Exception as ex:
                            self.showWarning("fs missing! Click calculate and then save the file!")

                        # output Tab

                        font = QtGui.QFont()
                        font.setFamily("MS Shell Dlg 2")
                        font.setPointSize(12)

                        data_dict = data["ir1"]
                        # Populate table cells
                        self.customizeInnerRingTableWidget1()
                        for i in range(len(data_dict)):
                            item = QTableWidgetItem(str(data_dict[i]))
                            item.setTextAlignment(QtCore.Qt.AlignCenter)
                            item.setFont(font)
                            if i > 7:
                                item.setBackground(QtGui.QColor(240, 240, 240))
                            self.ui.innerRingTableWidget1.setItem(i, 2, item)

                        data_dict = data["ir2"]
                        # Populate table cells
                        self.customizeInnerRingTableWidget2()
                        for i in range(len(data_dict)):
                            item = QTableWidgetItem(str(data_dict[i]))
                            item.setTextAlignment(QtCore.Qt.AlignCenter)
                            item.setFont(font)
                            if i > 2:
                                item.setBackground(QtGui.QColor(240, 240, 240))
                            self.ui.innerRingTableWidget2.setItem(i, 1, item)

                        data_dict = data["or1"]
                        # Populate table cells
                        self.customizeOuterRingTableWidget1()
                        for i in range(len(data_dict)):
                            item = QTableWidgetItem(str(data_dict[i]))
                            item.setTextAlignment(QtCore.Qt.AlignCenter)
                            item.setFont(font)
                            if i > 5:
                                item.setBackground(QtGui.QColor(240, 240, 240))
                            self.ui.outerRingTableWidget1.setItem(i, 2, item)

                        data_dict = data["or2"]
                        # Populate table cells
                        self.customizeOuterRingTableWidget2()
                        for i in range(len(data_dict)):
                            item = QTableWidgetItem(str(data_dict[i]))
                            item.setTextAlignment(QtCore.Qt.AlignCenter)
                            item.setFont(font)
                            if i > 2:
                                item.setBackground(QtGui.QColor(240, 240, 240))
                            self.ui.outerRingTableWidget2.setItem(i, 1, item)

                        data_dict = data["f1"]
                        # Populate table cells
                        self.customizeFlangeTableWidget1()
                        for i in range(len(data_dict)):
                            item = QTableWidgetItem(str(data_dict[i]))
                            item.setTextAlignment(QtCore.Qt.AlignCenter)
                            item.setFont(font)
                            if i > 3:
                                item.setBackground(QtGui.QColor(240, 240, 240))
                            self.ui.flangeTableWidget1.setItem(i, 2, item)

                        data_dict = data["f2"]
                        # Populate table cells
                        self.customizeFlangeTableWidget2()
                        for i in range(len(data_dict)):
                            item = QTableWidgetItem(str(data_dict[i]))
                            item.setTextAlignment(QtCore.Qt.AlignCenter)
                            item.setFont(font)
                            item.setBackground(QtGui.QColor(240, 240, 240))
                            self.ui.flangeTableWidget2.setItem(i, 1, item)

                        self.c.status = ""
                        self.c.fileName = file_to_open
                        self.c.filePath = file_to_open
                        self.updateWindowTitle()
                    else:
                        self.resetInput()
                        self.ui.errorLabel.setText("You opened a file for a different RKB program!")
                        self.resetOutput()

                        # msg = QMessageBox()
                        # msg.setWindowTitle("Error")
                        # msg.setText("<p align='center'> You opened a file for a different RKB program! </p>")
                        #
                        # msg.setIconPixmap(QPixmap(resource_path(WINDOW_ICO_FILE_PATH)))
                        # msg.setWindowIcon(QtGui.QIcon(resource_path(WINDOW_ICO_FILE_PATH)))
                        #
                        # x = msg.exec_()
                        self.c.reset()

                        self.updateWindowTitle()
                except Exception as ex:
                    # self.resetClicked()
                    # msg = QMessageBox()
                    # msg.setWindowTitle("Error")
                    # msg.setText("<p align='center'> The file cannot be opened! </p>")
                    #
                    #
                    # msg.setIconPixmap(QPixmap(resource_path(WINDOW_ICO_FILE_PATH)))
                    # msg.setWindowIcon(QtGui.QIcon(resource_path(WINDOW_ICO_FILE_PATH)))
                    #
                    # x = msg.exec_()
                    self.resetInput()
                    self.ui.errorLabel.setText("The file cannot be opened!")
                    self.resetOutput()

        self.adjust_tables()

    def adjust_tables(self):
        # screen = QGuiApplication.primaryScreen()
        # screen = self.screen()
        screen = QGuiApplication.screenAt(self.ui.MainWindow.frameGeometry().center())
        scale_factor = screen.logicalDotsPerInch() / 96.0

        self.ui.adjust_table_widget(self.ui.innerRingTableWidget1, scale_factor)
        self.ui.adjust_table_widget(self.ui.innerRingTableWidget2, scale_factor)
        self.ui.adjust_table_widget(self.ui.outerRingTableWidget1, scale_factor)
        self.ui.adjust_table_widget(self.ui.outerRingTableWidget2, scale_factor)
        self.ui.adjust_table_widget(self.ui.flangeTableWidget1, scale_factor)
        self.ui.adjust_table_widget(self.ui.flangeTableWidget2, scale_factor)

    def saveTriggered(self):
        '''
        Saves all the data into the open file
        :return:
        '''
        if self.c.fileName == "[Untitled]":
            self.saveAsTriggered()
        else:
            # try:
            textContent = self.getAllValues()
            saveMsg = "Are you sure you want to save the changes?"
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Question)
            # for local run
            msg.setWindowIcon(QtGui.QIcon(resource_path(WINDOW_ICO_FILE_PATH)))
            # for exe with --onefile option
            # msg.setWindowIcon(QtGui.QIcon(self.resource_path('.\\icon.ico')))

            msg.setWindowTitle("Save Changes")
            msg.setText(saveMsg)
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.Cancel)
            rsp = msg.exec_()

            if rsp == QMessageBox.Yes:
                with open(self.c.fileName, 'wb') as f:
                    pickle.dump(textContent, f)
                self.c.status = ""
                self.updateWindowTitle()
            # except Exception as ex:
            #     self.showError(str(ex))

    def saveCloseTriggered(self):
        '''
        Saves all the data into the open file
        :return:
        '''
        textContent = self.getAllValues()
        if self.c.fileName == "[Untitled]":
            self.saveAsTriggered()
        else:
            with open(self.c.fileName, 'wb') as f:
                pickle.dump(textContent, f)
            self.c.status = ""
            self.updateWindowTitle()

    def getAllValues(self):
        '''
        :return: returns as text the input and output values and error message
        '''
        textContent = {"programName": self.programName,
                       "programVersion": self.programVersion,
                       "lastUpdated": self.lastUpdated,
                       # input Tab
                       "partNo": self.ui.partNoTextEdit.text(),
                       "irSymmetry": self.ui.innerRingSymmetryListWidget.currentRow(),
                       "orSymmetry": self.ui.outerRingSymmetryListWidget.currentRow(),
                       "stiffness": self.ui.stiffnessLineEdit.text(),
                       "fs": self.ui.fsLineEdit.text(),
                       "bearingType": self.ui.bearingTypeListWidget.currentRow(),
                       "flangePresence": self.ui.flangePresenceListWidget.currentRow(),
                       "flangeType": self.ui.flangeTypeListWidget.currentRow(),
                       "d": self.ui.boreDiameterLineEdit.text(),
                       "D": self.ui.outerDiameterLineEdit.text(),
                       "D1": self.ui.orFlangeDiameterLineEdit.text(),
                       "precision": self.ui.precisionListWidget.currentRow(),
                       "error": self.ui.errorLabel.text(),
                       # output Tab
                       "ir1": {},
                       "ir2": {},
                       "or1": {},
                       "or2": {},
                       "f1": {},
                       "f2": {}
                       }

        # get data from innerRingTableWidget1
        rows = self.ui.innerRingTableWidget1.rowCount()
        ir1 = []
        for i in range(rows):
            if (self.ui.innerRingTableWidget1.item(i, 2) is not None) and (
                    self.ui.innerRingTableWidget1.item(i, 2).text() != ""):
                ir1.append((self.ui.innerRingTableWidget1.item(i, 2).text()))
        textContent["ir1"] = ir1

        # get data from innerRingTableWidget2
        rows = self.ui.innerRingTableWidget2.rowCount()
        ir2 = []
        for i in range(rows):
            if (self.ui.innerRingTableWidget2.item(i, 1) is not None) and (
                    self.ui.innerRingTableWidget2.item(i, 1).text() != ""):
                ir2.append((self.ui.innerRingTableWidget2.item(i, 1).text()))
        textContent["ir2"] = ir2

        # get data from outerRingTableWidget1
        rows = self.ui.outerRingTableWidget1.rowCount()
        or1 = []
        for i in range(rows):
            if (self.ui.outerRingTableWidget1.item(i, 2) is not None) and (
                    self.ui.outerRingTableWidget1.item(i, 2).text() != ""):
                or1.append(self.ui.outerRingTableWidget1.item(i, 2).text())
        textContent["or1"] = or1

        # get data from outerRingTableWidget2
        rows = self.ui.outerRingTableWidget2.rowCount()
        or2 = []
        for i in range(rows):
            if (self.ui.outerRingTableWidget2.item(i, 1) is not None) and (
                    self.ui.outerRingTableWidget2.item(i, 1).text() != ""):
                or2.append(self.ui.outerRingTableWidget2.item(i, 1).text())
        textContent["or2"] = or2

        # get data from flangeTableWidget1
        rows = self.ui.flangeTableWidget1.rowCount()
        f1 = []
        for i in range(rows):
            if (self.ui.flangeTableWidget1.item(i, 2) is not None) and (
                    self.ui.flangeTableWidget1.item(i, 2).text() != ""):
                f1.append(self.ui.flangeTableWidget1.item(i, 2).text())
        textContent["f1"] = f1

        # get data from flangeTableWidget2
        rows = self.ui.flangeTableWidget2.rowCount()
        f2 = []
        for i in range(rows):
            if (self.ui.flangeTableWidget2.item(i, 1) is not None) and (
                    self.ui.flangeTableWidget2.item(i, 1).text() != ""):
                f2.append(self.ui.flangeTableWidget2.item(i, 1).text())
        textContent["f2"] = f2

        return textContent

    def saveAsTriggered(self):
        '''
        Open a dialog box to allow saving all the data from the form in a file
        :return:
        '''
        date = datetime.today().strftime('%Y-%m-%d')
        path = self.resource_path(
            os.path.join('C:/', 'Documents and Settings', os.getlogin(), 'Documents/RKB App Data/'))
        pathname = QFileDialog.getSaveFileName(self, 'Save File',
                                               path + 'Tolerances_Metric_Radial_Tapered_Roller_Bearings-' + date,
                                               'Data files (*.rkb)')
        textContent = self.getAllValues()
        if (pathname[0]) != "":
            with open(pathname[0], 'wb') as f:
                pickle.dump(textContent, f)
            # self.currentPath = pathname[0]
            self.setWindowTitle(pathname[0])
            self.c.fileName = pathname[0]
            self.c.status = ""
            self.c.filePath = pathname[0]
            self.updateWindowTitle()

    def resource_path(self, relative_path):
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)

    def readExcelFileISO492(self, file, sheetName, nRows):

        # Read the Excel sheet without header
        df = pd.read_excel(file, sheet_name=sheetName, header=None)

        # Drop all-empty rows
        df.dropna(how='all', inplace=True)

        # Drop first nRows rows (header and spacing)
        df = df.iloc[nRows:]

        # Optional: reset index
        df.reset_index(drop=True, inplace=True)

        # Convert to list of lists
        table = df.values.tolist()
        return table

    def readExcelFileISO15(self, file, sheetName, nCol):

        df = pd.read_excel(file, sheet_name=sheetName)

        # Remove rows where all elements are NaN (empty rows)
        df = df.dropna(how='all')

        # Remove empty rows/cols
        df = df.drop(df.columns[0], axis=1)
        df = df.drop(df.columns[-nCol:], axis=1)
        df = df.iloc[2:]
        df.iat[0, 0] = "d"
        df.iat[0, 1] = "D"

        # Fix table header
        df.columns = df.iloc[0]
        df = df[1:]

        # Create dictionary
        df.reset_index(drop=True, inplace=True)
        df = df.drop_duplicates(subset=['d', 'D'])
        table = df.set_index(['d', 'D']).to_dict(orient='index')
        return table

    def getStiffnessSeries(self, D, d):
        fs = (D - d)/(d ** 0.9)
        self.ui.fsLineEdit.setText("{:.4f}".format(fs))
        if fs <= 0.535:
            return "A"
        elif fs <= 0.73:
            return "B"
        elif fs <= 1.3:
            return "C"
        else:
            return "S"

    def calculateClicked(self):
        '''
        calculates output and displays results or pop-ups for handling errors and warnings
        :return:
        '''

        # obtain input data and handle missing input errors

        errMsg = ""

        try:
            d = float(self.ui.boreDiameterLineEdit.text())
        except Exception as ex:
            errMsg = "The value for the bore diameter is missing!"

        try:
            D = float(self.ui.outerDiameterLineEdit.text())
        except Exception as ex:
            errMsg = "The value for the outer diameter is missing!"

        try:
            D1 = float(self.ui.orFlangeDiameterLineEdit.text())
        except Exception as ex:
            errMsg = "The value for the outer ring flange diameter is missing!"

        if errMsg != '':
            self.ui.errorLabel.setText(errMsg)

        if errMsg == '' and self.ui.errorLabel.text() == "":

            font = QtGui.QFont()
            font.setFamily("MS Shell Dlg 2")
            font.setPointSize(12)

            S = self.getStiffnessSeries(D, d)
            self.ui.stiffnessLineEdit.setText(S)

            t_delta_ir = self.get_t_delta_ir(d)
            for i in range(len(t_delta_ir)):
                value = t_delta_ir[i]
                try:
                    num = int(value)
                    if num > 0:
                        text = f"+{num}"
                    else:
                        text = str(num)
                except ValueError:
                    text = str(value)
                item = QTableWidgetItem(text)
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                item.setFont(font)
                if i > 7:
                    item.setBackground(QtGui.QColor(240, 240, 240))
                self.ui.innerRingTableWidget1.setItem(i, 2, item)

            t_V_ir = self.get_t_V_ir(d, S)
            for i in range(len(t_V_ir)):
                value = t_V_ir[i]
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                item.setFont(font)
                if i > 2:
                    item.setBackground(QtGui.QColor(240, 240, 240))
                self.ui.innerRingTableWidget2.setItem(i, 1, item)

            t_delta_or = self.get_t_delta_or(D, d)
            for i in range(len(t_delta_or)):
                value = t_delta_or[i]
                try:
                    num = int(value)
                    if num > 0:
                        text = f"+{num}"
                    else:
                        text = str(num)
                except ValueError:
                    text = str(value)
                item = QTableWidgetItem(text)
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                item.setFont(font)
                if i > 5:
                    item.setBackground(QtGui.QColor(240, 240, 240))
                self.ui.outerRingTableWidget1.setItem(i, 2, item)

            t_V_or = self.get_t_V_or(D, S)
            for i in range(len(t_V_or)):
                value = t_V_or[i]
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                item.setFont(font)
                if i > 2:
                    item.setBackground(QtGui.QColor(240, 240, 240))
                self.ui.outerRingTableWidget2.setItem(i, 1, item)

            t_delta_F = self.get_t_delta_F(D1, d)
            for i in range(len(t_delta_F)):
                value = t_delta_F[i]
                try:
                    num = int(value)
                    if num > 0:
                        text = f"+{num}"
                    else:
                        text = str(num)
                except ValueError:
                    text = str(value)
                item = QTableWidgetItem(text)
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                item.setFont(font)
                if i > 3:
                    item.setBackground(QtGui.QColor(240, 240, 240))
                self.ui.flangeTableWidget1.setItem(i, 2, item)

            t_V_F = self.get_t_V_F(D,D1)
            for i in range(len(t_V_F)):
                value = t_V_F[i]
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                item.setFont(font)
                item.setBackground(QtGui.QColor(240, 240, 240))
                self.ui.flangeTableWidget2.setItem(i, 1, item)

            # go to first tab with output
            self.ui.tabWidget.setCurrentIndex(1)

            self.c.status = '*'
            self.updateWindowTitle()
            # except Exception as ex:
            #     self.showError(str(ex))
        else:
            self.resetOutput()

    def searchCol_OC_LITE(self, x, A, colm, colM):
        n = len(A)
        if x < A[0][colm] or x > A[n - 1][colM]:
            return "N/A"
        if x == A[0][colm]:
            return 0
        for i in range(n):
            if A[i][colm] < x and x <= A[i][colM]:
                return i

    def get_t_delta_ir_P0(self, d):
        rez = []
        for i in range(12):
            rez.append("—")
        i_d = self.searchCol_OC_LITE(d, self.T14_ISO492_2023, 0, 1)
        if i_d == "N/A":
            rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
            return rez
        else:
            rez[0] = self.T14_ISO492_2023[i_d][2]
            rez[1] = self.T14_ISO492_2023[i_d][3]
        if self.ui.bearingTypeListWidget.currentItem().text() != "Single row":
            i_d_NSK = self.searchCol_OC_LITE(d, self.T233_NSK, 0, 1)
            if i_d_NSK == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDO":
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[4] = self.T233_NSK[i_d_NSK][2]
                        rez[5] = self.T233_NSK[i_d_NSK][3]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDI":
                    i_d_BCT = self.searchCol_OC_LITE(d, self.T16_ISO492_2023, 0, 1)
                    if i_d_BCT == "N/A":
                        rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                        return rez
                    else:
                        rez[2] = self.T16_ISO492_2023[i_d_BCT][2]
                        rez[3] = self.T16_ISO492_2023[i_d_BCT][3]
                        if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                            rez[6] = self.T233_NSK[i_d_NSK][2]
                            rez[7] = self.T233_NSK[i_d_NSK][3]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQO":
                    rez[2] = self.T233_NSK[i_d_NSK][4]
                    rez[3] = self.T233_NSK[i_d_NSK][5]
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[6] = self.T233_NSK[i_d_NSK][4]
                        rez[7] = self.T233_NSK[i_d_NSK][5]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQI":
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[6] = self.T233_NSK[i_d_NSK][4]
                        rez[7] = self.T233_NSK[i_d_NSK][5]
        if self.ui.bearingTypeListWidget.currentItem().text() == "Single row":
            i_d_BCT = self.searchCol_OC_LITE(d, self.T16_ISO492_2023, 0, 1)
            if i_d_BCT == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.innerRingSymmetryListWidget.currentRow() == 0:
                    rez[2] = self.T16_ISO492_2023[i_d_BCT][2]
                    rez[3] = self.T16_ISO492_2023[i_d_BCT][3]
                else:
                    rez[4] = self.T16_ISO492_2023[i_d_BCT][4]
                    rez[5] = self.T16_ISO492_2023[i_d_BCT][5]
                if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                    rez[6] = self.T16_ISO492_2023[i_d_BCT][10]
                    rez[7] = self.T16_ISO492_2023[i_d_BCT][11]
                rez[8] = self.T16_ISO492_2023[i_d_BCT][14]
                rez[9] = self.T16_ISO492_2023[i_d_BCT][15]

        return rez

    def get_t_delta_ir_P6X(self, d):
        rez = []
        for i in range(12):
            rez.append("—")
        i_d = self.searchCol_OC_LITE(d, self.T14_ISO492_2023, 0, 1)
        if i_d == "N/A":
            rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
            return rez
        else:
            rez[0] = self.T14_ISO492_2023[i_d][2]
            rez[1] = self.T14_ISO492_2023[i_d][3]
        if self.ui.bearingTypeListWidget.currentItem().text() != "Single row":
            i_d_NSK = self.searchCol_OC_LITE(d, self.T233_NSK, 0, 1)
            if i_d_NSK == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDO":
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[4] = self.T233_NSK[i_d_NSK][2]
                        rez[5] = self.T233_NSK[i_d_NSK][3]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDI":
                    i_d_BCT = self.searchCol_OC_LITE(d, self.T17_ISO492_2023, 0, 1)
                    if i_d_BCT == "N/A":
                        rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                        return rez
                    else:
                        rez[2] = self.T17_ISO492_2023[i_d_BCT][2]
                        rez[3] = self.T17_ISO492_2023[i_d_BCT][3]
                        if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                            rez[6] = self.T233_NSK[i_d_NSK][2]
                            rez[7] = self.T233_NSK[i_d_NSK][3]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQO":
                    rez[2] = self.T233_NSK[i_d_NSK][4]
                    rez[3] = self.T233_NSK[i_d_NSK][5]
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[6] = self.T233_NSK[i_d_NSK][4]
                        rez[7] = self.T233_NSK[i_d_NSK][5]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQI":
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[6] = self.T233_NSK[i_d_NSK][4]
                        rez[7] = self.T233_NSK[i_d_NSK][5]
        if self.ui.bearingTypeListWidget.currentItem().text() == "Single row":
            i_d_BCT = self.searchCol_OC_LITE(d, self.T17_ISO492_2023, 0, 1)
            if i_d_BCT == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.innerRingSymmetryListWidget.currentRow() == 0:
                    rez[2] = self.T17_ISO492_2023[i_d_BCT][2]
                    rez[3] = self.T17_ISO492_2023[i_d_BCT][3]
                else:
                    rez[4] = self.T17_ISO492_2023[i_d_BCT][4]
                    rez[5] = self.T17_ISO492_2023[i_d_BCT][5]
                if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                    rez[6] = self.T17_ISO492_2023[i_d_BCT][10]
                    rez[7] = self.T17_ISO492_2023[i_d_BCT][11]
                rez[8] = self.T17_ISO492_2023[i_d_BCT][14]
                rez[9] = self.T17_ISO492_2023[i_d_BCT][15]
        return rez

    def get_t_delta_ir_P5(self, d):
        rez = []
        for i in range(12):
            rez.append("—")
        i_d = self.searchCol_OC_LITE(d, self.T18_ISO492_2023, 0, 1)
        if i_d == "N/A":
            rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
            return rez
        else:
            rez[0] = self.T18_ISO492_2023[i_d][2]
            rez[1] = self.T18_ISO492_2023[i_d][3]
        if self.ui.bearingTypeListWidget.currentItem().text() != "Single row":
            i_d_NSK = self.searchCol_OC_LITE(d, self.T233_NSK, 0, 1)
            if i_d_NSK == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDO":
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[6] = self.T233_NSK[i_d_NSK][2]
                        rez[7] = self.T233_NSK[i_d_NSK][3]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDI":
                    i_d_BCT = self.searchCol_OC_LITE(d, self.T20_ISO492_2023, 0, 1)
                    if i_d_BCT == "N/A":
                        rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                        return rez
                    else:
                        rez[2] = self.T20_ISO492_2023[i_d_BCT][2]
                        rez[3] = self.T20_ISO492_2023[i_d_BCT][3]
                        if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                            rez[6] = self.T233_NSK[i_d_NSK][2]
                            rez[7] = self.T233_NSK[i_d_NSK][3]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQO":
                    rez[2] = self.T233_NSK[i_d_NSK][4]
                    rez[3] = self.T233_NSK[i_d_NSK][5]
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[6] = self.T233_NSK[i_d_NSK][4]
                        rez[7] = self.T233_NSK[i_d_NSK][5]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQI":
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[6] = self.T233_NSK[i_d_NSK][4]
                        rez[7] = self.T233_NSK[i_d_NSK][5]
        if self.ui.bearingTypeListWidget.currentItem().text() == "Single row":
            i_d_BCT = self.searchCol_OC_LITE(d, self.T20_ISO492_2023, 0, 1)
            if i_d_BCT == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.innerRingSymmetryListWidget.currentRow() == 0:
                    rez[2] = self.T20_ISO492_2023[i_d_BCT][2]
                    rez[3] = self.T20_ISO492_2023[i_d_BCT][3]
                else:
                    rez[4] = self.T20_ISO492_2023[i_d_BCT][4]
                    rez[5] = self.T20_ISO492_2023[i_d_BCT][5]
                if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                    rez[6] = self.T20_ISO492_2023[i_d_BCT][10]
                    rez[7] = self.T20_ISO492_2023[i_d_BCT][11]
                rez[8] = self.T20_ISO492_2023[i_d_BCT][14]
                rez[9] = self.T20_ISO492_2023[i_d_BCT][15]
        return rez

    def get_t_delta_ir_P4(self, d):
        rez = []
        for i in range(12):
            rez.append("—")
        i_d = self.searchCol_OC_LITE(d, self.T21_ISO492_2023, 0, 1)
        if i_d == "N/A":
            rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
            return rez
        else:
            S=self.ui.stiffnessLineEdit.text()
            if S=="C" or S=="S":
                rez[10] = self.T21_ISO492_2023[i_d][4]
                rez[11] = self.T21_ISO492_2023[i_d][5]
            else:
                rez[0] = self.T21_ISO492_2023[i_d][2]
                rez[1] = self.T21_ISO492_2023[i_d][3]
        if self.ui.bearingTypeListWidget.currentItem().text() != "Single row":
            i_d_NSK = self.searchCol_OC_LITE(d, self.T233_NSK, 0, 1)
            if i_d_NSK == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDO":
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[6] = self.T233_NSK[i_d_NSK][2]
                        rez[7] = self.T233_NSK[i_d_NSK][3]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDI":
                    i_d_BCT = self.searchCol_OC_LITE(d, self.T23_ISO492_2023, 0, 1)
                    if i_d_BCT == "N/A":
                        rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                        return rez
                    else:
                        rez[2] = self.T23_ISO492_2023[i_d_BCT][2]
                        rez[3] = self.T23_ISO492_2023[i_d_BCT][3]
                        if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                            rez[6] = self.T233_NSK[i_d_NSK][2]
                            rez[7] = self.T233_NSK[i_d_NSK][3]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQO":
                    rez[2] = self.T233_NSK[i_d_NSK][4]
                    rez[3] = self.T233_NSK[i_d_NSK][5]
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[6] = self.T233_NSK[i_d_NSK][4]
                        rez[7] = self.T233_NSK[i_d_NSK][5]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQI":
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[6] = self.T233_NSK[i_d_NSK][4]
                        rez[7] = self.T233_NSK[i_d_NSK][5]
        if self.ui.bearingTypeListWidget.currentItem().text() == "Single row":
            i_d_BCT = self.searchCol_OC_LITE(d, self.T23_ISO492_2023, 0, 1)
            if i_d_BCT == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.innerRingSymmetryListWidget.currentRow()==0:
                    rez[2] = self.T23_ISO492_2023[i_d_BCT][2]
                    rez[3] = self.T23_ISO492_2023[i_d_BCT][3]
                else:
                    rez[4] = self.T23_ISO492_2023[i_d_BCT][4]
                    rez[5] = self.T23_ISO492_2023[i_d_BCT][5]
                if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                    rez[6] = self.T23_ISO492_2023[i_d_BCT][10]
                    rez[7] = self.T23_ISO492_2023[i_d_BCT][11]
                rez[8] = self.T23_ISO492_2023[i_d_BCT][14]
                rez[9] = self.T23_ISO492_2023[i_d_BCT][15]
        return rez

    def get_t_delta_ir_P2(self, d):
        rez = []
        for i in range(12):
            rez.append("—")
        i_d = self.searchCol_OC_LITE(d, self.T24_ISO492_2023, 0, 1)
        if i_d == "N/A":
            rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
            return rez
        else:
            S=self.ui.stiffnessLineEdit.text()
            if S=="C" or S=="S":
                rez[10] = self.T24_ISO492_2023[i_d][4]
                rez[11] = self.T24_ISO492_2023[i_d][5]
            else:
                rez[0] = self.T24_ISO492_2023[i_d][2]
                rez[1] = self.T24_ISO492_2023[i_d][3]
        if self.ui.bearingTypeListWidget.currentItem().text() != "Single row":
            i_d_NSK = self.searchCol_OC_LITE(d, self.T233_NSK, 0, 1)
            if i_d_NSK == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDO":
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[6] = self.T233_NSK[i_d_NSK][2]
                        rez[7] = self.T233_NSK[i_d_NSK][3]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDI":
                    i_d_BCT = self.searchCol_OC_LITE(d, self.T26_ISO492_2023, 0, 1)
                    if i_d_BCT == "N/A":
                        rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                        return rez
                    else:
                        rez[2] = self.T26_ISO492_2023[i_d_BCT][2]
                        rez[3] = self.T26_ISO492_2023[i_d_BCT][3]
                        if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                            rez[6] = self.T233_NSK[i_d_NSK][2]
                            rez[7] = self.T233_NSK[i_d_NSK][3]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQO":
                    rez[2] = self.T233_NSK[i_d_NSK][4]
                    rez[3] = self.T233_NSK[i_d_NSK][5]
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[6] = self.T233_NSK[i_d_NSK][4]
                        rez[7] = self.T233_NSK[i_d_NSK][5]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQI":
                    if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                        rez[6] = self.T233_NSK[i_d_NSK][4]
                        rez[7] = self.T233_NSK[i_d_NSK][5]
        if self.ui.bearingTypeListWidget.currentItem().text() == "Single row":
            i_d_BCT = self.searchCol_OC_LITE(d, self.T26_ISO492_2023, 0, 1)
            if i_d_BCT == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.innerRingSymmetryListWidget.currentRow()==0:
                    rez[2] = self.T26_ISO492_2023[i_d_BCT][2]
                    rez[3] = self.T26_ISO492_2023[i_d_BCT][3]
                else:
                    rez[4] = self.T26_ISO492_2023[i_d_BCT][4]
                    rez[5] = self.T26_ISO492_2023[i_d_BCT][5]
                if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                    rez[6] = self.T26_ISO492_2023[i_d_BCT][10]
                    rez[7] = self.T26_ISO492_2023[i_d_BCT][11]
                rez[8] = self.T26_ISO492_2023[i_d_BCT][14]
                rez[9] = self.T26_ISO492_2023[i_d_BCT][15]
        return rez

    def get_t_delta_ir(self, d):
        if self.ui.precisionListWidget.currentRow() == 0:
            t_delta_ir = self.get_t_delta_ir_P0(d)
        elif self.ui.precisionListWidget.currentRow() == 1:
            t_delta_ir = self.get_t_delta_ir_P6X(d)
        elif self.ui.precisionListWidget.currentRow() == 2:
            t_delta_ir = self.get_t_delta_ir_P5(d)
        elif self.ui.precisionListWidget.currentRow() == 3:
            t_delta_ir = self.get_t_delta_ir_P4(d)
        elif self.ui.precisionListWidget.currentRow() == 4:
            t_delta_ir = self.get_t_delta_ir_P2(d)
        return t_delta_ir

    def get_t_V_ir_P0(self, d, S):
        rez = []
        for i in range(5):
            rez.append("—")
        i_d = self.searchCol_OC_LITE(d, self.T14_ISO492_2023, 0, 1)
        if i_d == "N/A":
            rez = ["N/A", "N/A", "N/A", "N/A", "N/A"]
            return rez
        else:
            if S == "A":
                rez[0] = self.T14_ISO492_2023[i_d][4]
            elif S == "B":
                rez[0] = self.T14_ISO492_2023[i_d][5]
            elif S == "C":
                rez[0] = self.T14_ISO492_2023[i_d][6]
            else:
                rez[0] = self.T14_ISO492_2023[i_d][7]
            rez[1] = self.T14_ISO492_2023[i_d][8]
            rez[2] = self.T14_ISO492_2023[i_d][9]
        return rez

    # get_t_V_ir_P6X is the same as get_t_V_ir_P0, so that function will be called

    def get_t_V_ir_P5(self, d, S):
        rez = []
        for i in range(5):
            rez.append("—")
        i_d = self.searchCol_OC_LITE(d, self.T18_ISO492_2023, 0, 1)
        if i_d == "N/A":
            rez = ["N/A", "N/A", "N/A", "N/A", "N/A"]
            return rez
        else:
            if S == "A":
                rez[0] = self.T18_ISO492_2023[i_d][4]
            elif S == "B":
                rez[0] = self.T18_ISO492_2023[i_d][5]
            elif S == "C":
                rez[0] = self.T18_ISO492_2023[i_d][6]
            else:
                rez[0] = self.T18_ISO492_2023[i_d][7]
            rez[1] = self.T18_ISO492_2023[i_d][8]
            rez[2] = self.T18_ISO492_2023[i_d][9]
            rez[3] = self.T18_ISO492_2023[i_d][10]
        return rez

    def get_t_V_ir_P4(self, d, S):
        rez = []
        for i in range(5):
            rez.append("—")
        i_d = self.searchCol_OC_LITE(d, self.T21_ISO492_2023, 0, 1)
        if i_d == "N/A":
            rez = ["N/A", "N/A", "N/A", "N/A", "N/A"]
            return rez
        else:
            if S=="A":
                rez[0] = self.T21_ISO492_2023[i_d][6]
            elif S=="B":
                rez[0] = self.T21_ISO492_2023[i_d][7]
            elif S=="C":
                rez[0] = self.T21_ISO492_2023[i_d][8]
            else:
                rez[0] = self.T21_ISO492_2023[i_d][9]
            rez[1] = self.T21_ISO492_2023[i_d][10]
            rez[2] = self.T21_ISO492_2023[i_d][11]
            rez[3] = self.T21_ISO492_2023[i_d][12]
            rez[4] = self.T21_ISO492_2023[i_d][13]
        return rez

    def get_t_V_ir_P2(self, d, S):
        rez = []
        for i in range(5):
            rez.append("—")
        i_d = self.searchCol_OC_LITE(d, self.T24_ISO492_2023, 0, 1)
        if i_d == "N/A":
            rez = ["N/A", "N/A", "N/A", "N/A", "N/A"]
            return rez
        else:
            rez[0] = self.T24_ISO492_2023[i_d][6]
            rez[1] = self.T24_ISO492_2023[i_d][7]
            rez[2] = self.T24_ISO492_2023[i_d][8]
            rez[3] = self.T24_ISO492_2023[i_d][9]
            rez[4] = self.T24_ISO492_2023[i_d][10]
        return rez

    def get_t_V_ir(self, d, S):
        if self.ui.precisionListWidget.currentRow() == 0:
            t_V_ir = self.get_t_V_ir_P0(d, S)
        elif self.ui.precisionListWidget.currentRow() == 1:
            t_V_ir = self.get_t_V_ir_P0(d,
                                        S)  # get_t_V_ir_P6X is the same as get_t_V_ir_P0, so get_t_V_ir_P0 is called in this case
        elif self.ui.precisionListWidget.currentRow() == 2:
            t_V_ir = self.get_t_V_ir_P5(d, S)
        elif self.ui.precisionListWidget.currentRow() == 3:
            t_V_ir = self.get_t_V_ir_P4(d, S)
        elif self.ui.precisionListWidget.currentRow() == 4:
            t_V_ir = self.get_t_V_ir_P2(d, S)
        return t_V_ir

    def get_t_delta_or_P0(self, D, d):
        rez = []
        for i in range(10):
            rez.append("—")
        i_D = self.searchCol_OC_LITE(D, self.T15_ISO492_2023, 0, 1)
        if i_D == "N/A":
            # rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
            # return rez
            rez[0] = "N/A"
            rez[1] = "N/A"
        else:
            rez[0] = self.T15_ISO492_2023[i_D][2]
            rez[1] = self.T15_ISO492_2023[i_D][3]
        if self.ui.bearingTypeListWidget.currentItem().text() != "Single row":
            i_d_NSK = self.searchCol_OC_LITE(d, self.T233_NSK, 0, 1)
            if i_d_NSK == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDO":
                    i_d_BCT = self.searchCol_OC_LITE(d, self.T16_ISO492_2023, 0, 1)
                    if i_d_BCT == "N/A":
                        rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                        return rez
                    else:
                        rez[2] = self.T16_ISO492_2023[i_d_BCT][6]
                        rez[3] = self.T16_ISO492_2023[i_d_BCT][7]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQI":
                    rez[2] = self.T233_NSK[i_d_NSK][4]
                    rez[3] = self.T233_NSK[i_d_NSK][5]
        if self.ui.bearingTypeListWidget.currentItem().text() == "Single row":
            i_d_BCT = self.searchCol_OC_LITE(d, self.T16_ISO492_2023, 0, 1)
            if i_d_BCT == "N/A":
                # rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                # return rez
                rez[0]="N/A"
                rez[1] = "N/A"
            if self.ui.outerRingSymmetryListWidget.currentRow() == 0:
                rez[2] = self.T16_ISO492_2023[i_d_BCT][6]
                rez[3] = self.T16_ISO492_2023[i_d_BCT][7]
            else:
                rez[4] = self.T16_ISO492_2023[i_d_BCT][8]
                rez[5] = self.T16_ISO492_2023[i_d_BCT][9]
            if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                rez[6] = self.T16_ISO492_2023[i_d_BCT][16]
                rez[7] = self.T16_ISO492_2023[i_d_BCT][17]
        return rez

    def get_t_delta_or_P6X(self, D, d):
        rez = []
        for i in range(10):
            rez.append("—")
        i_D = self.searchCol_OC_LITE(D, self.T15_ISO492_2023, 0, 1)
        if i_D == "N/A":
            # rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
            # return rez
            rez[0] = "N/A"
            rez[1] = "N/A"
        else:
            rez[0] = self.T15_ISO492_2023[i_D][2]
            rez[1] = self.T15_ISO492_2023[i_D][3]
        if self.ui.bearingTypeListWidget.currentItem().text() != "Single row":
            i_d_NSK = self.searchCol_OC_LITE(d, self.T233_NSK, 0, 1)
            if i_d_NSK == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDO":
                    i_d_BCT = self.searchCol_OC_LITE(d, self.T17_ISO492_2023, 0, 1)
                    if i_d_BCT == "N/A":
                        rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                        return rez
                    else:
                        rez[2] = self.T17_ISO492_2023[i_d_BCT][6]
                        rez[3] = self.T17_ISO492_2023[i_d_BCT][7]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQI":
                    rez[2] = self.T233_NSK[i_d_NSK][4]
                    rez[3] = self.T233_NSK[i_d_NSK][5]
        if self.ui.bearingTypeListWidget.currentItem().text() == "Single row":
            i_d_BCT = self.searchCol_OC_LITE(d, self.T17_ISO492_2023, 0, 1)
            if i_d_BCT == "N/A":
                # rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                # return rez
                rez[0] = "N/A"
                rez[1] = "N/A"
            if self.ui.outerRingSymmetryListWidget.currentRow() == 0:
                rez[2] = self.T17_ISO492_2023[i_d_BCT][6]
                rez[3] = self.T17_ISO492_2023[i_d_BCT][7]
            else:
                rez[4] = self.T17_ISO492_2023[i_d_BCT][8]
                rez[5] = self.T17_ISO492_2023[i_d_BCT][9]
            if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                rez[6] = self.T17_ISO492_2023[i_d_BCT][16]
                rez[7] = self.T17_ISO492_2023[i_d_BCT][17]
        return rez

    def get_t_delta_or_P5(self, D, d):
        rez = []
        for i in range(10):
            rez.append("—")
        i_D = self.searchCol_OC_LITE(D, self.T19_ISO492_2023, 0, 1)
        if i_D == "N/A":
            # rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
            # return rez
            rez[0] = "N/A"
            rez[1] = "N/A"
        else:
            rez[0] = self.T19_ISO492_2023[i_D][2]
            rez[1] = self.T19_ISO492_2023[i_D][3]
        if self.ui.bearingTypeListWidget.currentItem().text() != "Single row":
            i_d_NSK = self.searchCol_OC_LITE(d, self.T233_NSK, 0, 1)
            if i_d_NSK == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDO":
                    i_d_BCT = self.searchCol_OC_LITE(d, self.T20_ISO492_2023, 0, 1)
                    if i_d_BCT == "N/A":
                        rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                        return rez
                    else:
                        rez[2] = self.T20_ISO492_2023[i_d_BCT][6]
                        rez[3] = self.T20_ISO492_2023[i_d_BCT][7]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQI":
                    rez[2] = self.T233_NSK[i_d_NSK][4]
                    rez[3] = self.T233_NSK[i_d_NSK][5]
        if self.ui.bearingTypeListWidget.currentItem().text() == "Single row":
            i_d_BCT = self.searchCol_OC_LITE(d, self.T20_ISO492_2023, 0, 1)
            if i_d_BCT == "N/A":
                # rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                # return rez
                rez[0] = "N/A"
                rez[1] = "N/A"
            if self.ui.outerRingSymmetryListWidget.currentRow() == 0:
                rez[2] = self.T20_ISO492_2023[i_d_BCT][6]
                rez[3] = self.T20_ISO492_2023[i_d_BCT][7]
            else:
                rez[4] = self.T20_ISO492_2023[i_d_BCT][8]
                rez[5] = self.T20_ISO492_2023[i_d_BCT][9]
            if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                rez[6] = self.T20_ISO492_2023[i_d_BCT][16]
                rez[7] = self.T20_ISO492_2023[i_d_BCT][17]
        return rez

    def get_t_delta_or_P4(self, D, d):
        rez = []
        for i in range(10):
            rez.append("—")
        i_D = self.searchCol_OC_LITE(D, self.T22_ISO492_2023, 0, 1)
        if i_D == "N/A":
            # rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
            # return rez
            rez[0] = "N/A"
            rez[1] = "N/A"
        else:
            S=self.ui.stiffnessLineEdit.text()
            if S=="C" or S=="S":
                rez[8] = self.T22_ISO492_2023[i_D][4]
                rez[9] = self.T22_ISO492_2023[i_D][5]
            else:
                rez[0] = self.T22_ISO492_2023[i_D][2]
                rez[1] = self.T22_ISO492_2023[i_D][3]
        if self.ui.bearingTypeListWidget.currentItem().text() != "Single row":
            i_d_NSK = self.searchCol_OC_LITE(d, self.T233_NSK, 0, 1)
            if i_d_NSK == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDO":
                    i_d_BCT = self.searchCol_OC_LITE(d, self.T23_ISO492_2023, 0, 1)
                    if i_d_BCT == "N/A":
                        rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                        return rez
                    else:
                        rez[2] = self.T23_ISO492_2023[i_d_BCT][6]
                        rez[3] = self.T23_ISO492_2023[i_d_BCT][7]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQI":
                    rez[2] = self.T233_NSK[i_d_NSK][4]
                    rez[3] = self.T233_NSK[i_d_NSK][5]
        if self.ui.bearingTypeListWidget.currentItem().text() == "Single row":
            i_d_BCT = self.searchCol_OC_LITE(d, self.T23_ISO492_2023, 0, 1)
            if i_d_BCT == "N/A":
                # rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                # return rez
                rez[0] = "N/A"
                rez[1] = "N/A"
            if self.ui.outerRingSymmetryListWidget.currentRow()==0:
                rez[2] = self.T23_ISO492_2023[i_d_BCT][6]
                rez[3] = self.T23_ISO492_2023[i_d_BCT][7]
            else:
                rez[4] = self.T23_ISO492_2023[i_d_BCT][8]
                rez[5] = self.T23_ISO492_2023[i_d_BCT][9]
            if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                rez[6] = self.T23_ISO492_2023[i_d_BCT][16]
                rez[7] = self.T23_ISO492_2023[i_d_BCT][17]
        return rez

    def get_t_delta_or_P2(self, D, d):
        rez = []
        for i in range(10):
            rez.append("—")
        i_D = self.searchCol_OC_LITE(D, self.T25_ISO492_2023, 0, 1)
        if i_D == "N/A":
            # rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
            # return rez
            rez[0] = "N/A"
            rez[1] = "N/A"
        else:
            S=self.ui.stiffnessLineEdit.text()
            if S=="C" or S=="S":
                rez[8] = self.T25_ISO492_2023[i_D][4]
                rez[9] = self.T25_ISO492_2023[i_D][5]
            else:
                rez[0] = self.T25_ISO492_2023[i_D][2]
                rez[1] = self.T25_ISO492_2023[i_D][3]
        if self.ui.bearingTypeListWidget.currentItem().text() != "Single row":
            i_d_NSK = self.searchCol_OC_LITE(d, self.T233_NSK, 0, 1)
            if i_d_NSK == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.bearingTypeListWidget.currentItem().text() == "TDO":
                    i_d_BCT = self.searchCol_OC_LITE(d, self.T26_ISO492_2023, 0, 1)
                    if i_d_BCT == "N/A":
                        rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                        return rez
                    else:
                        rez[2] = self.T26_ISO492_2023[i_d_BCT][6]
                        rez[3] = self.T26_ISO492_2023[i_d_BCT][7]
                if self.ui.bearingTypeListWidget.currentItem().text() == "TQI":
                    rez[2] = self.T233_NSK[i_d_NSK][4]
                    rez[3] = self.T233_NSK[i_d_NSK][5]
        if self.ui.bearingTypeListWidget.currentItem().text() == "Single row":
            i_d_BCT = self.searchCol_OC_LITE(d, self.T26_ISO492_2023, 0, 1)
            if i_d_BCT == "N/A":
                # rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                # return rez
                rez[0] = "N/A"
                rez[1] = "N/A"
            if self.ui.outerRingSymmetryListWidget.currentRow()==0:
                rez[2] = self.T26_ISO492_2023[i_d_BCT][6]
                rez[3] = self.T26_ISO492_2023[i_d_BCT][7]
            else:
                rez[4] = self.T26_ISO492_2023[i_d_BCT][8]
                rez[5] = self.T26_ISO492_2023[i_d_BCT][9]
            if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                rez[6] = self.T26_ISO492_2023[i_d_BCT][16]
                rez[7] = self.T26_ISO492_2023[i_d_BCT][17]
        return rez

    def get_t_delta_or(self, D, d):
        if self.ui.precisionListWidget.currentRow() == 0:
            t_delta_or = self.get_t_delta_or_P0(D, d)
        elif self.ui.precisionListWidget.currentRow() == 1:
            t_delta_or = self.get_t_delta_or_P6X(D, d)
        elif self.ui.precisionListWidget.currentRow() == 2:
            t_delta_or = self.get_t_delta_or_P5(D, d)
        elif self.ui.precisionListWidget.currentRow() == 3:
            t_delta_or = self.get_t_delta_or_P4(D, d)
        elif self.ui.precisionListWidget.currentRow() == 4:
            t_delta_or = self.get_t_delta_or_P2(D, d)
        return t_delta_or

    def get_t_V_or_P0(self, D, S):
        rez = []
        for i in range(5):
            rez.append("—")
        i_D = self.searchCol_OC_LITE(D, self.T15_ISO492_2023, 0, 1)
        if i_D == "N/A":
            rez = ["N/A", "N/A", "N/A", "N/A", "N/A"]
            return rez
        else:
            if S == "A":
                rez[0] = self.T15_ISO492_2023[i_D][4]
            elif S == "B":
                rez[0] = self.T15_ISO492_2023[i_D][5]
            elif S == "C":
                rez[0] = self.T15_ISO492_2023[i_D][6]
            else:
                rez[0] = self.T15_ISO492_2023[i_D][7]
            rez[1] = self.T15_ISO492_2023[i_D][8]
            rez[2] = self.T15_ISO492_2023[i_D][9]
        return rez

    # get_t_V_or_P6X is the same as get_t_V_or_P0, so that function will be called

    def get_t_V_or_P5(self, D, S):
        rez = []
        for i in range(5):
            rez.append("—")
        i_D = self.searchCol_OC_LITE(D, self.T19_ISO492_2023, 0, 1)
        if i_D == "N/A":
            rez = ["N/A", "N/A", "N/A", "N/A", "N/A"]
            return rez
        else:
            if S == "A":
                rez[0] = self.T19_ISO492_2023[i_D][4]
            elif S == "B":
                rez[0] = self.T19_ISO492_2023[i_D][5]
            elif S == "C":
                rez[0] = self.T19_ISO492_2023[i_D][6]
            else:
                rez[0] = self.T19_ISO492_2023[i_D][7]
            rez[1] = self.T19_ISO492_2023[i_D][8]
            rez[2] = self.T19_ISO492_2023[i_D][9]
            if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                rez[3] = self.T19_ISO492_2023[i_D][10]
        return rez

    def get_t_V_or_P4(self, D, S):
        rez = []
        for i in range(5):
            rez.append("—")
        i_D = self.searchCol_OC_LITE(D, self.T22_ISO492_2023, 0, 1)
        if i_D == "N/A":
            rez = ["N/A", "N/A", "N/A", "N/A", "N/A"]
            return rez
        else:
            if S=="A":
                rez[0] = self.T22_ISO492_2023[i_D][6]
            elif S=="B":
                rez[0] = self.T22_ISO492_2023[i_D][7]
            elif S=="C":
                rez[0] = self.T22_ISO492_2023[i_D][8]
            else:
                rez[0] = self.T22_ISO492_2023[i_D][9]
            rez[1] = self.T22_ISO492_2023[i_D][10]
            rez[2] = self.T22_ISO492_2023[i_D][11]
            if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                rez[3] = self.T22_ISO492_2023[i_D][12]
                rez[4] = self.T22_ISO492_2023[i_D][14]
        return rez

    def get_t_V_or_P2(self, D, S):
        rez = []
        for i in range(5):
            rez.append("—")
        i_D = self.searchCol_OC_LITE(D, self.T25_ISO492_2023, 0, 1)
        if i_D == "N/A":
            rez = ["N/A", "N/A", "N/A", "N/A", "N/A"]
            return rez
        else:
            rez[0] = self.T25_ISO492_2023[i_D][6]
            rez[1] = self.T25_ISO492_2023[i_D][7]
            rez[2] = self.T25_ISO492_2023[i_D][8]
            if self.ui.flangePresenceListWidget.currentItem().text() == "Normal":
                rez[3] = self.T25_ISO492_2023[i_D][9]
                rez[4] = self.T25_ISO492_2023[i_D][11]
        return rez

    def get_t_V_or(self, D, S):
        if self.ui.precisionListWidget.currentRow() == 0:
            t_V_or = self.get_t_V_or_P0(D, S)
        elif self.ui.precisionListWidget.currentRow() == 1:
            t_V_or = self.get_t_V_or_P0(D,
                                        S)  # get_t_V_ir_P6X is the same as get_t_V_ir_P0, so get_t_V_ir_P0 is called in this case
        elif self.ui.precisionListWidget.currentRow() == 2:
            t_V_or = self.get_t_V_or_P5(D, S)
        elif self.ui.precisionListWidget.currentRow() == 3:
            t_V_or = self.get_t_V_or_P4(D, S)
        elif self.ui.precisionListWidget.currentRow() == 4:
            t_V_or = self.get_t_V_or_P2(D, S)
        return t_V_or

    def get_t_delta_F_P0(self, D1, d):
        rez = []
        for i in range(6):
            rez.append("—")
        if self.ui.flangePresenceListWidget.currentItem().text() == "Flanged outer ring":
            i_D1 = self.searchCol_OC_LITE(D1, self.T27_ISO492_2023, 0, 1)
            if i_D1 == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.flangeTypeListWidget.currentItem().text() == "Locating flange":
                    rez[0] = self.T27_ISO492_2023[i_D1][2]
                    rez[1] = self.T27_ISO492_2023[i_D1][3]
                else:
                    rez[0] = self.T27_ISO492_2023[i_D1][4]
                    rez[1] = self.T27_ISO492_2023[i_D1][5]
                i_d_BCT = self.searchCol_OC_LITE(d, self.T16_ISO492_2023, 0, 1)
                if i_d_BCT == "N/A":
                    rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                    return rez
                else:
                    rez[2] = self.T16_ISO492_2023[i_d_BCT][10]
                    rez[3] = self.T16_ISO492_2023[i_d_BCT][11]
                    rez[4] = self.T16_ISO492_2023[i_d_BCT][18]
                    rez[5] = self.T16_ISO492_2023[i_d_BCT][19]
        return rez

    def get_t_delta_F_P6X(self, D1, d):
        rez = []
        for i in range(6):
            rez.append("—")
        if self.ui.flangePresenceListWidget.currentItem().text() == "Flanged outer ring":
            i_D1 = self.searchCol_OC_LITE(D1, self.T27_ISO492_2023, 0, 1)
            if i_D1 == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.flangeTypeListWidget.currentItem().text() == "Locating flange":
                    rez[0] = self.T27_ISO492_2023[i_D1][2]
                    rez[1] = self.T27_ISO492_2023[i_D1][3]
                else:
                    rez[0] = self.T27_ISO492_2023[i_D1][4]
                    rez[1] = self.T27_ISO492_2023[i_D1][5]
                i_d_BCT = self.searchCol_OC_LITE(d, self.T17_ISO492_2023, 0, 1)
                if i_d_BCT == "N/A":
                    rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                    return rez
                else:
                    rez[2] = self.T17_ISO492_2023[i_d_BCT][12]
                    rez[3] = self.T17_ISO492_2023[i_d_BCT][13]
                    rez[4] = self.T17_ISO492_2023[i_d_BCT][18]
                    rez[5] = self.T17_ISO492_2023[i_d_BCT][19]
        return rez

    def get_t_delta_F_P5(self, D1, d):
        rez = []
        for i in range(6):
            rez.append("—")
        if self.ui.flangePresenceListWidget.currentItem().text() == "Flanged outer ring":
            i_D1 = self.searchCol_OC_LITE(D1, self.T27_ISO492_2023, 0, 1)
            if i_D1 == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.flangeTypeListWidget.currentItem().text() == "Locating flange":
                    rez[0] = self.T27_ISO492_2023[i_D1][2]
                    rez[1] = self.T27_ISO492_2023[i_D1][3]
                else:
                    rez[0] = self.T27_ISO492_2023[i_D1][4]
                    rez[1] = self.T27_ISO492_2023[i_D1][5]
                i_d_BCT = self.searchCol_OC_LITE(d, self.T20_ISO492_2023, 0, 1)
                if i_d_BCT == "N/A":
                    rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                    return rez
                else:
                    rez[2] = self.T20_ISO492_2023[i_d_BCT][12]
                    rez[3] = self.T20_ISO492_2023[i_d_BCT][13]
                    rez[4] = self.T20_ISO492_2023[i_d_BCT][18]
                    rez[5] = self.T20_ISO492_2023[i_d_BCT][19]
        return rez

    def get_t_delta_F_P4(self, D1, d):
        rez = []
        for i in range(6):
            rez.append("—")
        if self.ui.flangePresenceListWidget.currentItem().text() == "Flanged outer ring":
            i_D1 = self.searchCol_OC_LITE(D1, self.T27_ISO492_2023, 0, 1)
            if i_D1 == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.flangeTypeListWidget.currentItem().text() == "Locating flange":
                    rez[0] = self.T27_ISO492_2023[i_D1][2]
                    rez[1] = self.T27_ISO492_2023[i_D1][3]
                else:
                    rez[0] = self.T27_ISO492_2023[i_D1][4]
                    rez[1] = self.T27_ISO492_2023[i_D1][5]
                i_d_BCT = self.searchCol_OC_LITE(d, self.T23_ISO492_2023, 0, 1)
                if i_d_BCT == "N/A":
                    rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                    return rez
                else:
                    rez[2] = self.T23_ISO492_2023[i_d_BCT][12]
                    rez[3] = self.T23_ISO492_2023[i_d_BCT][13]
                    rez[4] = self.T23_ISO492_2023[i_d_BCT][18]
                    rez[5] = self.T23_ISO492_2023[i_d_BCT][19]
        return rez

    def get_t_delta_F_P2(self, D1, d):
        rez = []
        for i in range(6):
            rez.append("—")
        if self.ui.flangePresenceListWidget.currentItem().text() == "Flanged outer ring":
            i_D1 = self.searchCol_OC_LITE(D1, self.T27_ISO492_2023, 0, 1)
            if i_D1 == "N/A":
                rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                return rez
            else:
                if self.ui.flangeTypeListWidget.currentItem().text() == "Locating flange":
                    rez[0] = self.T27_ISO492_2023[i_D1][2]
                    rez[1] = self.T27_ISO492_2023[i_D1][3]
                else:
                    rez[0] = self.T27_ISO492_2023[i_D1][4]
                    rez[1] = self.T27_ISO492_2023[i_D1][5]
                i_d_BCT = self.searchCol_OC_LITE(d, self.T26_ISO492_2023, 0, 1)
                if i_d_BCT == "N/A":
                    rez = ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]
                    return rez
                else:
                    rez[2] = self.T26_ISO492_2023[i_d_BCT][12]
                    rez[3] = self.T26_ISO492_2023[i_d_BCT][13]
                    rez[4] = self.T26_ISO492_2023[i_d_BCT][18]
                    rez[5] = self.T26_ISO492_2023[i_d_BCT][19]
        return rez

    def get_t_delta_F(self, D1, d):
        if self.ui.precisionListWidget.currentRow() == 0:
            t_delta_F = self.get_t_delta_F_P0(D1, d)
        elif self.ui.precisionListWidget.currentRow() == 1:
            t_delta_F = self.get_t_delta_F_P6X(D1, d)
        elif self.ui.precisionListWidget.currentRow() == 2:
            t_delta_F = self.get_t_delta_F_P5(D1, d)
        elif self.ui.precisionListWidget.currentRow() == 3:
            t_delta_F = self.get_t_delta_F_P4(D1, d)
        elif self.ui.precisionListWidget.currentRow() == 4:
            t_delta_F = self.get_t_delta_F_P2(D1, d)
        return t_delta_F

    def get_t_V_F_P0(self, D,D1):
        rez = []
        for i in range(2):
            rez.append("—")
        return rez

    def get_t_V_F_P6X(self, D,D1):
        rez = []
        for i in range(2):
            rez.append("—")
        return rez

    def get_t_V_F_P5(self, D,D1):
        rez = []
        for i in range(2):
            rez.append("—")
        if self.ui.flangePresenceListWidget.currentItem().text() == "Flanged outer ring":
            i_D1 = self.searchCol_OC_LITE(D1, self.T27_ISO492_2023, 0, 1)
            if i_D1 == "N/A":
                rez = ["N/A", "N/A"]
                return rez
            else:
                i_D = self.searchCol_OC_LITE(D, self.T19_ISO492_2023, 0, 1)
                if i_D == "N/A":
                    rez = ["N/A", "N/A"]
                    return rez
                else:
                    rez[0] = self.T19_ISO492_2023[i_D][11]
        return rez

    def get_t_V_F_P4(self, D,D1):
        rez = []
        for i in range(2):
            rez.append("—")
        if self.ui.flangePresenceListWidget.currentItem().text() == "Flanged outer ring":
            i_D1 = self.searchCol_OC_LITE(D1, self.T27_ISO492_2023, 0, 1)
            if i_D1 == "N/A":
                rez = ["N/A", "N/A"]
                return rez
            else:
                i_D = self.searchCol_OC_LITE(D, self.T22_ISO492_2023, 0, 1)
                if i_D == "N/A":
                    rez = ["N/A", "N/A"]
                    return rez
                else:
                    rez[0] = self.T22_ISO492_2023[i_D][13]
                    rez[1] = self.T22_ISO492_2023[i_D][15]
        return rez

    def get_t_V_F_P2(self,D, D1):
        rez = []
        for i in range(2):
            rez.append("—")
        if self.ui.flangePresenceListWidget.currentItem().text() == "Flanged outer ring":
            i_D1 = self.searchCol_OC_LITE(D1, self.T27_ISO492_2023, 0, 1)
            if i_D1 == "N/A":
                rez = ["N/A", "N/A"]
                return rez
            else:
                i_D = self.searchCol_OC_LITE(D, self.T25_ISO492_2023, 0, 1)
                if i_D == "N/A":
                    rez = ["N/A", "N/A"]
                    return rez
                else:
                    rez[0] = self.T25_ISO492_2023[i_D][10]
                    rez[1] = self.T25_ISO492_2023[i_D][12]
        return rez

    def get_t_V_F(self,D, D1):
        if self.ui.precisionListWidget.currentRow() == 0:
            t_V_F = self.get_t_V_F_P0(D,D1)
        elif self.ui.precisionListWidget.currentRow() == 1:
            t_V_F = self.get_t_V_F_P6X(D,D1)
        elif self.ui.precisionListWidget.currentRow() == 2:
            t_V_F = self.get_t_V_F_P5(D,D1)
        elif self.ui.precisionListWidget.currentRow() == 3:
            t_V_F = self.get_t_V_F_P4(D,D1)
        elif self.ui.precisionListWidget.currentRow() == 4:
            t_V_F = self.get_t_V_F_P2(D,D1)
        return t_V_F

    def showError(self, errMsg):
        '''

        :param errMsg: error message to be shown in a popup
        :return:
        '''
        msg = QMessageBox()
        msg.setWindowTitle("Error")
        text = "<p align='center'>" + errMsg + "</p>"
        msg.setText(text)

        # for local run
        msg.setWindowIcon(QtGui.QIcon(resource_path(WINDOW_ICO_FILE_PATH)))
        msg.setIcon(QMessageBox.Critical)
        x = msg.exec_()

    def showMessage(self, mesg):
        '''

        :param errMsg: error message to be shown in a popup
        :return:
        '''
        msg = QMessageBox()
        msg.setWindowTitle("Printed successfully")
        text = "<p align='center'>" + mesg + "</p>"
        msg.setText(text)

        # for local run
        msg.setWindowIcon(QtGui.QIcon(resource_path(LOGO_FILE_PATH2)))
        msg.setIconPixmap(QPixmap(resource_path(WINDOW_ICO_FILE_PATH)))
        x = msg.exec_()

    def showWarning(self, warningText):
        text = "<p align='center'>" + warningText + "</p>"

        msg = QMessageBox.warning(self, "Warning", warningText)
        # msg.setWindowTitle("Warning")
        #
        # msg.setText(text)
        #
        # # for local run
        # msg.setWindowIcon(QtGui.QIcon(resource_path(LOGO_FILE_PATH2)))
        # #msg.setIcon(QMessageBox.warning)
        # x = msg.exec_()

    def updateWindowTitle(self):
        _translate = QtCore.QCoreApplication.translate
        self.ui.MainWindow.setWindowTitle(_translate("MainWindow",
                                                     self.programName + " - " + self.programVersion + " - " + self.c.fileName + " " + self.c.status))
        self.ui.menubar.setToolTip(
            self.programName + " - " + self.programVersion + " - " + self.c.fileName + " " + self.c.status)

    def loadLists(self):
        self.ui.innerRingSymmetryListWidget.setCurrentRow(0)
        self.ui.outerRingSymmetryListWidget.setCurrentRow(0)

        # bearingTypeListWidget
        self.ui.bearingTypeListWidget.clear()
        self.ui.bearingTypeListWidget.insertItem(0, "Single row")
        self.ui.bearingTypeListWidget.insertItem(1, "TDO")
        self.ui.bearingTypeListWidget.insertItem(2, "TDI")
        self.ui.bearingTypeListWidget.insertItem(3, "TQO")
        self.ui.bearingTypeListWidget.insertItem(4, "TQI")
        self.ui.bearingTypeListWidget.setCurrentRow(0)

        # singleDoubleDirectionListWidget
        self.ui.flangePresenceListWidget.clear()
        self.ui.flangePresenceListWidget.insertItem(0, "Normal")
        self.ui.flangePresenceListWidget.insertItem(1, "Flanged outer ring")
        self.ui.flangePresenceListWidget.setCurrentRow(0)

        # precisionListWidget
        self.ui.precisionListWidget.clear()
        self.ui.precisionListWidget.insertItem(0, "Normal  (P0)")
        self.ui.precisionListWidget.insertItem(1, "Class 6X (P6X)")
        self.ui.precisionListWidget.insertItem(2, "Class 5 (P5)")
        self.ui.precisionListWidget.insertItem(3, "Class 4 (P4)")
        self.ui.precisionListWidget.insertItem(4, "Class 2 (P2)")
        self.ui.precisionListWidget.setCurrentRow(0)

        # flangeTypeListWidget
        self.ui.flangeTypeListWidget.clear()

    def flangePresenceChanged(self):
        if self.ui.flangePresenceListWidget.currentRow() == 0:
            self.ui.flangeTypeListWidget.clear()
        elif self.ui.flangePresenceListWidget.currentRow() == 1:
            self.ui.flangeTypeListWidget.clear()
            self.ui.flangeTypeListWidget.insertItem(0, "Locating flange")
            self.ui.flangeTypeListWidget.insertItem(1, "Non-locating flange")
            self.ui.flangeTypeListWidget.setCurrentRow(0)

    def inputChanged(self):
        self.c.status = "*"
        self.updateWindowTitle()
        self.resetOutput()
        self.ui.errorLabel.setText("")

    def stiffnessChanged(self):
        self.c.status = "*"
        self.updateWindowTitle()
        self.ui.stiffnessLineEdit.setText("")
        self.ui.fsLineEdit.setText("")
        self.resetOutput()

    def resetClicked(self):
        self.resetInput()
        self.ui.errorLabel.setText("")
        self.resetOutput()
        self.c.reset()
        self.updateWindowTitle()

    def resetOutput(self):

        # Output Tab
        self.ui.innerRingTableWidget1.clearContents()
        self.customizeInnerRingTableWidget1()
        self.ui.innerRingTableWidget2.clearContents()
        self.customizeInnerRingTableWidget2()
        self.ui.outerRingTableWidget1.clearContents()
        self.customizeOuterRingTableWidget1()
        self.ui.outerRingTableWidget2.clearContents()
        self.customizeOuterRingTableWidget2()
        self.ui.flangeTableWidget1.clearContents()
        self.customizeFlangeTableWidget1()
        self.ui.flangeTableWidget2.clearContents()
        self.customizeFlangeTableWidget2()

        self.adjust_tables()

    def resetInput(self):
        '''
        resets all input fields
        :return:
        '''
        # input tab
        self.ui.partNoTextEdit.setText("")
        self.ui.stiffnessLineEdit.setText("")
        self.ui.fsLineEdit.setText("")
        self.ui.boreDiameterLineEdit.setText("")
        self.ui.outerDiameterLineEdit.setText("")
        self.ui.orFlangeDiameterLineEdit.setText("")
        self.loadLists()

    def print_tabs_to_pdf(self):
        try:
            # Set up printer
            printer = QPrinter(QPrinter.HighResolution)
            printer.setResolution(300)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setPageSize(QPrinter.A4)
            # Set the output file name
            filename, _ = QFileDialog.getSaveFileName(None, "Save PDF", "", "PDF files (*.pdf)")
            if filename:
                self.setStyleSheet("background-color: rgb(255,255,255);")
                printer.setOutputFileName(filename)
                printer.setPageMargins(10, 10, 10, 10, QPrinter.Millimeter)
                # Set up painter
                painter = QPainter()
                if not painter.begin(printer):
                    self.showError(str("Error: Could not open printer. Verify if the document is already open!"))
                    return
                # Loop through each tab to print it on a new page
                x = self.ui.tabWidget.count()
                for tab_index in range(self.ui.tabWidget.count()):
                    if (tab_index != 2):
                        self.ui.tabWidget.setCurrentIndex(tab_index)
                        tab_widget = self.ui.tabWidget.currentWidget()
                        if isinstance(tab_widget, QScrollArea):
                            tab_widget = tab_widget.widget()
                        tab_widget.adjustSize()
                        pixmap = QPixmap(tab_widget.size())
                        tab_widget.render(pixmap)
                        scaled_pixmap = pixmap.scaled(printer.pageRect().size(), Qt.KeepAspectRatio,
                                                      Qt.SmoothTransformation)
                        painter.drawPixmap(0, 0, scaled_pixmap)
                        if tab_index < self.ui.tabWidget.count() - 2:
                            printer.newPage()
                date = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
                painter.drawText(30, printer.pageRect().height() - 100,
                                 "Printed at " + date + " © RKB Bearings Industries Switzerland")
                painter.end()
                self.setStyleSheet("")
                self.showMessage(str("Document successfully printed as a PDF."))
                self.ui.tabWidget.setCurrentIndex(2)

        except Exception as ex:
            self.showError(str(ex))

    def moveEvent(self, event):
        """ Detect when the window moves to a different display. """
        current_screen = QGuiApplication.screenAt(self.ui.MainWindow.frameGeometry().center())
        screen_rect = current_screen.availableGeometry()
        available_height = screen_rect.height() * 0.90  # 90% of available screen height
        self.ui.MainWindow.w = 810
        self.ui.MainWindow.h = 938
        # MainWindow.resize(MainWindow.w, MainWindow.h)
        self.ui.MainWindow.setFixedSize(self.ui.MainWindow.w, self.ui.MainWindow.h)
        if current_screen:
            # print(f"Window moved to: {current_screen.name()}")
            self.ui.update_sizes()
        event.accept()

    def closeEvent(self, event):
        '''
                handling the exit from the program
                :param event:exit event
                :return:
                '''
        if self.c.status == "*":
            if self.c.fileName == "[Untitled]":
                exitMsg = "Are you sure you want to exit without saving?"
                rsp = QMessageBox.question(self, "Save Changes", exitMsg,
                                           QMessageBox.Yes | QMessageBox.Save | QMessageBox.Cancel)
                if rsp == QMessageBox.Yes:
                    event.accept()
                elif rsp == QMessageBox.Save:
                    self.saveAsTriggered()
                    event.accept()
                else:
                    event.ignore()
            else:
                # file_path = os.path.join(self.c.filePath, self.c.fileName)
                exitMsg = "Are you sure you want to exit without saving?"
                rsp = QMessageBox.question(self, "Save Changes", exitMsg,
                                           QMessageBox.Yes | QMessageBox.Save | QMessageBox.Cancel)
                if rsp == QMessageBox.Yes:
                    event.accept()
                elif rsp == QMessageBox.Save:
                    self.saveCloseTriggered()
                    event.accept()
                else:
                    event.ignore()


if __name__ == "__main__":
    configure_logging()
    app = QApplication(sys.argv)
    app.setAttribute(Qt.AA_EnableHighDpiScaling)
    app.setAttribute(Qt.AA_UseHighDpiPixmaps)
    window = Window()
    window.show()
    sys.exit(app.exec_())
