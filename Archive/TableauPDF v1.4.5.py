import sys
import logging
import configparser
import tempfile
import os
import pandas as pd
import requests
import webbrowser
import time
import json
import fitz # Import PyMuPDF for PDF operations
import traceback
from PIL import Image, ImageChops, ImageOps
from tableauserverclient import Server, PersonalAccessTokenAuth, PDFRequestOptions, ImageRequestOptions, RequestOptions, Filter
from threading import Thread, Event
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QPushButton, QCheckBox, QStyledItemDelegate, QStyle, QStyleOptionViewItem, QScrollArea,
    QLineEdit, QLabel, QComboBox, QFileDialog, QProgressBar, QTextEdit, QMessageBox, QToolButton, QMenu, QAction,
    QGroupBox, QDialog, QListWidget, QListWidgetItem, QFrame, QStyledItemDelegate, QRadioButton, QSizePolicy, QDialogButtonBox)
from PyQt5.QtCore import Qt, QEvent, pyqtSignal, pyqtSlot, QSize, Qt, QRect, QPoint, QModelIndex 
from PyQt5.QtGui import QFont, QPalette, QFontMetrics, QStandardItem, QTextCursor, QPixmap, QPainter, QPen, QColor, QIcon, QStandardItemModel

CURRENT_VERSION = "v1.4.5"
HIDDEN_TEMP_DIR = os.path.join(tempfile.gettempdir(), '.pdf_export_tool')
APPDATA_DIR = os.path.join(os.getenv('APPDATA'), 'PDFExportTool')
RECENT_FILES_PATH = os.path.join(APPDATA_DIR, 'recent_configs.json')

class CheckableComboBox(QComboBox):
    # Custom delegate to adjust item height
    class Delegate(QStyledItemDelegate):
        def sizeHint(self, option, index):
            size = super().sizeHint(option, index)
            size.setHeight(20) # Adjust item height for better spacing
            return size

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setModel(QStandardItemModel(self))
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        palette = self.lineEdit().palette()
        palette.setBrush(QPalette.Base, palette.button())
        self.lineEdit().setPalette(palette)
        self.setItemDelegate(CheckableComboBox.Delegate())
        self.model().dataChanged.connect(self.updateText)
        self.lineEdit().installEventFilter(self)
        self.view().viewport().installEventFilter(self)
        self.closeOnLineEditClick = False

    def setModel(self, model):
        super().setModel(model)
        if self.model():
            try:
                self.model().dataChanged.disconnect(self.updateText)
            except TypeError: pass
            self.model().dataChanged.connect(self.updateText)

    def clear(self):
        if self.model():
            self.model().clear()
        self.updateText()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.updateText()

    def eventFilter(self, obj, event):
        if obj == self.lineEdit():
            if event.type() == QEvent.MouseButtonRelease:
                if self.closeOnLineEditClick: self.hidePopup()
                else: self.showPopup()
                return True
        if obj == self.view().viewport():
            if event.type() == QEvent.MouseButtonRelease:
                index = self.view().indexAt(event.pos())
                if index.isValid():
                    item = self.model().itemFromIndex(index)
                    if item and item.flags() & Qt.ItemIsUserCheckable:
                        current_state = item.checkState()
                        item.setCheckState(Qt.Checked if current_state == Qt.Unchecked else Qt.Unchecked)
                        return True
        return False

    def showPopup(self):
        super().showPopup()
        self.closeOnLineEditClick = True

    def hidePopup(self):
        super().hidePopup()
        self.startTimer(50)
        self.closeOnLineEditClick = False

    def timerEvent(self, event):
        self.killTimer(event.timerId())
        self.closeOnLineEditClick = False

    def updateText(self):
        texts = []
        if self.model():
            for i in range(self.model().rowCount()):
                item = self.model().item(i)
                if item and item.checkState() == Qt.Checked: texts.append(item.text())
        if not texts: display_text = "None selected"
        elif len(texts) == 1: display_text = texts[0]
        elif self.model() and len(texts) == self.model().rowCount(): display_text = "All selected"
        else: display_text = f"{len(texts)} items selected"
        metrics = QFontMetrics(self.lineEdit().font())
        elidedText = metrics.elidedText(display_text, Qt.ElideRight, self.lineEdit().width() - 15)
        self.lineEdit().setText(elidedText)

    def addItem(self, text, data=None, checked=False):
        if not self.model(): return
        item = QStandardItem()
        item.setText(text)
        item.setData(data if data is not None else text)
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
        item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
        self.model().appendRow(item)

    def addItems(self, texts, datalist=None):
        if not self.model(): return
        self.model().blockSignals(True)
        for i, text in enumerate(texts):
            data = None
            if datalist:
                try: data = datalist[i]
                except (TypeError, IndexError): pass
            self.addItem(text, data)
        self.model().blockSignals(False)
        self.model().layoutChanged.emit()
        self.updateText()

    def currentData(self):
        res = []
        if self.model():
            for i in range(self.model().rowCount()):
                item = self.model().item(i)
                if item and item.checkState() == Qt.Checked: res.append(item.data())
        return res

    def getCheckedItemsText(self):
        texts = []
        if self.model():
            for i in range(self.model().rowCount()):
                item = self.model().item(i)
                if item and item.checkState() == Qt.Checked: texts.append(item.text())
        return texts

    def setCheckedByData(self, data_items_to_check):
        if not self.model(): return
        data_set = set(str(d) for d in data_items_to_check)
        self.model().blockSignals(True)
        for i in range(self.model().rowCount()):
            item = self.model().item(i)
            if item:
                 if str(item.data()) in data_set: item.setCheckState(Qt.Checked)
                 else: item.setCheckState(Qt.Unchecked)
        self.model().blockSignals(False)
        self.model().layoutChanged.emit()
        self.updateText()
# --- End of CheckableComboBox ---

class RedXCheckDelegate(QStyledItemDelegate):
    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex):
        # Get original option to modify
        opts = QStyleOptionViewItem(option)
        self.initStyleOption(opts, index)

        style = opts.widget.style() if opts.widget else QApplication.style()

        # --- Determine State ---
        is_checked = index.data(Qt.CheckStateRole) == Qt.Checked
        is_enabled = opts.state & QStyle.State_Enabled

        # --- Draw Background and Text ---
        # Let the default implementation handle selection highlighting etc.
        # But adjust text color if disabled
        original_text_role = opts.palette.currentColorGroup()
        if not is_enabled:
            opts.palette.setCurrentColorGroup(QPalette.Disabled) # Use disabled palette

        # Draw standard background, focus rect etc. (excluding check indicator and text for now)
        style.drawPrimitive(QStyle.PE_PanelItemViewItem, opts, painter, opts.widget)

        # Draw text (using adjusted palette if disabled)
        text_rect = style.subElementRect(QStyle.SE_ItemViewItemText, opts, opts.widget)
        # Adjust text rect slightly away from where the checkbox will be
        text_rect.setLeft(text_rect.left() + 20) # Add space for indicator
        style.drawItemText(painter, text_rect, opts.displayAlignment, opts.palette, is_enabled, opts.text, QPalette.Text)

        # Restore original palette color group
        opts.palette.setCurrentColorGroup(original_text_role)

        # --- Draw Custom Check Indicator ---
        check_rect = style.subElementRect(QStyle.SE_ItemViewItemCheckIndicator, opts, opts.widget)
        # Center the drawing area slightly
        indicator_size = 14
        draw_rect = QRect(
            check_rect.x() + (check_rect.width() - indicator_size) // 2,
            check_rect.y() + (check_rect.height() - indicator_size) // 2,
            indicator_size,
            indicator_size
        )

        painter.save() # Save painter state

        if is_checked:
            # Draw Red X
            pen_width = 2
            pen_color = QColor(Qt.red)
            if not is_enabled:
                # Slightly muted red if disabled but checked (globally excluded)
                 pen_color = QColor("#FF8080") # Lighter red

            pen = QPen(pen_color, pen_width)
            pen.setCapStyle(Qt.RoundCap) # Nicer line ends
            painter.setPen(pen)
            painter.setRenderHint(QPainter.Antialiasing, True)

            # Draw the 'X' lines within the draw_rect
            margin = 3 # Margin inside the indicator rect
            painter.drawLine(draw_rect.topLeft() + QPoint(margin, margin),
                             draw_rect.bottomRight() - QPoint(margin, margin))
            painter.drawLine(draw_rect.topRight() + QPoint(-margin, margin),
                             draw_rect.bottomLeft() - QPoint(margin, -margin))
        else:
             # Draw Empty Box (optional: could let default draw it, but this ensures consistency)
             pen_width = 1
             pen_color = QColor("#a0a0a0") # Default border color from QSS
             if not is_enabled:
                 pen_color = QColor("#dcdcdc") # Disabled border color

             painter.setPen(QPen(pen_color, pen_width))
             painter.setBrush(Qt.NoBrush) # No fill
             # Slightly smaller rect for the border
             border_rect = QRect(draw_rect.x()+1, draw_rect.y()+1, draw_rect.width()-2, draw_rect.height()-2)
             painter.drawRoundedRect(border_rect, 2, 2) # Draw rounded border


        painter.restore() # Restore painter state

    # Override sizeHint to provide adequate space
    def sizeHint(self, option: QStyleOptionViewItem, index: QModelIndex):
        size = super().sizeHint(option, index)
        # Add some padding to ensure space for checkbox + text
        size.setHeight(max(size.height(), 22)) # Ensure min height
        return size

# --- Server Configuration Dialog (No changes) ---
class ServerConfigDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_app = parent # Store reference to main application window
        self.setWindowTitle("Server Configuration")
        self.setMinimumWidth(500) # Set a reasonable minimum width

        # --- Layout ---
        layout = QVBoxLayout(self)
        grid_layout = QGridLayout()
        grid_layout.setSpacing(10)

        # --- Widgets ---
        self.server_url_edit = QLineEdit()
        self.token_name_edit = QLineEdit()
        self.token_secret_edit = QLineEdit()
        self.token_secret_edit.setEchoMode(QLineEdit.Password)
        self.site_id_edit = QLineEdit()

        # --- Arrange Widgets ---
        grid_layout.addWidget(QLabel("Server URL:"), 0, 0)
        grid_layout.addWidget(self.server_url_edit, 0, 1)
        grid_layout.addWidget(QLabel("Site ID:"), 1, 0)
        grid_layout.addWidget(self.site_id_edit, 1, 1)
        grid_layout.addWidget(QLabel("Token Name:"), 2, 0)
        grid_layout.addWidget(self.token_name_edit, 2, 1)
        grid_layout.addWidget(QLabel("Token Secret:"), 3, 0)
        grid_layout.addWidget(self.token_secret_edit, 3, 1)

        layout.addLayout(grid_layout)

        # --- Dialog Buttons ---
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept) # Connect Ok to accept()
        button_box.rejected.connect(self.reject) # Connect Cancel to reject()
        layout.addWidget(button_box)

        self.load_settings() # Load current settings when dialog opens

    def load_settings(self):
        """Loads current server settings from the parent application."""
        if self.parent_app:
            self.server_url_edit.setText(getattr(self.parent_app, 'server_url_text', ''))
            self.token_name_edit.setText(getattr(self.parent_app, 'token_name_text', ''))
            self.token_secret_edit.setText(getattr(self.parent_app, 'token_secret_text', ''))
            self.site_id_edit.setText(getattr(self.parent_app, 'site_id_text', ''))

    def accept(self):
        """Saves settings back to the parent application when Ok is clicked."""
        if self.parent_app:
            self.parent_app.server_url_text = self.server_url_edit.text().strip()
            self.parent_app.token_name_text = self.token_name_edit.text().strip()
            self.parent_app.token_secret_text = self.token_secret_edit.text() # Don't strip secret
            self.parent_app.site_id_text = self.site_id_edit.text().strip()
            self.parent_app.logger.info("Server configuration updated via dialog.")
            # Optionally update a status label on the main window if needed
            # self.parent_app.update_server_status_display()
        super().accept() # Close the dialog
# --- End Server Configuration Dialog ---

class PDFExportApp(QWidget):
    log_message_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)

    def __init__(self):
        super().__init__()
        print("Initializing PDFExportApp...")

        # --- Set Window Icon ---
        try:
            # Define the icon filename you want to use
            icon_filename = 'icon.png'
            # Use self.current_dir() to get the correct base path
            icon_path = os.path.join(self.current_dir(), 'icons', icon_filename)

            if os.path.exists(icon_path):
                app_icon = QIcon(icon_path)
                self.setWindowIcon(app_icon)
                print(f"Window icon set using: {icon_path}") # Log success
            else:
                # Log a warning if the specific icon file isn't found
                print(f"Warning: Window icon file '{icon_filename}' not found at: {icon_path}")
                # Optionally, you could try loading icon.svg here as a fallback if desired
                # icon_svg_path = os.path.join(self.current_dir(), 'icons', 'icon.svg')
                # if os.path.exists(icon_svg_path):
                #    app_icon = QIcon(icon_svg_path)
                #    self.setWindowIcon(app_icon)
                #    print(f"Window icon set using fallback: {icon_svg_path}")

        except Exception as e:
            # Log any unexpected error during icon setting
            print(f"Error setting window icon: {e}")
        # --- End Set Window Icon ---
        
        self.conditions = []
        self.parameters = []
        self.filters = [] # For Excel list filters (now QLineEdit + QListWidget)
        self.excel_path = ""
        self.tableau_views = [] # Stores names of views loaded from Tableau
        self.stop_event = Event() # Used to signal the worker thread to stop
        self.excluded_views_for_export = [] # Views globally excluded from export
        self.is_custom_theme_active = False # Start with default theme initially
        self.export_settings_section_visible = True
        self.logic_section_visible = True
        self.qss_file_path = os.path.join(self.current_dir(), "styles", "macos_style.qss")
        self.trim_pdf_enabled = False # To store the state of the new checkbox
        self.merge_pdfs_enabled = False # New: To store the state of the merge checkbox

        # Attributes to store server config (updated by dialog)
        self.server_url_text = ""
        self.token_name_text = ""
        self.token_secret_text = ""
        self.site_id_text = ""


        # Set window properties
        self.setWindowTitle(f'Tableau PDF/PNG Export Tool {CURRENT_VERSION}') # More descriptive title
        self.setGeometry(100, 100, 950, 800) # Adjusted size for potentially larger header/content

        # Initialize logging and UI components
        self.init_logger()
        self.initUI() # Call the revamped UI initializer

        # Connect signals for logging and progress updates
        self.log_message_signal.connect(self.update_log)
        self.progress_signal.connect(self.update_progress)

        # Check for application updates on startup
        self.check_for_updates()
        self._update_theme_button_text()

        # Style and Theme Management:

    def apply_custom_theme(self):
        """Loads and applies the custom QSS file."""
        app = QApplication.instance() # Get the application instance
        if not app:
            self.logger.error("Could not get QApplication instance to apply theme.")
            return False

        try:
            with open(self.qss_file_path, "r", encoding="utf-8") as f:
                style = f.read()
                app.setStyleSheet(style)
            self.logger.info(f"Applied custom theme from {self.qss_file_path}")
            self.is_custom_theme_active = True
            return True
        except FileNotFoundError:
            self.logger.error(f"Custom stylesheet file not found: {self.qss_file_path}")
            QMessageBox.warning(self, "Theme Error", f"Stylesheet file not found:\n{self.qss_file_path}")
            return False
        except Exception as e:
            self.logger.error(f"Error loading custom stylesheet: {e}", exc_info=True)
            QMessageBox.critical(self, "Theme Error", f"Could not load or apply stylesheet:\n{e}")
            # Optionally revert to default theme on error
            # self.apply_default_theme()
            return False

    def apply_default_theme(self):
        """Applies the base style by clearing the custom stylesheet."""
        app = QApplication.instance()
        if not app:
            self.logger.error("Could not get QApplication instance to clear theme.")
            return

        # Set empty stylesheet to revert to the base style set on app creation
        app.setStyleSheet("")

        self.logger.info("Reverted to default theme.")
        self.is_custom_theme_active = False

    def toggle_theme(self):
        """Switches between custom and default themes."""
        self.logger.debug(f"Toggling theme. Current custom active: {self.is_custom_theme_active}")
        if self.is_custom_theme_active:
            self.apply_default_theme()
        else:
            # Only set state to True if applying succeeds
            if not self.apply_custom_theme():
                # If applying custom fails, ensure we stay in default state
                self.apply_default_theme()

        self._update_theme_button_text()

    def _update_theme_button_text(self):
        """Updates the toggle button text based on the current theme state."""
        if hasattr(self, 'toggle_theme_btn'): # Check if button exists yet
            if self.is_custom_theme_active:
                self.toggle_theme_btn.setText("🎨")
                self.toggle_theme_btn.setToolTip("Switch back to the default style")
            else:
                self.toggle_theme_btn.setText("🎨")
                self.toggle_theme_btn.setToolTip("Switch to the custom macOS-like style")

    # --- HELPER: Safely remove widgets from a layout ---
    def _clear_layout_items(self, layout):
        """Recursively removes widgets and sub-layouts from a given layout."""
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    # Disconnect signals explicitly if needed, though deleteLater should handle most
                    widget.setParent(None)
                    widget.deleteLater()
                else:
                    sub_layout = item.layout()
                    if sub_layout is not None:
                        self._clear_layout_items(sub_layout) # Recursive call


    def init_logger(self):
        """Initializes the logging configuration."""
        print("Initializing logger...")
        self.logger = logging.getLogger('PDFExportApp')
        self.logger.setLevel(logging.DEBUG) # Set base level to DEBUG

        # Prevent adding multiple handlers if called again
        if not self.logger.handlers:
            # File handler - logs everything (DEBUG level)
            log_file_name = 'app.log'
            log_file_path = os.path.join(self.current_dir(), log_file_name) # Use current_dir()

            try:
                # Use rotating file handler for larger logs if needed
                # from logging.handlers import RotatingFileHandler
                # fh = RotatingFileHandler(log_file_path, maxBytes=5*1024*1024, backupCount=2, mode='w', encoding='utf-8')

                # *** FIX: Specify UTF-8 encoding for the file handler ***
                fh = logging.FileHandler(log_file_path, mode='w', encoding='utf-8') # Overwrite log on each start
                fh.setLevel(logging.DEBUG)
                # Use a formatter that handles Unicode characters well
                file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(funcName)s:%(lineno)d] - %(message)s')
                fh.setFormatter(file_formatter)
                self.logger.addHandler(fh)
                # Log the actual path being used
                print(f"File logger initialized. Log file at: {log_file_path}")
                self.logger.info(f"File logging directed to: {log_file_path}")
            except Exception as e:
                # Log error with the intended path
                print(f"Error setting up file logger ({log_file_path}): {e}")
                # Optionally, disable file logging or try an alternative path if this fails


            # Console handler - logs INFO and above
            # The console handler usually handles Unicode better automatically,
            # but specifying encoding might be needed in some environments.
            ch = logging.StreamHandler(sys.stdout) # Explicitly use stdout
            ch.setLevel(logging.INFO) # Log less verbose messages to console
            # Try setting encoding for console if issues persist, though often not needed
            # ch.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
            # try:
            #     sys.stdout.reconfigure(encoding='utf-8') # Attempt to reconfigure stdout
            # except AttributeError: # In case reconfigure is not available (e.g., older Python)
            #     pass
            console_formatter = logging.Formatter('%(levelname)s: %(message)s')
            ch.setFormatter(console_formatter)
            self.logger.addHandler(ch)

        self.logger.info(f"PDF Export Tool {CURRENT_VERSION} started.")


    def browse_output_folder(self):
        """Opens a dialog to select the output folder."""
        print("Browsing output folder...")
        current_output_dir = self.output_folder.text() or self.current_dir()
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder", current_output_dir)
        if folder:
            self.output_folder.setText(folder)
            self.logger.info(f"Selected output folder: {folder}")
        else:
            self.logger.info("Output folder selection cancelled.")


    def current_dir(self):
         """Returns the directory of the currently running script or executable."""
         if getattr(sys, 'frozen', False):
             # If running as a bundled executable (e.g., PyInstaller)
             return os.path.dirname(sys.executable)
         else:
             # If running as a script
             try:
                 # __file__ might not be defined in some environments (e.g., interactive)
                 return os.path.dirname(os.path.abspath(__file__))
             except NameError:
                 return os.getcwd() # Fallback to current working directory
             
    def show_help_info(self):
        """Placeholder function for showing help/info."""
        self.logger.info("Help/Info button clicked.")
        QMessageBox.information(self, "About PDF Export Tool",
                                f"Tableau PDF/PNG Export Tool\n"
                                f"Version: {CURRENT_VERSION}\n\n"
                                "This tool automates the process of exporting views from Tableau Server "
                                "based on criteria defined in an Excel file or exports selected views directly.\n\n"
                                "Developed by Haytham Soufi.\n"
                                "Find more details on GitHub (click the GitHub icon).")

    def initUI(self):
        """Initializes the main UI with scrolling, height limits, and fixed width."""
        print("Initializing UI (Scrollable Content, Height Limit, Fixed Width)...")
        self.logger.debug("Initializing UI with QScrollArea and fixed width.")

        # --- Main Layout (Header + ScrollArea) ---
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(0) # No space between header and scroll area
        main_layout.setContentsMargins(0, 0, 0, 0) # No margins around the main layout

        # --- Polished Header Section (Keep as before) ---
        self.header_widget = QWidget() # Store as instance variable
        self.header_widget.setObjectName("polishedHeaderWidget")
        self.header_height = 85 # Store header height
        self.header_widget.setFixedHeight(self.header_height) # Fixed height for the header
        header_main_layout = QHBoxLayout(self.header_widget)
        header_main_layout.setContentsMargins(20, 0, 15, 0) # L/R padding, no T/B padding
        header_main_layout.setSpacing(15)

        # == Left Section: Icon + Text Block ==
        left_section_widget = QWidget()
        left_section_layout = QHBoxLayout(left_section_widget)
        left_section_layout.setContentsMargins(0, 0, 0, 0)
        left_section_layout.setSpacing(12)

        icon_label = QLabel()
        icon_size = 80
        try:
            # Try SVG first, then PNG
            icon_path_svg = os.path.join(self.current_dir(), 'icons', 'icon.svg')
            icon_path_png = os.path.join(self.current_dir(), 'icons', 'icon.png')
            icon_path = ""
            if os.path.exists(icon_path_svg):
                icon_path = icon_path_svg
            elif os.path.exists(icon_path_png):
                icon_path = icon_path_png

            if icon_path:
                pixmap = QPixmap(icon_path)
                icon_label.setPixmap(pixmap.scaled(icon_size, icon_size, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            else:
                self.logger.warning("Header icon file (icon.svg or icon.png) not found.")
                icon_label.setText("?") # Fallback if no icon
                icon_label.setStyleSheet("color: white; font-size: 24pt; font-weight: bold;")
            icon_label.setFixedSize(icon_size, icon_size)
        except Exception as e:
            self.logger.error(f"Error loading header icon: {e}")
            icon_label.setText("!") # Error fallback
            icon_label.setStyleSheet("color: red; font-size: 24pt; font-weight: bold;")
            icon_label.setFixedSize(icon_size, icon_size)
        left_section_layout.addWidget(icon_label, alignment=Qt.AlignVCenter) # Vertically center icon

        # Text Block (Title, Version, Tagline)
        text_block_layout = QVBoxLayout()
        text_block_layout.setSpacing(1) # Very tight spacing
        text_block_layout.setContentsMargins(0, 0, 0, 0)
        title_label = QLabel("Tableau Bulk Exporter")
        title_label.setObjectName("headerTitleLabel")
        title_label.setFont(QFont('Segoe UI', 16, QFont.Medium)) # Medium weight, slightly smaller
        version_label = QLabel(f"{CURRENT_VERSION}") # Version only, less prominent
        version_label.setObjectName("headerVersionLabel")
        version_label.setFont(QFont('Segoe UI', 9))
        tagline_label = QLabel("Automated PDF & PNG Exports from Tabluea API") # Shorter tagline
        tagline_label.setObjectName("headerTaglineLabel")
        tagline_label.setFont(QFont('Segoe UI', 10))
        text_block_layout.addStretch(1) # Push text down towards vertical center
        text_block_layout.addWidget(title_label)
        text_block_layout.addWidget(version_label)
        text_block_layout.addWidget(tagline_label)
        text_block_layout.addStretch(1) # Push text up towards vertical center
        left_section_layout.addLayout(text_block_layout)
        header_main_layout.addWidget(left_section_widget) # Add left section to main header

        # == Spacer ==
        header_main_layout.addStretch(1) # Push buttons to the right

        # == Right Section: Action Buttons ==
        right_buttons_widget = QWidget()
        right_buttons_layout = QHBoxLayout(right_buttons_widget)
        right_buttons_layout.setContentsMargins(0, 0, 0, 0)
        right_buttons_layout.setSpacing(5) # Tighter spacing for action buttons
        button_icon_size = QSize(50, 50) # Standard size for action icons
        tooltip_style = "QToolTip { color: #ffffff; background-color: #333333; border: 1px solid #444444; }"

        # Help Button
        help_button = QToolButton()
        help_button.setObjectName("headerActionButton")
        help_button.setToolTip("Help / About")
        help_button.setStyleSheet(tooltip_style)
        help_button.setAutoRaise(True)
        help_button.setCursor(Qt.PointingHandCursor)
        help_button.setIconSize(button_icon_size)
        try:
            help_icon_path_svg = os.path.join(self.current_dir(), 'icons', 'help_icon.svg') # Assume white/light SVG
            help_icon_path_png = os.path.join(self.current_dir(), 'icons', 'help_icon.png')
            help_icon_path = ""
            if os.path.exists(help_icon_path_svg): help_icon_path = help_icon_path_svg
            elif os.path.exists(help_icon_path_png): help_icon_path = help_icon_path_png

            if help_icon_path: help_button.setIcon(QIcon(help_icon_path))
            else: self.logger.warning("Help icon not found."); help_button.setText("?")
        except Exception: help_button.setText("?") # Error fallback
        help_button.clicked.connect(self.show_help_info)
        right_buttons_layout.addWidget(help_button)

        # GitHub Button
        github_button = QToolButton()
        github_button.setObjectName("headerActionButton")
        github_button.setToolTip("View Project on GitHub")
        github_button.setStyleSheet(tooltip_style)
        github_button.setAutoRaise(True)
        github_button.setCursor(Qt.PointingHandCursor)
        github_button.setIconSize(button_icon_size)
        try:
            github_icon_path_svg = os.path.join(self.current_dir(), 'icons', 'github_icon.svg') # Assume white/light SVG
            github_icon_path_png = os.path.join(self.current_dir(), 'icons', 'github_icon.png')
            github_icon_path = ""
            if os.path.exists(github_icon_path_svg): github_icon_path = github_icon_path_svg
            elif os.path.exists(github_icon_path_png): github_icon_path = github_icon_path_png

            if github_icon_path: github_button.setIcon(QIcon(github_icon_path))
            else: self.logger.warning("GitHub icon not found."); github_button.setText("GH")
        except Exception: github_button.setText("GH") # Error fallback
        github_button.clicked.connect(lambda: webbrowser.open("https://github.com/haythamsoufi/TableauPDF"))
        right_buttons_layout.addWidget(github_button)

        header_main_layout.addWidget(right_buttons_widget, alignment=Qt.AlignVCenter) # Vertically center buttons

        # Apply Stylesheet to Header
        header_base_color = "#2c3e50"; header_text_color = "#ecf0f1"; header_hover_color = "#34495e"
        self.header_widget.setStyleSheet(f"""
            QWidget#polishedHeaderWidget {{ background-color: {header_base_color}; }}
            QWidget#polishedHeaderWidget QLabel {{ color: {header_text_color}; background-color: transparent; }}
            QLabel#headerVersionLabel {{ color: #bdc3c7; }}
            QLabel#headerTaglineLabel {{ color: #bdc3c7; }}
            QToolButton#headerActionButton {{ background-color: transparent; border: none; padding: 5px; border-radius: 4px; color: {header_text_color}; font-weight: bold; }}
            QToolButton#headerActionButton:hover {{ background-color: {header_hover_color}; }}
            QToolButton#headerActionButton:pressed {{ background-color: rgba(0, 0, 0, 0.2); }}
            {tooltip_style}
        """)

        main_layout.addWidget(self.header_widget) # Add the finished header
        # --- End of Polished Header Section ---

        # --- Scroll Area for Content ---
        self.scroll_area = QScrollArea() # Store scroll area reference
        self.scroll_area.setWidgetResizable(True) # Allows the inner widget to resize
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded) # Show scrollbar only when needed
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff) # Disable horizontal scrollbar
        self.scroll_area.setFrameShape(QFrame.NoFrame) # Remove scroll area border

        # --- Content Area Wrapper (Goes INSIDE ScrollArea) ---
        # This widget will contain all the sections below the header
        self.content_widget = QWidget() # Store content widget reference
        # Set an object name for potential styling of the background within the scroll area
        self.content_widget.setObjectName("scrollableContentWidget")
        # The main vertical layout for the content area
        content_layout = QVBoxLayout(self.content_widget)
        content_layout.setSpacing(15)
        content_layout.setContentsMargins(15, 15, 15, 15) # Padding for content area

        # --- Configuration Section ---
        # The setup_export_settings_panel now creates its own container widget
        # and adds it directly to the content_layout passed to it.
        self.setup_export_settings_panel(content_layout) # Pass content_layout

        # --- Combined Logic Section ---
        # The setup_combined_logic_panel also adds its container to content_layout.
        self.setup_combined_logic_panel(content_layout) # Pass content_layout

        # --- Progress and Control Section ---
        # These setup functions add their group boxes directly to content_layout.
        self.setup_progress_panel(content_layout)
        self.setup_control_buttons(content_layout)

        # --- Set Content Widget into Scroll Area ---
        # The content_widget now holds all the sections (export, logic, progress, controls)
        self.scroll_area.setWidget(self.content_widget)

        # --- Add Scroll Area to Main Layout ---
        main_layout.addWidget(self.scroll_area) # Add scroll area below header

        # --- Set Maximum Window Height ---
        self.max_height = 800 # Default max height if screen info fails
        try:
            screen = QApplication.primaryScreen()
            if screen:
                available_geometry = screen.availableGeometry() # Geometry excluding taskbar etc.
                # Set max height slightly less than available height to avoid overlap
                self.max_height = available_geometry.height() - 50 # Subtract a margin (adjust as needed)
                self.setMaximumHeight(self.max_height)
                self.logger.info(f"Set maximum window height to: {self.max_height}px")
            else:
                 self.logger.warning("Could not get primary screen information. Using default max height.")
                 self.setMaximumHeight(self.max_height) # Use default if screen info fails
        except Exception as e:
            self.logger.error(f"Could not determine screen geometry or set max height: {e}")
            self.setMaximumHeight(self.max_height) # Use default on error

        # --- Set Fixed Width ---
        # Use the width from the original setGeometry call or define explicitly
        self.initial_width = 1000 # Store initial width
        self.setFixedWidth(self.initial_width)
        self.logger.info(f"Set fixed window width to: {self.initial_width}px")

        # --- Set Initial Height ---
        self._adjust_window_height() # Call helper to set initial height


        # --- Final Connections and Initial State ---
        if hasattr(self, 'mode_selection_dropdown'):
            self.mode_selection_dropdown.currentIndexChanged.connect(self.toggle_mode_selection)
            self.toggle_mode_selection() # Set initial state based on mode
        else:
            self.logger.error("mode_selection_dropdown not found after setup_export_settings_panel.")

    def _adjust_window_height(self):
        """Calculates required height and resizes window, respecting max height."""
        QApplication.processEvents() # Ensure layout is updated

        # Calculate the required height based on the content widget's size hint
        content_height = self.content_widget.sizeHint().height()

        # Add height of the header
        total_required_height = content_height + self.header_height

        # Add margins/spacing from the main layout (usually 0) and scroll area (usually 0 for NoFrame)
        main_margins = self.layout().contentsMargins()
        total_required_height += main_margins.top() + main_margins.bottom()
        # Add a small buffer for safety/aesthetics
        total_required_height += 10

        # Clamp the height between a minimum (optional) and the maximum screen height
        # min_height = 400 # Example minimum height
        # target_height = max(min_height, total_required_height)
        target_height = min(total_required_height, self.max_height)

        self.logger.debug(f"Adjusting height: ContentHint={content_height}, Required={total_required_height}, Target={target_height}")
        self.resize(self.initial_width, target_height) # Resize with fixed width

    def reset_fields(self):
        """Resets most input fields and internal state to default."""
        print("Resetting fields...")
        self.logger.info("Resetting input fields.")
        # Server Config Attributes
        self.server_url_text = ""
        self.token_name_text = ""
        self.token_secret_text = ""
        self.site_id_text = ""
        # Export Settings
        self.workbook_name.clear()
        self.excel_file.clear()
        self.sheet_name.clear()
        self.sheet_name.setEnabled(False)
        self.tableau_filter_field_dropdown.clear()
        self.tableau_filter_field_dropdown.setEnabled(False)
        self.output_folder.clear()
        self.excel_path = ""
        self.file_naming_option.clear()
        self.file_naming_option.addItem("By view")
        self.organize_by_dropdown.clear()
        self.organize_by_dropdown.addItem("None")
        self.organize_by2_dropdown.clear()
        self.organize_by2_dropdown.addItem("None")
        self.pdf_radio.setChecked(True)
        self.numbering_checkbox.setChecked(True)
        self.trim_pdf_checkbox.setChecked(False) # Reset trim checkbox
        self.merge_pdfs_checkbox.setChecked(False) # Reset merge checkbox
        self.mode_selection_dropdown.setCurrentIndex(0) # Default to "Automate for a list"
        # Reset combined logic section
        self.reset_combined_logic()
        # Reset internal state
        self.tableau_views = []
        self.excluded_views_for_export = []
        # Clear log and progress
        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0%")


    def reset_server_configuration(self):
        """Resets only the server configuration attributes."""
        print("Resetting server configuration attributes...")
        self.server_url_text = ""
        self.token_name_text = ""
        self.token_secret_text = ""
        self.site_id_text = ""
        self.workbook_name.clear() # Workbook name is logically tied to server config
        self.tableau_views = [] # Clear loaded views if server config resets
        self.logger.info("Server configuration attributes reset.")


    def setup_export_settings_panel(self, main_layout_for_section):
        """Sets up the collapsible panel for export-related settings with the updated layout."""
        print("Setting up export settings panel (collapsible)...")
        self.logger.debug("Setting up collapsible export settings panel.")

        # --- Main Container for this whole section ---
        self.export_settings_container_widget = QWidget()
        self.export_settings_container_widget.setObjectName("exportSettingsContainerWidget")
        container_layout = QVBoxLayout(self.export_settings_container_widget)
        container_layout.setContentsMargins(0, 5, 0, 5)
        container_layout.setSpacing(0)

        # --- Custom "Title Bar" ---
        title_bar_widget = QWidget()
        title_bar_layout = QHBoxLayout(title_bar_widget)
        title_bar_layout.setContentsMargins(5, 3, 5, 3)
        title_bar_layout.setSpacing(8)

        # Title Label
        title_label = QLabel("Export Settings")
        title_label.setObjectName("exportSettingsTitleLabel")
        title_label.setStyleSheet("font-weight: bold; color: #444444;")

        # Toggle Button
        self.export_settings_toggle_button = QToolButton()
        self.export_settings_toggle_button.setObjectName("exportSettingsToggleButton")
        self.export_settings_toggle_button.setToolTip("Show/Hide Export Settings")
        self.export_settings_toggle_button.setCursor(Qt.PointingHandCursor)
        self.export_settings_toggle_button.setCheckable(True)
        self.export_settings_toggle_button.setChecked(self.export_settings_section_visible)
        self.export_settings_toggle_button.setStyleSheet("""
            QToolButton#exportSettingsToggleButton {
                border: none;
                background-color: transparent;
                padding: 2px;
            }
            QToolButton#exportSettingsToggleButton:hover {
                background-color: rgba(0, 0, 0, 0.1);
                border-radius: 3px; }
        """)
        self.export_settings_toggle_button.clicked.connect(self.toggle_export_settings_section)

        # Add elements to title bar layout
        title_bar_layout.addWidget(self.export_settings_toggle_button)
        title_bar_layout.addWidget(title_label)
        title_bar_layout.addStretch(1)

        # Add title bar to the main container layout
        container_layout.addWidget(title_bar_widget)

        # --- Content Widget (Holds the actual settings) ---
        self.exportSettingsContentWidget = QWidget()
        self.exportSettingsContentWidget.setObjectName("exportSettingsContentWidget")
        export_settings_layout = QGridLayout()
        export_settings_layout.setSpacing(8)
        self.exportSettingsContentWidget.setLayout(export_settings_layout)

        # --- Define Widgets ---
        # Row 0: Workbook, Test Connection, Mode
        self.workbook_name = QLineEdit()
        self.workbook_name.setPlaceholderText("Tableau Workbook Name")
        load_views_btn = QPushButton("🔍 Test")
        load_views_btn.setToolTip("Connect to Tableau, verify workbook, and load view names")
        self.mode_selection_dropdown = QComboBox()
        self.mode_selection_dropdown.addItems(["Automate for a list", "Export All Views Once"])
        self.mode_selection_dropdown.setToolTip("Choose export mode")

        # Row 1: Excel File, Browse, Sheet Name
        self.excel_file_label = QLabel("Excel File:")
        self.excel_file = QLineEdit()
        self.excel_file.setPlaceholderText("Path to .xlsx file")
        self.browse_excel = QPushButton("📂 Browse")
        self.sheet_name_label = QLabel("Sheet Name:")
        self.sheet_name = QComboBox()
        self.sheet_name.setEnabled(False)

        # Row 2: File Naming, Organize 1, Organize 2
        self.file_naming_label = QLabel("File Naming:")
        self.file_naming_option = QComboBox()
        self.file_naming_option.addItem("By view")
        self.file_naming_option.setToolTip("Choose how exported files are named (using view name or Excel column)")
        self.organize_by_label = QLabel("Organize by 1:")
        self.organize_by_dropdown = QComboBox()
        self.organize_by_dropdown.addItem("None")
        self.organize_by_dropdown.setToolTip("Create subfolders based on this Excel column (optional)")
        self.organize_by2_label = QLabel("Organize by 2:")
        self.organize_by2_dropdown = QComboBox()
        self.organize_by2_dropdown.addItem("None")
        self.organize_by2_dropdown.setToolTip("Create nested subfolders based on this Excel column (optional)")

        # Row 3: Key Field, Output Folder, Browse Output
        self.tableau_filter_field_label = QLabel("Key Field:")
        self.tableau_filter_field_dropdown = QComboBox()
        self.tableau_filter_field_dropdown.setEnabled(False)
        self.tableau_filter_field_dropdown.setToolTip("Excel column whose value will filter the Tableau view (optional)")
        self.output_folder_label = QLabel("Output Folder:")
        self.output_folder = QLineEdit()
        self.output_folder.setPlaceholderText("Path where files will be saved")
        browse_output_btn = QPushButton("📂 Browse")

        # Row 4: Format, Numbering, Trim, Merge, Buttons
        self.format_label = QLabel("Format:")
        self.pdf_radio = QRadioButton("PDF")
        self.png_radio = QRadioButton("PNG")
        self.pdf_radio.setChecked(True)
        self.numbering_checkbox = QCheckBox("Add numbering")
        self.numbering_checkbox.setChecked(True)
        self.numbering_checkbox.setToolTip("Prefix filenames with numbers (e.g., 01_, 02_)")
        
        # New: Trim PDF/PNG Checkbox
        self.trim_pdf_checkbox = QCheckBox("Trim Output")
        self.trim_pdf_checkbox.setToolTip(
            "Attempt to trim empty whitespace from PDF/PNG after export.\n"
            "Requires local libraries (e.g., PyMuPDF for PDF, Pillow for PNG)."
        )
        self.trim_pdf_checkbox.setChecked(self.trim_pdf_enabled) # Set initial state
        self.trim_pdf_checkbox.setEnabled(self.pdf_radio.isChecked()) # Only enabled for PDF initially

        # New: Merge PDFs Checkbox
        self.merge_pdfs_checkbox = QCheckBox("Merge PDFs")
        self.merge_pdfs_checkbox.setToolTip(
            "If 'Export All Views Once' mode, all PDFs are merged into one file.\n"
            "If 'Automate for a list' mode, PDFs are merged at the lowest 'Organize by' level."
        )
        self.merge_pdfs_checkbox.setChecked(self.merge_pdfs_enabled) # Set initial state
        self.merge_pdfs_checkbox.setEnabled(self.pdf_radio.isChecked()) # Only enabled for PDF

        reset_settings_btn = QPushButton("🔄 Reset")
        reset_settings_btn.setToolTip("Reset all fields in this section")
        select_views_btn = QPushButton("📑 Select Views")
        select_views_btn.setToolTip("Choose which Tableau views to include/exclude globally")


        # --- Arrange Widgets in Grid ---
        # Row 0 (Workbook, Test, Mode)
        export_settings_layout.addWidget(QLabel("Workbook Name:"), 0, 0)
        export_settings_layout.addWidget(self.workbook_name, 0, 1, 1, 2) # Span 2 columns
        export_settings_layout.addWidget(load_views_btn, 0, 3)
        export_settings_layout.addWidget(QLabel("Export Mode:"), 0, 4)
        export_settings_layout.addWidget(self.mode_selection_dropdown, 0, 5, 1, 3) # Span 3 columns

        # Row 1 (Excel, Browse, Sheet)
        export_settings_layout.addWidget(self.excel_file_label, 1, 0)
        export_settings_layout.addWidget(self.excel_file, 1, 1, 1, 3) # Span 3 columns
        export_settings_layout.addWidget(self.browse_excel, 1, 4)
        export_settings_layout.addWidget(self.sheet_name_label, 1, 5)
        export_settings_layout.addWidget(self.sheet_name, 1, 6, 1, 2) # Span 2 columns

        # Row 2 (File Naming, Org1, Org2)
        export_settings_layout.addWidget(self.file_naming_label, 2, 0)
        export_settings_layout.addWidget(self.file_naming_option, 2, 1)
        export_settings_layout.addWidget(self.organize_by_label, 2, 2)
        export_settings_layout.addWidget(self.organize_by_dropdown, 2, 3)
        export_settings_layout.addWidget(self.organize_by2_label, 2, 4)
        export_settings_layout.addWidget(self.organize_by2_dropdown, 2, 5, 1, 3) # Span 3 columns

        # Row 3 (Key Field, Output Folder, Browse Output)
        export_settings_layout.addWidget(self.tableau_filter_field_label, 3, 0)
        export_settings_layout.addWidget(self.tableau_filter_field_dropdown, 3, 1, 1, 2) # Span 2 columns
        export_settings_layout.addWidget(self.output_folder_label, 3, 3)
        export_settings_layout.addWidget(self.output_folder, 3, 4, 1, 3) # Span 3 columns
        export_settings_layout.addWidget(browse_output_btn, 3, 7)

        # Row 4 (Format, Numbering, Trim, Merge)
        export_settings_layout.addWidget(self.format_label, 4, 0)
        format_layout = QHBoxLayout()
        format_layout.addWidget(self.pdf_radio)
        format_layout.addWidget(self.png_radio)
        format_layout.addStretch()
        export_settings_layout.addLayout(format_layout, 4, 1)
        export_settings_layout.addWidget(self.numbering_checkbox, 4, 2)
        export_settings_layout.addWidget(self.trim_pdf_checkbox, 4, 3)
        export_settings_layout.addWidget(self.merge_pdfs_checkbox, 4, 4) # New Merge checkbox

        # Row 5 (Buttons)
        button_layout_row5 = QHBoxLayout()
        button_layout_row5.setSpacing(10)
        button_layout_row5.addStretch(1) # Push buttons right
        button_layout_row5.addWidget(select_views_btn)
        button_layout_row5.addWidget(reset_settings_btn)
        export_settings_layout.addLayout(button_layout_row5, 5, 0, 1, 8) # Span all 8 columns

        # --- Connect Signals ---
        load_views_btn.clicked.connect(self.load_tableau_views)
        self.browse_excel.clicked.connect(self.browse_excel_file)
        browse_output_btn.clicked.connect(self.browse_output_folder)
        reset_settings_btn.clicked.connect(self.reset_export_settings)
        select_views_btn.clicked.connect(self.select_views_for_export)

        # Connect radio buttons to update checkbox enabled states
        self.pdf_radio.toggled.connect(self.update_trim_merge_checkbox_states)
        self.png_radio.toggled.connect(self.update_trim_merge_checkbox_states)

        # Add the content widget to the main container layout
        container_layout.addWidget(self.exportSettingsContentWidget)
        self.exportSettingsContentWidget.setVisible(self.export_settings_section_visible)

        # Add the container widget to the main layout (QVBoxLayout)
        main_layout_for_section.addWidget(self.export_settings_container_widget)

        # Set the initial icon for the toggle button
        if hasattr(self, '_update_export_settings_toggle_icon'):
            self._update_export_settings_toggle_icon()
        else:
             self.logger.warning("_update_export_settings_toggle_icon method not found yet.")

    def update_trim_merge_checkbox_states(self):
        """Updates the enabled state of the trim and merge checkboxes based on format selection."""
        is_pdf_selected = self.pdf_radio.isChecked()
        self.trim_pdf_checkbox.setEnabled(True) # Trim is available for both PDF and PNG
        self.merge_pdfs_checkbox.setEnabled(is_pdf_selected) # Merge only for PDF

    def reset_export_settings(self):
        """Resets fields within the Export Settings group."""
        print("Resetting export settings...")
        self.logger.info("Resetting export settings fields.")
        # Don't reset server config here
        self.workbook_name.clear()
        self.excel_file.clear()
        self.sheet_name.clear()
        self.sheet_name.setEnabled(False)
        self.tableau_filter_field_dropdown.clear()
        self.tableau_filter_field_dropdown.setEnabled(False)
        self.output_folder.clear()
        self.excel_path = ""
        self.file_naming_option.clear()
        self.file_naming_option.addItem("By view")
        self.organize_by_dropdown.clear()
        self.organize_by_dropdown.addItem("None")
        self.organize_by2_dropdown.clear()
        self.organize_by2_dropdown.addItem("None")
        self.pdf_radio.setChecked(True)
        self.numbering_checkbox.setChecked(True)
        self.trim_pdf_checkbox.setChecked(False)
        self.merge_pdfs_checkbox.setChecked(False) # Reset merge checkbox
        self.mode_selection_dropdown.setCurrentIndex(0)
        # Keep excluded_views_for_export as is, user might want to keep the selection
        self.reset_combined_logic() # Reset the combined logic section
        self.update_trim_merge_checkbox_states() # Update enabled states after reset

    def select_views_for_export(self):
            print("Selecting views for global export exclusion...") # Log updated purpose
            if not self.tableau_views:
                # ... (keep existing logic to load views first if needed) ...
                try:
                    self.load_tableau_views()
                except Exception as e:
                    QMessageBox.warning(self, "Connection Error", f"Failed to connect and load views: {e}")
                    self.logger.error(f"Failed to load views for global selection: {e}")
                    return

                if not self.tableau_views: # Check again after loading attempt
                    QMessageBox.warning(self, "No Views Loaded", "No views available after attempting connection.")
                    return


            dialog = QDialog(self)
            dialog.setWindowTitle("Select Views to EXCLUDE Globally") # Title updated
            dialog.setMinimumWidth(400)
            layout = QVBoxLayout(dialog) # Set layout on dialog

            # *** UPDATED INSTRUCTION ***
            info_label = QLabel("Check views to EXCLUDE from all exports.")
            layout.addWidget(info_label)

            list_widget = QListWidget()
            list_widget.setSpacing(2)
            # *** APPLY CUSTOM DELEGATE ***
            list_widget.setItemDelegate(RedXCheckDelegate(list_widget))

            current_excluded_set = set(getattr(self, 'excluded_views_for_export', []))

            for view_name in self.tableau_views:
                item = QListWidgetItem(view_name)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                # *** REVERSED LOGIC: Check if it IS excluded ***
                item.setCheckState(Qt.Checked if view_name in current_excluded_set else Qt.Unchecked)
                list_widget.addItem(item)

            button_layout = QHBoxLayout()
            # *** UPDATE BUTTON LABELS/TOOLTIPS ***
            tick_all_btn = QPushButton("Exclude All")
            tick_all_btn.setToolTip("Check all views to exclude them.")
            untick_all_btn = QPushButton("Include All")
            untick_all_btn.setToolTip("Uncheck all views to include them.")
            # Keep connections the same, meaning changes now
            tick_all_btn.clicked.connect(lambda: self.set_all_check_states(list_widget, Qt.Checked))
            untick_all_btn.clicked.connect(lambda: self.set_all_check_states(list_widget, Qt.Unchecked))
            button_layout.addWidget(untick_all_btn) # Include All (Uncheck) first might be more logical
            button_layout.addWidget(tick_all_btn)

            button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            button_box.accepted.connect(lambda: self.apply_view_selection_for_export(dialog, list_widget))
            button_box.rejected.connect(dialog.reject)

            layout.addWidget(list_widget)
            layout.addLayout(button_layout)
            layout.addWidget(button_box)
            # Removed custom OK button, using standard DialogButtonBox

            dialog.exec_()

        # Keep set_all_check_states as it is

    def apply_view_selection_for_export(self, dialog, list_widget):
        """Applies the view selections from the dialog to the global exclusion list."""
        print("Applying global view exclusion selection...")
        newly_excluded = []
        included_count = 0
        for index in range(list_widget.count()):
            item = list_widget.item(index)
            if item:
                # *** REVERSED LOGIC: Checked means EXCLUDE ***
                if item.checkState() == Qt.Checked:
                    newly_excluded.append(item.text())
                else:
                    included_count += 1

        self.excluded_views_for_export = newly_excluded
        self.logger.info(f"Global view selection updated. Included: {included_count}, Excluded: {len(newly_excluded)}")
        self.log_message_signal.emit(f"Global view selection updated ({len(newly_excluded)} excluded).")
        dialog.accept()

    def set_all_check_states(self, list_widget, state):
        """Sets the check state of all items in a QListWidget."""
        for index in range(list_widget.count()):
            item = list_widget.item(index)
            if item: # Check if item exists
                item.setCheckState(state)


    def toggle_mode_selection(self):
        """Shows/hides UI elements based on the selected export mode, including the logic section container."""
        print("Toggling mode selection...")
        # Ensure mode_selection_dropdown exists before accessing it
        if not hasattr(self, 'mode_selection_dropdown'):
            self.logger.error("mode_selection_dropdown not found in toggle_mode_selection.")
            return

        mode = self.mode_selection_dropdown.currentText()
        self.logger.info(f"Export mode changed to: {mode}")

        is_automate_mode = (mode == "Automate for a list")

        # Toggle visibility of widgets specific to "Automate for a list" mode
        # These are typically inside the "Export Settings" group box
        # List the *attribute names* as strings for safer checking
        widget_attribute_names = [
            'excel_file_label', 'excel_file', 'browse_excel',
            'sheet_name_label', 'sheet_name',
            'tableau_filter_field_label', 'tableau_filter_field_dropdown',
            'file_naming_label', 'file_naming_option',
            'organize_by_label', 'organize_by_dropdown',
            'organize_by2_label', 'organize_by2_dropdown',
        ]

        for attr_name in widget_attribute_names:
            if hasattr(self, attr_name): # Check if the app object has this attribute
                widget = getattr(self, attr_name) # Get the widget object using the name
                if widget: # Check if the attribute holds a valid widget object (not None)
                    widget.setVisible(is_automate_mode)
                else:
                    # Log if the attribute exists but is None (less common issue)
                    self.logger.debug(f"Attribute '{attr_name}' exists but is None during mode toggle.")
            # else: # Optional: Log if an expected attribute name is missing entirely
                # self.logger.warning(f"Attribute '{attr_name}' not found on self during mode toggle.")


        # Toggle visibility of the ENTIRE logic container based *only* on the mode
        if hasattr(self, 'logic_container_widget'):
             self.logic_container_widget.setVisible(is_automate_mode)
             # The visibility of the content *within* the container is handled separately
             # by self.logic_section_visible and the toggle button.
        else:
            self.logger.warning("logic_container_widget not found during mode toggle.")


        if not is_automate_mode:
            # Reset fields specific to automate mode when switching away
            # Use hasattr checks for robustness before clearing/disabling
            if hasattr(self, 'excel_file'): self.excel_file.clear()
            if hasattr(self, 'sheet_name'):
                self.sheet_name.clear()
                self.sheet_name.setEnabled(False)
            if hasattr(self, 'tableau_filter_field_dropdown'):
                self.tableau_filter_field_dropdown.clear()
                self.tableau_filter_field_dropdown.setEnabled(False)
            if hasattr(self, 'file_naming_option'):
                self.file_naming_option.clear()
                self.file_naming_option.addItem("By view")
            if hasattr(self, 'organize_by_dropdown'):
                self.organize_by_dropdown.clear()
                self.organize_by_dropdown.addItem("None")
            if hasattr(self, 'organize_by2_dropdown'):
                self.organize_by2_dropdown.clear()
                self.organize_by2_dropdown.addItem("None")

            # Reset the actual filter/condition/parameter lines
            # Ensure reset_combined_logic method exists before calling
            if hasattr(self, 'reset_combined_logic'):
                self.reset_combined_logic()
            else:
                self.logger.error("reset_combined_logic method not found during mode toggle reset.")

    def browse_excel_file(self):
        """Opens a dialog to select the Excel file."""
        print("Browsing for Excel file...")
        current_excel_dir = os.path.dirname(self.excel_file.text()) if self.excel_file.text() else self.current_dir()
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", current_excel_dir, "Excel Files (*.xlsx)")
        if file_name:
            self.excel_file.setText(file_name)
            self.excel_path = file_name
            self.logger.info(f"Selected Excel file: {file_name}")
            self.load_sheets() # Load sheets immediately after selecting file
        else:
            self.logger.info("Excel file selection cancelled.")


    def load_sheets(self):
        """Loads sheet names from the selected Excel file into the sheet_name combobox."""
        print("Loading sheets...")
        if not self.excel_path:
            self.logger.warning("Load sheets called without an Excel file path.")
            return

        while True: # Loop to allow retry on PermissionError
            try:
                # Use context manager for file handling
                with pd.ExcelFile(self.excel_path) as xls:
                    sheet_names = xls.sheet_names
                    self.sheet_name.clear()
                    self.sheet_name.addItems(sheet_names)
                    self.sheet_name.setEnabled(True)
                    self.tableau_filter_field_dropdown.setEnabled(True) # Enable dependent dropdown

                    # Disconnect first to avoid multiple connections if called again
                    try:
                        self.sheet_name.currentIndexChanged.disconnect(self.on_sheet_selection_changed)
                    except TypeError: # No connection existed
                        pass
                    # Reconnect the signal
                    self.sheet_name.currentIndexChanged.connect(self.on_sheet_selection_changed)

                    # Trigger initial population for the first sheet if available
                    if sheet_names:
                        self.logger.info(f"Loaded sheets: {sheet_names}. Selecting first sheet '{sheet_names[0]}'.")
                        self.on_sheet_selection_changed() # Populate dropdowns based on the first sheet
                    else:
                         self.logger.warning(f"Excel file '{os.path.basename(self.excel_path)}' contains no sheets.")
                         self.on_sheet_selection_changed() # Call to clear dependent dropdowns

                    break # Exit while loop on success

            except PermissionError:
                self.logger.error(f"Permission denied for Excel file: {self.excel_path}. It might be open.")
                retry = QMessageBox.warning(
                    self, "File Locked",
                    f"The Excel file '{os.path.basename(self.excel_path)}' appears to be open or locked.\n\nPlease close the file and press Retry.",
                    QMessageBox.Retry | QMessageBox.Cancel, QMessageBox.Retry
                )
                if retry == QMessageBox.Cancel:
                    self.logger.info("User cancelled retry for locked Excel file.")
                    self.excel_file.clear() # Clear field if user cancels
                    self.excel_path = ""
                    self.sheet_name.clear() # Clear sheets dropdown
                    self.sheet_name.setEnabled(False)
                    self.on_sheet_selection_changed() # Clear dependent dropdowns
                    return # Exit method

            except FileNotFoundError:
                 self.logger.error(f"Excel file not found: {self.excel_path}")
                 QMessageBox.critical(self, "Error", f"Excel file not found:\n{self.excel_path}")
                 self.excel_file.clear()
                 self.excel_path = ""
                 self.sheet_name.clear()
                 self.sheet_name.setEnabled(False)
                 self.on_sheet_selection_changed() # Clear dependent dropdowns
                 break # Exit loop

            except Exception as e:
                self.logger.error(f"Could not load sheets from '{self.excel_path}': {e}", exc_info=True)
                QMessageBox.critical(self, "Error", f"Could not load sheets from the Excel file:\n{e}")
                self.excel_file.clear() # Clear field on error
                self.excel_path = ""
                self.sheet_name.clear()
                self.sheet_name.setEnabled(False)
                self.on_sheet_selection_changed() # Clear dependent dropdowns
                break # Exit while loop on other error


    def on_sheet_selection_changed(self):
        """Handles changes in the selected sheet, populating dependent dropdowns."""
        print("Sheet selection changed...")
        sheet = self.sheet_name.currentText()
        self.logger.info(f"Sheet selection changed to: '{sheet}'")

        # Clear dependent UI elements first
        self.file_naming_option.clear()
        self.file_naming_option.addItem("By view")
        self.tableau_filter_field_dropdown.clear()
        self.organize_by_dropdown.clear()
        self.organize_by_dropdown.addItem("None")
        self.organize_by2_dropdown.clear()
        self.organize_by2_dropdown.addItem("None")
        self.reset_combined_logic() # Reset filters, conditions, params

        if sheet and self.excel_path:
            try:
                # Read only header row first to get columns quickly
                df_header = pd.read_excel(self.excel_path, sheet_name=sheet, nrows=1)
                columns = list(df_header.columns)
                self.logger.info(f"Loaded columns from sheet '{sheet}': {columns}")

                # Populate File Naming options
                self.file_naming_option.addItems(columns)

                # Populate Tableau Filter Field dropdown
                self.populate_tableau_filter_field_dropdown(columns)

                # Populate Organize By dropdowns
                self.populate_organize_by_dropdown(columns)

                # Populate Field dropdowns in existing Filter and Condition lines
                # (These functions need to exist and handle the new structure if needed)
                self.load_columns_for_filters(columns)
                self.load_columns_for_conditions(columns)


            except Exception as e:
                self.logger.error(f"Error reading columns from sheet '{sheet}': {e}", exc_info=True)
                QMessageBox.critical(self, "Error", f"Error reading columns from sheet '{sheet}':\n{e}")
                # Ensure dropdowns remain cleared/disabled
                self.tableau_filter_field_dropdown.setEnabled(False)
                self.organize_by_dropdown.setEnabled(False)
                self.organize_by2_dropdown.setEnabled(False)
        else:
             # Disable dependent dropdowns if no sheet is selected
             self.logger.info("No sheet selected or Excel path missing, clearing dependent dropdowns.")
             self.tableau_filter_field_dropdown.setEnabled(False)
             self.organize_by_dropdown.setEnabled(False)
             self.organize_by2_dropdown.setEnabled(False)


    def populate_organize_by_dropdown(self, columns):
        """Populates the 'Organize by' dropdowns with Excel columns."""
        print("Populating organize by dropdown...")
        try:
            options = ['None'] + columns # Add 'None' option
            # Dropdown 1
            self.organize_by_dropdown.clear()
            self.organize_by_dropdown.addItems(options)
            self.organize_by_dropdown.setCurrentText('None') # Default to None
            self.organize_by_dropdown.setEnabled(True)
            # Dropdown 2
            self.organize_by2_dropdown.clear()
            self.organize_by2_dropdown.addItems(options)
            self.organize_by2_dropdown.setCurrentText('None') # Default to None
            self.organize_by2_dropdown.setEnabled(True)
        except Exception as e:
            self.logger.error(f"Could not load 'Organize by' columns: {e}", exc_info=True)
            QMessageBox.critical(self, "Error", f"Could not load 'Organize by' columns:\n{e}")
            self.organize_by_dropdown.setEnabled(False)
            self.organize_by2_dropdown.setEnabled(False)


    def populate_tableau_filter_field_dropdown(self, columns):
        """Populates the 'Tableau Filter Field' dropdown with Excel columns."""
        print("Populating Tableau filter field dropdown...")
        try:
            self.tableau_filter_field_dropdown.clear()
            self.tableau_filter_field_dropdown.addItems(columns)
            self.tableau_filter_field_dropdown.setEnabled(True)
            self.tableau_filter_field_dropdown.setCurrentIndex(-1) # No initial selection
            self.tableau_filter_field_dropdown.setPlaceholderText("Select Excel Column (Optional)")
        except Exception as e:
            self.logger.error(f"Could not load Tableau filter fields: {e}", exc_info=True)
            QMessageBox.critical(self, "Error", f"Could not load Tableau filter fields:\n{e}")
            self.tableau_filter_field_dropdown.setEnabled(False)

    def _update_logic_toggle_icon(self):
        """Sets the correct icon (expand/collapse) on the toggle button."""
        if not hasattr(self, 'logic_toggle_button'):
            return # Button not created yet

        icon_size = QSize(16, 16) # Adjust size as needed
        self.logic_toggle_button.setIconSize(icon_size)
        icon_path = ""

        if self.logic_section_visible: # Section is visible, show "collapse" icon
            icon_path_svg = os.path.join(self.current_dir(), 'icons', 'chevron_down.svg')
            icon_path_png = os.path.join(self.current_dir(), 'icons', 'chevron_down.png')
            fallback_text = "▼"
        else: # Section is hidden, show "expand" icon
            icon_path_svg = os.path.join(self.current_dir(), 'icons', 'chevron_right.svg')
            icon_path_png = os.path.join(self.current_dir(), 'icons', 'chevron_right.png')
            fallback_text = "►"

        # Prefer SVG
        if os.path.exists(icon_path_svg):
            icon_path = icon_path_svg
        elif os.path.exists(icon_path_png):
            icon_path = icon_path_png

        try:
            if icon_path:
                self.logic_toggle_button.setIcon(QIcon(icon_path))
                self.logic_toggle_button.setText("") # Clear text when icon loads
            else:
                self.logger.warning(f"Toggle icon not found for state visible={self.logic_section_visible}")
                self.logic_toggle_button.setIcon(QIcon()) # Clear icon
                self.logic_toggle_button.setText(fallback_text) # Use text fallback
        except Exception as e:
            self.logger.error(f"Error loading toggle icon: {e}")
            self.logic_toggle_button.setIcon(QIcon())
            self.logic_toggle_button.setText(fallback_text)

    def _update_export_settings_toggle_icon(self):
        """Sets the correct icon (expand/collapse) on the Export Settings toggle button."""
        if not hasattr(self, 'export_settings_toggle_button'):
            return # Button not created yet

        icon_size = QSize(16, 16) # Adjust size as needed
        self.export_settings_toggle_button.setIconSize(icon_size)
        icon_path = ""
        fallback_text = "" # Initialize fallback_text

        if self.export_settings_section_visible: # Section is visible, show "collapse" icon
            icon_path_svg = os.path.join(self.current_dir(), 'icons', 'chevron_down.svg')
            icon_path_png = os.path.join(self.current_dir(), 'icons', 'chevron_down.png')
            fallback_text = "▼"
        else: # Section is hidden, show "expand" icon
            icon_path_svg = os.path.join(self.current_dir(), 'icons', 'chevron_right.svg')
            icon_path_png = os.path.join(self.current_dir(), 'icons', 'chevron_right.png')
            fallback_text = "►"

        # Prefer SVG
        if os.path.exists(icon_path_svg):
            icon_path = icon_path_svg
        elif os.path.exists(icon_path_png):
            icon_path = icon_path_png

        try:
            if icon_path:
                self.export_settings_toggle_button.setIcon(QIcon(icon_path))
                self.export_settings_toggle_button.setText("") # Clear text when icon loads
            else:
                self.logger.warning(f"Toggle icon not found for Export Settings state visible={self.export_settings_section_visible}")
                self.export_settings_toggle_button.setIcon(QIcon()) # Clear icon
                self.export_settings_toggle_button.setText(fallback_text) # Use text fallback
        except Exception as e:
            self.logger.error(f"Error loading Export Settings toggle icon: {e}")
            self.export_settings_toggle_button.setIcon(QIcon())
            self.export_settings_toggle_button.setText(fallback_text)

    # Replace this method inside PDFExportApp class
    def toggle_logic_section(self):
        """Shows or hides the logic content widget."""
        # The button's checked state drives the visibility state
        if hasattr(self, 'logic_toggle_button') and hasattr(self, 'logic_content_widget'):
            self.logic_section_visible = self.logic_toggle_button.isChecked() # Update state from button
            self.logic_content_widget.setVisible(self.logic_section_visible)
            self._update_logic_toggle_icon() # Update the icon
            self.logger.debug(f"Logic section visibility toggled to: {self.logic_section_visible}")
            # *** Use helper to adjust height ***
            self._adjust_window_height()
            self.logger.debug(f"Window size after logic toggle: {self.size()}")

    def toggle_export_settings_section(self):
        """Shows or hides the export settings content widget."""
        if hasattr(self, 'export_settings_toggle_button') and hasattr(self, 'exportSettingsContentWidget'):
            # Update state based on the button's checked status
            self.export_settings_section_visible = self.export_settings_toggle_button.isChecked()
            self.exportSettingsContentWidget.setVisible(self.export_settings_section_visible)
            self._update_export_settings_toggle_icon() # Update the icon
            self.logger.debug(f"Export Settings section visibility toggled to: {self.export_settings_section_visible}")
            # *** Use helper to adjust height ***
            self._adjust_window_height()
            self.logger.debug(f"Window size after export settings toggle: {self.size()}")
  
    # --- Combined Logic Panel Setup ---
    # Inside PDFExportApp class
    def setup_combined_logic_panel(self, main_content_layout):
        """Sets up the Filtering/Logic section with an inline toggle."""
        print("Setting up combined logic panel (inline toggle)...")

        # --- Main Container for this whole section ---
        # This widget will be hidden/shown based on the export mode
        self.logic_container_widget = QWidget()
        self.logic_container_widget.setObjectName("logicContainerWidget")
        container_layout = QVBoxLayout(self.logic_container_widget)
        container_layout.setContentsMargins(0, 5, 0, 5) # Add some vertical margin
        container_layout.setSpacing(0) # No space between title bar and content

        # --- Custom "Title Bar" ---
        title_bar_widget = QWidget()
        title_bar_layout = QHBoxLayout(title_bar_widget)
        title_bar_layout.setContentsMargins(5, 3, 5, 3) # Padding inside title bar
        title_bar_layout.setSpacing(8)

        # Title Label (Styled like a GroupBox title)
        title_label = QLabel("Filtering & Conditional Logic")
        title_label.setObjectName("logicTitleLabel")
        title_label.setStyleSheet("font-weight: bold; color: #444444;") # Style the title

        # Toggle Button (Icon Only)
        self.logic_toggle_button = QToolButton()
        self.logic_toggle_button.setObjectName("logicToggleButton")
        self.logic_toggle_button.setToolTip("Show/Hide Logic Details")
        self.logic_toggle_button.setCursor(Qt.PointingHandCursor)
        self.logic_toggle_button.setCheckable(True) # The button state tracks visibility
        self.logic_toggle_button.setChecked(self.logic_section_visible) # Set initial check state
        self.logic_toggle_button.setStyleSheet("""
            QToolButton#logicToggleButton {
                border: none;
                background-color: transparent;
                padding: 2px;
            }
            QToolButton#logicToggleButton:hover {
                background-color: rgba(0, 0, 0, 0.1);
                border-radius: 3px;
            }
        """)
        self.logic_toggle_button.clicked.connect(self.toggle_logic_section)
        self._update_logic_toggle_icon() # Set initial icon

        # Add elements to title bar layout
        title_bar_layout.addWidget(self.logic_toggle_button) # Icon first
        title_bar_layout.addWidget(title_label)
        title_bar_layout.addStretch(1) # Push title/icon left

        # Add title bar to the main container layout
        container_layout.addWidget(title_bar_widget)

        # --- Content Widget (Holds the actual filters/conditions) ---
        self.logic_content_widget = QWidget()
        self.logic_content_widget.setObjectName("logicContentWidget")
        # Use the stored layout reference for the content
        self.combined_logic_layout = QVBoxLayout(self.logic_content_widget)
        self.combined_logic_layout.setContentsMargins(10, 5, 10, 10) # Padding inside content area
        self.combined_logic_layout.setSpacing(8) # Spacing between filter/condition lines

        # --- Add/Reset Buttons Row --- (Goes INSIDE logic_content_widget)
        buttons_layout = QHBoxLayout()
        add_filter_btn = QPushButton("➕ Filter")
        add_filter_btn.setToolTip("Add Excel list filter")
        add_filter_btn.clicked.connect(self.add_filter_line)

        add_condition_btn = QPushButton("➕ Condition")
        add_condition_btn.setToolTip("Exclude views based on Excel data")
        add_condition_btn.clicked.connect(self.add_condition_line)

        add_parameter_btn = QPushButton("➕ Parameter")
        add_parameter_btn.setToolTip("Override Tableau Parameters")
        add_parameter_btn.clicked.connect(self.add_parameter_line)

        reset_all_btn = QPushButton("🔄 Reset")
        reset_all_btn.setToolTip("Remove all Filters, Conditions, and Parameters")
        reset_all_btn.clicked.connect(self.reset_combined_logic)

        buttons_layout.addWidget(add_filter_btn)
        buttons_layout.addWidget(add_condition_btn)
        buttons_layout.addWidget(add_parameter_btn)
        buttons_layout.addStretch()
        buttons_layout.addWidget(reset_all_btn)

        buttons_container = QWidget()
        buttons_container.setLayout(buttons_layout)
        self.combined_logic_layout.addWidget(buttons_container) # Add button bar

        # Add a stretch item to push dynamic lines to the top within the content
        self.combined_logic_layout.addStretch(1)

        # Add the content widget (initially hidden/shown based on state)
        container_layout.addWidget(self.logic_content_widget)
        self.logic_content_widget.setVisible(self.logic_section_visible)

        # Add the entire container to the main application layout
        main_content_layout.addWidget(self.logic_container_widget)

        # --- Final Visibility Check based on Mode ---
        current_mode = self.mode_selection_dropdown.currentText() if hasattr(self, 'mode_selection_dropdown') else "Automate for a list"
        is_automate_mode = (current_mode == "Automate for a list")
        self.logic_container_widget.setVisible(is_automate_mode) # Hide WHOLE section if not automate mode
    
    # --- NEW HELPER FUNCTION ---
    def filter_list_items(self, text, list_widget):
        """Filters items in the QListWidget based on the search text."""
        search_text = text.lower()
        for i in range(list_widget.count()):
            item = list_widget.item(i)
            if item:
                item_text = item.text().lower()
                # Hide item if text doesn't contain the search term
                is_hidden = search_text not in item_text
                list_widget.setItemHidden(item, is_hidden)

    def add_filter_line(self, filter_data=None):
        """Adds a new row for defining a filter criterion (Field + Value List + Apply as Param)."""
        print("Adding filter line (QLineEdit + QListWidget + Param Checkbox)...")
        self.logger.debug("Adding new filter line to UI.")

        # --- Main Horizontal Layout for the Filter Row ---
        hbox = QHBoxLayout()
        hbox.setSpacing(5)

        # --- Field Selection ---
        field_label = QLabel("Filter:")
        field_combo = QComboBox()
        field_combo.setPlaceholderText("Select Field")
        field_combo.setMinimumWidth(180)
        field_combo.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        field_combo.setEnabled(False)

        # --- Value Selection Area (Vertical Layout) ---
        value_area_layout = QVBoxLayout()
        value_area_layout.setSpacing(2) # Compact spacing

        # Combined dropdown for search + check
        values_combo = CheckableComboBox()
        values_combo.setMinimumWidth(250)
        values_combo.setEnabled(False)

        # Add Filter Edit and List Widget to the vertical layout
        value_area_layout.addWidget(values_combo)
        # --- End Value Selection Area ---

        # --- ADD: Apply as Parameter Checkbox ---
        apply_as_param_checkbox = QCheckBox("Apply as Param")
        apply_as_param_checkbox.setToolTip("If checked, apply this filter's field and first selected value as a Tableau Parameter override.\nIn Tableau, the parameter name should be [Filter field nam]&'Param'")
        apply_as_param_checkbox.setChecked(False) # Default to unchecked
        # --- End Add Checkbox ---


        # --- Delete Button ---
        del_btn = QPushButton("🗑️")
        del_btn.setObjectName("deleteItemButton") # Set object name for QSS styling
        del_btn.setToolTip("Delete this filter")
        del_btn.setCursor(Qt.PointingHandCursor)
        # --- End Delete Button ---

        # --- Store Filter Information ---
        # IMPORTANT: Added 'apply_as_param_checkbox'
        filter_info = {
            'hbox': hbox,
            'field_combo': field_combo,
            'values_combo': values_combo,
            'apply_as_param_checkbox': apply_as_param_checkbox, # Store the checkbox
            'selected_values': [], # Store checked values (text) when loading/saving
            'apply_param_on_load': False # Temporary flag for loading config
        }
        self.filters.append(filter_info)
        # --- End Storage ---

        # --- Connect Signals ---
        # Populate list when field changes
        field_combo.currentIndexChanged.connect(
            lambda index, fc=field_combo, vc=values_combo: self.populate_values_list(fc, vc)
        )

        del_btn.clicked.connect(lambda: self.remove_filter_line(hbox))
        # --- End Signal Connections ---

        # --- Populate Widgets if Loading from Configuration ---
        saved_field = '' # Initialize
        if filter_data and isinstance(filter_data, dict):
            saved_field = filter_data.get('field', '')
            saved_values_str = filter_data.get('values', '')
            # Store the text of values to be checked later in populate_values_list
            filter_info['selected_values'] = [v.strip() for v in saved_values_str.split(',') if v.strip()]
            # Store the checkbox state for later application
            filter_info['apply_param_on_load'] = filter_data.get('apply_as_param', False)
            # The checkbox itself will be checked after columns are loaded and field is potentially set

        # --- Add Widgets to Main Horizontal Layout ---
        hbox.addWidget(field_label)
        hbox.addWidget(field_combo, 2) # Stretch factor 2
        hbox.addWidget(QLabel("Values:"), 0) # No stretch for label
        hbox.addLayout(value_area_layout, 3) # Add the vertical layout (stretch 3)
        hbox.addWidget(apply_as_param_checkbox, 1) # Add the checkbox
        hbox.addWidget(del_btn, 0) # No stretch
        # --- End Adding Widgets ---

        # --- Insert Layout into UI ---
        stretch_index = self.combined_logic_layout.count() - 1
        self.combined_logic_layout.insertLayout(stretch_index, hbox)
        # --- End Insertion ---

        # --- Attempt to Load Columns for the New Line's Field Combo ---
        if self.excel_path and self.sheet_name.currentText():
            try:
                df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name.currentText(), nrows=1)
                columns = list(df.columns)
                field_combo.clear()
                field_combo.addItems(columns)
                field_combo.setEnabled(True)

                # Now set the field combo text if loading and field is valid
                if saved_field and saved_field in columns:
                    field_combo.blockSignals(True)
                    field_combo.setCurrentText(saved_field)
                    field_combo.blockSignals(False)
                    # Trigger population for the loaded field (this will also handle checking values)
                    self.populate_values_list(field_combo, values_combo)
                    # Now that values are potentially loaded, check the checkbox based on the temp flag
                    apply_as_param_checkbox.setChecked(filter_info['apply_param_on_load'])
                else:
                    # If not loading a valid saved field, ensure checkbox state reflects default/flag
                    apply_as_param_checkbox.setChecked(filter_info['apply_param_on_load'])

            except Exception as e:
                self.logger.warning(f"Could not immediately load columns/populate for new filter line: {e}")
                # Still try to set the checkbox state based on the loaded data if available
                apply_as_param_checkbox.setChecked(filter_info.get('apply_param_on_load', False))

    # --- REPLACE THIS FUNCTION ---
    def populate_values_list(self, field_combo, values_combo):
        """Populates the QListWidget with unique checkable values from the selected field."""
        selected_field = field_combo.currentText()
        values_combo.clear() # Clear previous values
        values_combo.setEnabled(False) # Disable until populated
        self.logger.debug(f"Populating values list for filter field: '{selected_field}'")

        if not selected_field or not self.excel_path or not self.sheet_name.currentText():
            self.logger.debug("Skipping value population: Field/Excel/Sheet not selected.")
            return

        # Find the corresponding filter_info to access 'selected_values' for loading
        filter_info = next((f for f in self.filters if f['field_combo'] == field_combo), None)
        values_to_check_on_load = set(filter_info['selected_values']) if filter_info else set()
        if filter_info:
            filter_info['selected_values'] = [] # Clear temp storage after getting values

        try:
            df_col = pd.read_excel(self.excel_path, sheet_name=self.sheet_name.currentText(), usecols=[selected_field])
            # Ensure values are strings for consistent comparison
            unique_values = sorted([str(v) for v in df_col[selected_field].dropna().unique()])

            if unique_values:
                self.logger.debug(f"Found {len(unique_values)} unique values for '{selected_field}'. Populating list.")
                values_combo.blockSignals(True)
                for value_text in unique_values:
                    values_combo.addItem(value_text, checked=(value_text in values_to_check_on_load))
                values_combo.blockSignals(False)
                values_combo.setEnabled(True)

            else:
                 self.logger.warning(f"No unique values found for field '{selected_field}' in sheet '{self.sheet_name.currentText()}'.")

        except KeyError:
             self.logger.error(f"Field '{selected_field}' not found in sheet '{self.sheet_name.currentText()}' during value population.")
             QMessageBox.warning(self, "Error", f"Column '{selected_field}' not found in the selected sheet.")
        except Exception as e:
            self.logger.error(f"Error populating filter values list for field '{selected_field}': {e}", exc_info=True)
            QMessageBox.critical(self, "Error", f"Could not load values for field '{selected_field}':\n{e}")


    # --- REPLACE THIS FUNCTION ---
    def remove_filter_line(self, hbox):
        """Removes a filter row (hbox) and its associated data."""
        print("Removing filter line...")
        filter_to_remove = next((f for f in self.filters if f['hbox'] == hbox), None)
        if filter_to_remove:
            field_name = "N/A"
            try:
                # Attempt to get field name, handle potential deletion errors
                if filter_to_remove.get('field_combo'):
                    field_name = filter_to_remove['field_combo'].currentText()
            except RuntimeError: # Handles cases where the widget might already be deleted
                self.logger.warning("Could not get field name during filter removal, widget might be deleted.")

            self.logger.debug(f"Removing filter: {field_name}")

            # Use the helper function to clear items from the layout
            self._clear_layout_items(hbox)

            # Remove the hbox itself from its parent layout
            parent_layout = self.combined_logic_layout # Assuming this is the direct parent
            if parent_layout:
                parent_layout.removeItem(hbox)
            else:
                 self.logger.error("Could not find parent layout for filter hbox during removal.")

            hbox.deleteLater() # Schedule the layout itself for deletion

            # Remove the entry from the internal list
            try:
                self.filters.remove(filter_to_remove)
            except ValueError:
                 self.logger.warning("Attempted to remove filter_info that was already removed from the list.")

            # Optional: Force layout update if needed, though usually not required
            # self.combined_logic_layout.update()
            # self.layout().activate()
        else:
            self.logger.error("Error: Could not find filter layout/info to remove.")


    # --- MODIFY THIS FUNCTION ---
    def load_columns_for_filters(self, column_names, specific_filter=None):
        """Populates the field dropdown for filter lines."""
        filters_to_update = [specific_filter] if specific_filter else self.filters
        self.logger.debug(f"Loading columns {column_names} for {len(filters_to_update)} filter field(s).")
        for filt_info in filters_to_update:
             # Check for the new structure
             if isinstance(filt_info, dict) and 'field_combo' in filt_info and 'values_combo' in filt_info:
                field_combo = filt_info['field_combo']
                values_combo = filt_info['values_combo'] # Corrected to values_combo
                current_field_selection = field_combo.currentText()

                field_combo.clear()
                field_combo.addItems(column_names)
                field_combo.setEnabled(True)

                if current_field_selection in column_names:
                    field_combo.blockSignals(True)
                    field_combo.setCurrentText(current_field_selection)
                    field_combo.blockSignals(False)
                    # Re-populate the list for the selected field
                    self.populate_values_list(field_combo, values_combo) # Corrected to values_combo
                else:
                    field_combo.setCurrentIndex(-1)
                    field_combo.setPlaceholderText("Select Field")
                    values_combo.clear() # Corrected to values_combo
                    values_combo.setEnabled(False) # Corrected to values_combo
             else:
                  self.logger.warning(f"Skipping column load for invalid/outdated filter structure: {filt_info}")

    def add_condition_line(self, condition=None):
        """Adds a new row for defining a view exclusion condition."""
        print("Adding condition line...")
        self.logger.debug("Adding new condition line to UI.")
        hbox = QHBoxLayout()
        hbox.setSpacing(5)

        column_txt = QComboBox()
        column_txt.setEnabled(False); column_txt.setPlaceholderText("Select Excel Field"); column_txt.setMinimumWidth(150)
        type_choice = QComboBox()
        type_choice.addItems(['Equals', 'Not Equals', 'Greater Than', 'Less Than', 'Is Blank', 'Is Not Blank']); type_choice.setMinimumWidth(100)
        value_txt = QLineEdit(); value_txt.setPlaceholderText("Value"); value_txt.setMinimumWidth(120)
        views_btn = QPushButton("Exclude"); views_btn.setToolTip("Select views to EXCLUDE if this condition is met"); views_btn.setMinimumWidth(100)
        del_btn = QPushButton("🗑️") # Keep Unicode character
        del_btn.setObjectName("deleteItemButton") # Set SAME object name
        del_btn.setToolTip("Delete this condition")
        del_btn.setCursor(Qt.PointingHandCursor) # Add for consistency

        exclude_views = []
        saved_column = '' # Initialize
        if condition and isinstance(condition, dict):
            saved_column = condition.get('column', '')
            saved_type = condition.get('type', 'Equals')
            saved_value = condition.get('value', '')
            saved_excluded_views = condition.get('excluded_views', [])
            exclude_views = [str(v) for v in saved_excluded_views if isinstance(v, (str, int, float))]
            type_choice.setCurrentText(saved_type)
            value_txt.setText(saved_value)

        condition_info = { 'hbox': hbox, 'column_txt': column_txt, 'type_choice': type_choice, 'value_txt': value_txt, 'views_btn': views_btn, 'excluded_views': exclude_views }
        self.conditions.append(condition_info)

        views_btn.clicked.connect(lambda checked, h=hbox, ev_list_ref=exclude_views: self.select_views_for_condition(h, ev_list_ref))
        type_choice.currentIndexChanged.connect(lambda index, tc=type_choice, vt=value_txt: self.on_condition_type_changed(tc, vt))
        del_btn.clicked.connect(lambda: self.remove_condition_line(hbox))
        self.on_condition_type_changed(type_choice, value_txt) # Set initial state

        hbox.addWidget(QLabel("If:"))
        hbox.addWidget(column_txt, 2); hbox.addWidget(type_choice, 1); hbox.addWidget(value_txt, 2)
        hbox.addWidget(QLabel("Then"))
        hbox.addWidget(views_btn, 1); hbox.addWidget(del_btn, 0, Qt.AlignRight)

        # Insert into the *combined* layout
        stretch_index = self.combined_logic_layout.count() - 1
        self.combined_logic_layout.insertLayout(stretch_index, hbox)

        if self.excel_path and self.sheet_name.currentText():
            try:
                 df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name.currentText(), nrows=1)
                 columns = list(df.columns)
                 self.load_columns_for_conditions(columns, specific_condition=condition_info)
                 if saved_column and saved_column in columns:
                      column_txt.setCurrentText(saved_column)
            except Exception as e:
                 self.logger.warning(f"Could not immediately load columns for new condition: {e}")


    def on_condition_type_changed(self, type_choice_widget, value_txt_widget):
        """Enables/disables the value QLineEdit based on condition type."""
        condition_type = type_choice_widget.currentText()
        is_blank_type = condition_type in ['Is Blank', 'Is Not Blank']
        value_txt_widget.setEnabled(not is_blank_type)
        is_numeric_type = condition_type in ['Greater Than', 'Less Than']
        if is_blank_type:
            value_txt_widget.clear(); value_txt_widget.setPlaceholderText("")
        elif is_numeric_type: value_txt_widget.setPlaceholderText("Enter Number")
        else: value_txt_widget.setPlaceholderText("Enter Text Value")


    def add_parameter_line(self, parameter=None):
        """Adds a new row for defining a Tableau parameter override."""
        print("Adding parameter line...")
        self.logger.debug("Adding new parameter line to UI.")
        hbox = QHBoxLayout(); hbox.setSpacing(5)

        param_name_txt = QLineEdit(); param_name_txt.setPlaceholderText("Parameter Name on Tableau"); param_name_txt.setMinimumWidth(150)
        param_value_txt = QLineEdit(); param_value_txt.setPlaceholderText("Value (Static or Excel Column Name)"); param_value_txt.setMinimumWidth(200)
        del_btn = QPushButton("🗑️") # Keep Unicode character
        del_btn.setObjectName("deleteItemButton") # Set SAME object name
        del_btn.setToolTip("Delete this parameter override")
        del_btn.setCursor(Qt.PointingHandCursor) # Add for consistency

        if parameter and isinstance(parameter, dict):
            param_name_txt.setText(parameter.get('name', ''))
            param_value_txt.setText(parameter.get('value', ''))

        self.parameters.append({ 'hbox': hbox, 'param_name_txt': param_name_txt, 'param_value_txt': param_value_txt })
        del_btn.clicked.connect(lambda: self.remove_parameter_line(hbox))

        hbox.addWidget(QLabel("Parameter:")); hbox.addWidget(param_name_txt, 2)
        hbox.addWidget(QLabel("Value:")); hbox.addWidget(param_value_txt, 3)
        hbox.addWidget(del_btn, 0, Qt.AlignRight)

        # Insert into the *combined* layout
        stretch_index = self.combined_logic_layout.count() - 1
        self.combined_logic_layout.insertLayout(stretch_index, hbox)


    def remove_condition_line(self, hbox):
         """Removes a condition row (hbox) and its associated data."""
         print("Removing condition line...")
         condition_to_remove = next((cond for cond in self.conditions if cond['hbox'] == hbox), None)
         if condition_to_remove:
             cond_name = "N/A"
             try:
                 if condition_to_remove.get('column_txt'):
                     cond_name = condition_to_remove['column_txt'].currentText()
             except RuntimeError:
                 self.logger.warning("Could not get condition name during removal, widget might be deleted.")

             self.logger.debug(f"Removing condition: {cond_name}")
             self._clear_layout_items(hbox) # Use helper to clear layout
             parent_layout = self.combined_logic_layout
             if parent_layout: parent_layout.removeItem(hbox)
             hbox.deleteLater()
             try:
                 self.conditions.remove(condition_to_remove)
             except ValueError:
                 self.logger.warning("Attempted to remove condition_info that was already removed from the list.")
         else:
             self.logger.error("Error: Could not find condition layout/info to remove.")


    def remove_parameter_line(self, hbox):
         """Removes a parameter row (hbox) and its associated data."""
         print("Removing parameter line...")
         parameter_to_remove = next((param for param in self.parameters if param['hbox'] == hbox), None)
         if parameter_to_remove:
              param_name = "N/A"
              try:
                  if parameter_to_remove.get('param_name_txt'):
                      param_name = parameter_to_remove['param_name_txt'].text()
              except RuntimeError:
                  self.logger.warning("Could not get parameter name during removal, widget might be deleted.")

              self.logger.debug(f"Removing parameter: {param_name}")
              self._clear_layout_items(hbox) # Use helper to clear layout
              parent_layout = self.combined_logic_layout
              if parent_layout: parent_layout.removeItem(hbox)
              hbox.deleteLater()
              try:
                  self.parameters.remove(parameter_to_remove)
              except ValueError:
                  self.logger.warning("Attempted to remove parameter_info that was already removed from the list.")
         else:
              self.logger.error("Error: Could not find parameter layout/info to remove.")


    def select_views_for_condition(self, hbox, current_excluded_views_list):
        print("Selecting views for conditional exclusion...") # Log updated purpose
        if not self.tableau_views:
            QMessageBox.warning(self, "No Views Loaded", "Please test connection and load views first.")
            self.logger.warning("Select views for condition attempted before loading views.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Select Views to EXCLUDE for this Condition") # Title updated
        dialog.setMinimumWidth(400)
        layout = QVBoxLayout(dialog) # Set layout on dialog

        # *** UPDATED INSTRUCTION ***
        info_label = QLabel("Check views to EXCLUDE only when this condition is met.\n"
                            "(Globally excluded views are disabled and marked).")
        layout.addWidget(info_label)

        list_widget = QListWidget()
        list_widget.setSpacing(2)
         # *** APPLY CUSTOM DELEGATE ***
        list_widget.setItemDelegate(RedXCheckDelegate(list_widget))


        current_condition_excluded_set = set(current_excluded_views_list) # Views excluded by this specific condition
        globally_excluded_set = set(self.excluded_views_for_export) # Views excluded everywhere

        for view_name in self.tableau_views:
            item = QListWidgetItem(view_name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)

            if view_name in globally_excluded_set:
                # *** Handle Globally Excluded ***
                item.setFlags(item.flags() & ~Qt.ItemIsEnabled) # Disable item
                item.setCheckState(Qt.Checked) # Mark as checked (excluded)
                # Delegate will handle grey text and red X
            else:
                # *** Handle Conditionally Excluded ***
                # Check if it IS excluded by THIS condition
                item.setCheckState(Qt.Checked if view_name in current_condition_excluded_set else Qt.Unchecked)
                # Delegate will handle red X if checked

            list_widget.addItem(item)

        button_layout = QHBoxLayout()
        # *** UPDATE BUTTON LABELS/TOOLTIPS ***
        tick_all_btn = QPushButton("Exclude All Applicable")
        tick_all_btn.setToolTip("Check all enabled views to exclude them for this condition.")
        untick_all_btn = QPushButton("Include All Applicable")
        untick_all_btn.setToolTip("Uncheck all enabled views to include them for this condition.")

        # Modify lambda to only affect enabled items
        tick_all_btn.clicked.connect(lambda: self.set_enabled_check_states(list_widget, Qt.Checked))
        untick_all_btn.clicked.connect(lambda: self.set_enabled_check_states(list_widget, Qt.Unchecked))

        button_layout.addWidget(untick_all_btn)
        button_layout.addWidget(tick_all_btn)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        # Pass the original list reference to be modified
        button_box.accepted.connect(lambda: self.apply_condition_view_selection(dialog, list_widget, current_excluded_views_list))
        button_box.rejected.connect(dialog.reject)

        layout.addWidget(list_widget)
        layout.addLayout(button_layout)
        layout.addWidget(button_box)

        dialog.exec_()

    def set_enabled_check_states(self, list_widget, state):
        """ Sets the check state of all ENABLED items in a QListWidget. """
        for index in range(list_widget.count()):
            item = list_widget.item(index)
            if item and item.flags() & Qt.ItemIsEnabled: # Only change enabled items
                item.setCheckState(state)

    def apply_condition_view_selection(self, dialog, list_widget, condition_excluded_list_ref):
        """Applies view selections from the dialog directly to the condition's exclusion list."""
        print("Applying conditional view exclusion selection...")
        newly_excluded_for_condition = []
        # Iterate through all items to build the list for this condition
        for index in range(list_widget.count()):
            item = list_widget.item(index)
            if item:
                 # *** REVERSED LOGIC: Include if checked AND ENABLED ***
                 # We only care about conditionally excluded items here.
                 # Globally excluded items remain globally excluded regardless of this dialog.
                 is_enabled = item.flags() & Qt.ItemIsEnabled
                 if is_enabled and item.checkState() == Qt.Checked:
                    newly_excluded_for_condition.append(item.text())

        # IMPORTANT: Modify the list passed by reference
        condition_excluded_list_ref.clear()
        condition_excluded_list_ref.extend(newly_excluded_for_condition)

        # Logging/Feedback
        cond_text = "condition" # Simple fallback text
        # You might want to find the specific condition text again if needed for logging
        # ... (code to find condition text based on hbox if necessary) ...

        self.logger.info(f"View exclusion updated for {cond_text}. Conditionally excluding: {len(newly_excluded_for_condition)} view(s).")
        dialog.accept()


    def setup_progress_panel(self, layout):
        """Sets up the GroupBox for the progress bar and log output."""
        print("Setting up progress panel...")
        progress_box = QGroupBox("Progress Log")
        # progress_box.setFont(QFont('Arial', 10, QFont.Normal)) # Font set globally
        progress_layout = QVBoxLayout()
        progress_box.setLayout(progress_layout)
        self.progress_bar = QProgressBar(); self.progress_bar.setTextVisible(True)
        self.progress_bar.setAlignment(Qt.AlignCenter); self.progress_bar.setFormat("0%")
        self.log_text = QTextEdit(); self.log_text.setReadOnly(True)
        # self.log_text.setFont(QFont('Consolas', 9)); # Use global font
        self.log_text.setLineWrapMode(QTextEdit.WidgetWidth)
        progress_layout.addWidget(self.progress_bar); progress_layout.addWidget(self.log_text)
        layout.addWidget(progress_box)


    def setup_control_buttons(self, layout):
        """Sets up the main control buttons (Start, Stop, Load, Save, Server Config)."""
        print("Setting up control buttons...")
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)

        self.start_btn = QPushButton("▶️ Start Export")
        self.start_btn.setToolTip("Start export process")

        self.stop_btn = QPushButton("⏹️ Stop Export")
        self.stop_btn.setEnabled(False)
        self.stop_btn.setToolTip("Stop running export")

        self.load_btn = QToolButton(self) # Keep as QToolButton
        self.load_btn.setText("📂 Load Config")
        self.load_btn.setPopupMode(QToolButton.MenuButtonPopup) # Keep menu mode
        self.load_btn.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.load_btn.setMinimumHeight(28)
        self.load_btn.setToolTip("Load saved settings or select recent") # Updated tooltip
        self.load_btn.setObjectName("loadConfigButton")

        load_menu = QMenu(self.load_btn)
        self.load_btn.setMenu(load_menu)
        load_action = QAction("Load New Configuration...", self)
        load_action.triggered.connect(self.load_configuration_file) # Menu item still works
        load_menu.addAction(load_action)
        self.load_recent_files(load_menu)

        # *** ADDED: Connect main button click to load function ***
        self.load_btn.clicked.connect(self.load_configuration_file)

        self.save_btn = QPushButton("💾 Save Config")
        self.save_btn.setToolTip("Save current settings")

        self.server_config_btn = QPushButton("⚙️ Server Config")
        self.server_config_btn.setToolTip("Configure Tableau Server connection")

        self.toggle_theme_btn = QPushButton("🎨 Theme") # Shortened Text
        self.toggle_theme_btn.setToolTip("Switch between custom and default styles")
        self.toggle_theme_btn.setObjectName("toggleThemeButton")
        self.toggle_theme_btn.clicked.connect(self.toggle_theme)
        self._update_theme_button_text() # Set initial tooltip based on theme


        buttons = [self.start_btn, self.stop_btn, self.load_btn, self.save_btn, self.server_config_btn, self.toggle_theme_btn]
        for btn in buttons:
            btn.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
            if isinstance(btn, QPushButton):
                btn.setFixedHeight(28)


        button_layout.addWidget(self.start_btn)
        button_layout.addWidget(self.stop_btn)
        button_layout.addStretch(1)
        button_layout.addWidget(self.toggle_theme_btn)
        button_layout.addWidget(self.server_config_btn)
        button_layout.addWidget(self.load_btn)
        button_layout.addWidget(self.save_btn)

        self.start_btn.clicked.connect(self.OnStart)
        self.stop_btn.clicked.connect(self.OnStop)
        self.save_btn.clicked.connect(self.OnSave) # Changed to OnSave
        self.server_config_btn.clicked.connect(self.open_server_config_dialog)

        layout.addLayout(button_layout)
        
    def open_server_config_dialog(self):
        """Opens the Server Configuration dialog."""
        self.logger.info("Opening Server Configuration dialog.")
        dialog = ServerConfigDialog(self) # Pass main window as parent
        dialog.exec_() # Show the dialog modally


    def load_configuration_file(self):
        """Opens a file dialog to load a configuration (.ini) file."""
        print("Loading configuration file...")
        self.logger.info("Opening file dialog to load configuration.")

        # *** MODIFIED: Set default directory to 'config' subfolder ***
        app_dir = self.current_dir()
        config_dir = os.path.join(app_dir, "config")

        if os.path.isdir(config_dir):
            start_dir = config_dir
            self.logger.debug(f"Defaulting load dialog to existing folder: {start_dir}")
        else:
            start_dir = app_dir # Fallback to app directory if 'config' doesn't exist
            self.logger.debug(f"Defaulting load dialog to application folder (config subfolder not found): {start_dir}")

        file_name, _ = QFileDialog.getOpenFileName(self, "Load Configuration File", start_dir, "INI Files (*.ini)") # Use start_dir

        if file_name:
            self.load_configuration(file_name) # Load the selected file
            self.add_recent_file(file_name) # Add it to recent list
        else:
            self.logger.info("Configuration file loading cancelled.")

    def load_recent_files(self, menu):
        """Loads recent file paths from JSON and populates the Load button menu."""
        # Clear existing recent file actions (keep the "Load New..." action)
        actions = menu.actions()
        if len(actions) > 1: # If more than just "Load New..."
            # Remove separators and recent file actions
            for i in range(len(actions) - 1, 0, -1): # Iterate backwards from last item down to index 1
                 action = actions[i]
                 # Check if it's a separator or a recent file action (has file_path attribute)
                 if action.isSeparator() or hasattr(action, 'file_path'):
                     menu.removeAction(action)

        recent_files = []
        if os.path.exists(RECENT_FILES_PATH):
            try:
                with open(RECENT_FILES_PATH, 'r') as f:
                    content = f.read()
                    if content:
                         loaded_files = json.loads(content)
                         # Basic validation
                         if isinstance(loaded_files, list) and all(isinstance(p, str) for p in loaded_files):
                             recent_files = loaded_files
                         else:
                             self.logger.warning(f"Corrupted recent files data in {RECENT_FILES_PATH}. Resetting."); recent_files = []
            except (json.JSONDecodeError, ValueError, TypeError) as e:
                self.logger.error(f"Error decoding recent files JSON {RECENT_FILES_PATH}: {e}. Resetting."); recent_files = []
            except Exception as e:
                self.logger.error(f"Error reading recent files {RECENT_FILES_PATH}: {e}", exc_info=True); recent_files = []

        if recent_files:
             menu.addSeparator() # Add separator only if there are recent files
             for file_path in recent_files:
                 if file_path and os.path.exists(file_path):
                     file_name = os.path.basename(file_path)
                     action = QAction(file_name, self); action.setToolTip(file_path)
                     # Connect triggered signal to load_configuration lambda
                     action.triggered.connect(lambda checked=False, path=file_path: self.load_configuration(path))
                     action.file_path = file_path # Store path for identification
                     menu.addAction(action)
                 else:
                     self.logger.warning(f"Recent file path not found or invalid, skipping: {file_path}")


    def ensure_appdata_dir_exists(self):
        """Ensures the AppData directory for storing configuration exists."""
        if not os.path.exists(APPDATA_DIR):
            try:
                os.makedirs(APPDATA_DIR);
                self.logger.info(f"Created AppData directory: {APPDATA_DIR}")
            except Exception as e:
                self.logger.error(f"Error creating AppData directory {APPDATA_DIR}: {e}", exc_info=True);
                QMessageBox.critical(self, "Error", f"Failed to create application data directory:\n{APPDATA_DIR}\n\nError: {e}");
                return False
        return True


    def add_recent_file(self, file_path):
        """Adds a file path to the list of recent files, stored in JSON."""
        if not file_path or not isinstance(file_path, str):
            self.logger.warning(f"Attempted to add invalid recent file path: {file_path}"); return
        if not self.ensure_appdata_dir_exists(): return

        recent_files = []
        if os.path.exists(RECENT_FILES_PATH):
            try:
                with open(RECENT_FILES_PATH, 'r') as f:
                    content = f.read()
                    if content:
                         loaded_files = json.loads(content)
                         if isinstance(loaded_files, list): recent_files = loaded_files
            except (json.JSONDecodeError, ValueError, TypeError) as e:
                self.logger.warning(f"Could not parse recent files JSON, starting fresh: {e}"); recent_files = []
            except Exception as e:
                self.logger.error(f"Error reading recent files {RECENT_FILES_PATH}: {e}", exc_info=True); recent_files = []

        # Remove if already exists, then insert at the beginning
        if file_path in recent_files: recent_files.remove(file_path)
        recent_files.insert(0, file_path)

        # Limit the number of recent files
        max_recent = 5; recent_files = recent_files[:max_recent]

        try:
             with open(RECENT_FILES_PATH, 'w') as f: json.dump(recent_files, f, indent=4)
             self.logger.info(f"Added '{os.path.basename(file_path)}' to recent files.")
        except Exception as e:
            self.logger.error(f"Error writing recent files JSON {RECENT_FILES_PATH}: {e}", exc_info=True);
            QMessageBox.warning(self, "Warning", f"Could not save recent files list:\n{e}")

        # Update the menu immediately
        self.update_recent_files_menu()


    def update_recent_files_menu(self):
        """Updates the recent files menu associated with the Load button."""
        if hasattr(self, 'load_btn') and self.load_btn:
             load_menu = self.load_btn.menu()
             if load_menu:
                 self.load_recent_files(load_menu) # Call the function that rebuilds the menu
             else:
                 self.logger.warning("Load button found, but has no associated menu.")
        else:
            self.logger.warning("Load button reference ('self.load_btn') not found for updating recent files menu.")


    def save_configuration_file(self):
        """Opens a dialog to save the current configuration to an .ini file."""
        print("Saving configuration file...")
        self.logger.info("Opening file dialog to save configuration.")

        # --- MODIFIED: Set default directory to 'config' subfolder ---
        app_dir = self.current_dir()
        config_dir = os.path.join(app_dir, "config")

        if os.path.isdir(config_dir):
            start_dir = config_dir
            self.logger.debug(f"Defaulting save dialog to existing folder: {start_dir}")
        else:
            # Attempt to create the config directory if it doesn't exist but prefer app_dir if creation fails
            try:
                os.makedirs(config_dir, exist_ok=True)
                start_dir = config_dir
                self.logger.debug(f"Defaulting save dialog to created folder: {start_dir}")
            except OSError as e:
                self.logger.warning(f"Could not create config directory {config_dir}: {e}. Defaulting save dialog to application folder.")
                start_dir = app_dir # Fallback to app directory if 'config' doesn't exist or creation fails
        # --- End MODIFIED ---

        default_filename = os.path.join(start_dir, "pdf_export_config.ini")
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Configuration File", default_filename, "INI Files (*.ini)")
        if file_name:
            # Ensure the file has a .ini extension
            if not file_name.lower().endswith('.ini'):
                file_name += '.ini'
            self.OnSave(file_name)
            self.add_recent_file(file_name) # Add to recent list after successful save
        else:
            self.logger.info("Configuration file saving cancelled.")

    def load_configuration(self, file_path):
        """Loads configuration settings from the specified .ini file."""
        self.logger.info(f"Loading configuration from: {file_path}")
        print(f"Loading configuration from: {file_path}")
        self.reset_fields() # Reset everything first

        try:
            config = configparser.ConfigParser(interpolation=None)
            read_ok = config.read(file_path)
            if not read_ok:
                QMessageBox.critical(self, "Error", f"Could not read configuration file:\n{file_path}");
                self.logger.error(f"Failed to read configuration file: {file_path}"); return

            # --- Load Server (into attributes) (Keep as before) ---
            self.server_url_text = config.get('Server', 'url', fallback='')
            self.token_name_text = config.get('Server', 'token_name', fallback='')
            self.token_secret_text = config.get('Server', 'token_secret', fallback='')
            self.site_id_text = config.get('Server', 'site_id', fallback='')
            self.workbook_name.setText(config.get('Server', 'workbook_name', fallback=''))

            # --- Load Paths (Keep as before) ---
            self.excel_file.setText(config.get('Paths', 'excel_file', fallback=''))
            self.output_folder.setText(config.get('Paths', 'output_folder', fallback=''))
            self.excel_path = self.excel_file.text() # Update internal path variable

            # --- Load Sheets and Dependent Dropdowns (Keep as before) ---
            saved_sheet_name = config.get('Paths', 'sheet_name', fallback='')
            columns_loaded_correctly = False
            if self.excel_path and os.path.exists(self.excel_path):
                try:
                    self.load_sheets()
                    sheet_index = self.sheet_name.findText(saved_sheet_name, Qt.MatchFixedString) # Exact match
                    if sheet_index != -1:
                        self.logger.info(f"Setting sheet index to {sheet_index} ('{saved_sheet_name}') from config.")
                        self.sheet_name.setCurrentIndex(sheet_index)
                        columns_loaded_correctly = True
                    elif self.sheet_name.count() > 0:
                        self.logger.warning(f"Saved sheet '{saved_sheet_name}' not found. Using first sheet '{self.sheet_name.itemText(0)}'.")
                        columns_loaded_correctly = True
                    else:
                        self.logger.warning("Excel file loaded, but no sheets found.")
                        self.on_sheet_selection_changed()
                except Exception as e:
                    self.logger.error(f"Error during sheet loading/setting in load_configuration: {e}", exc_info=True)
                    QMessageBox.critical(self, "Sheet Loading Error", f"Error processing sheets from configuration:\n{e}")
                    self.on_sheet_selection_changed()
            else:
                self.logger.info("No valid Excel path in config, clearing sheet-dependent fields.")
                self.on_sheet_selection_changed()

            # --- Load General Settings (Keep as before) ---
            self.mode_selection_dropdown.setCurrentText(config.get('General', 'export_mode', fallback='Automate for a list'))
            self.toggle_mode_selection() # Apply visibility based on loaded mode

            if columns_loaded_correctly:
                self.file_naming_option.setCurrentText(config.get('General', 'file_naming', fallback='By view'))
                self.tableau_filter_field_dropdown.setCurrentText(config.get('Paths', 'tableau_filter_field', fallback=''))
                self.organize_by_dropdown.setCurrentText(config.get('General', 'organize_by', fallback='None'))
                self.organize_by2_dropdown.setCurrentText(config.get('General', 'organize_by2', fallback='None'))
            else:
                self.logger.warning("Skipping setting column-dependent options as columns failed to load.")

            export_format = config.get('General', 'export_format', fallback='PDF')
            self.pdf_radio.setChecked(export_format == 'PDF'); self.png_radio.setChecked(export_format == 'PNG')
            self.numbering_checkbox.setChecked(config.getboolean('General', 'numbering_enabled', fallback=True))
            
            # Load new checkboxes
            self.trim_pdf_checkbox.setChecked(config.getboolean('General', 'trim_pdf_enabled', fallback=False))
            self.merge_pdfs_checkbox.setChecked(config.getboolean('General', 'merge_pdfs_enabled', fallback=False))
            self.update_trim_merge_checkbox_states() # Update enabled states after loading

            excluded_views_str = config.get('General', 'excluded_views_for_export', fallback='')
            self.excluded_views_for_export = [view.strip() for view in excluded_views_str.split(',') if view.strip()]


            # --- Load Filters (MODIFIED) ---
            # Filters are reset in reset_fields, just load new ones
            filters_count = config.getint('Filters', 'count', fallback=0)
            loaded_filters_data = []
            for idx in range(filters_count):
                section = f'Filter_{idx}'
                if config.has_section(section):
                    filter_data = {
                        'field': config.get(section, 'field', fallback=''),
                        'values': config.get(section, 'values', fallback=''), # Comma-separated string of checked values
                        # *** ADDED: Load checkbox state ***
                        'apply_as_param': config.getboolean(section, 'apply_as_param', fallback=False) # Load boolean state
                    }
                    loaded_filters_data.append(filter_data)

            # Add filter lines using the loaded data
            # The 'selected_values' and 'apply_param_on_load' will be stored temporarily in filter_info
            # and applied when populate_values_list/columns are loaded during add_filter_line
            for filter_data in loaded_filters_data:
                self.add_filter_line(filter_data)


            # --- Load Conditions (Keep as before) ---
            conditions_count = config.getint('Conditions', 'count', fallback=0)
            loaded_conditions_data = []
            for idx in range(conditions_count):
                section = f'Condition_{idx}'
                if config.has_section(section):
                    condition_data = {
                        'column': config.get(section, 'column', fallback=''),
                        'type': config.get(section, 'type', fallback='Equals'),
                        'value': config.get(section, 'value', fallback=''),
                        'excluded_views': [v.strip() for v in config.get(section, 'excluded_views', fallback='').split(',') if v.strip()]
                    }
                    loaded_conditions_data.append(condition_data)
            for condition_data in loaded_conditions_data:
                self.add_condition_line(condition_data)


            # --- Load Parameters (Keep as before) ---
            parameters_count = config.getint('Parameters', 'count', fallback=0)
            loaded_parameters_data = []
            for idx in range(parameters_count):
                section = f'Parameter_{idx}'
                if config.has_section(section):
                    parameter_data = {
                        'name': config.get(section, 'name', fallback=''),
                        'value': config.get(section, 'value', fallback='')
                    }
                    loaded_parameters_data.append(parameter_data)
            for parameter_data in loaded_parameters_data:
                self.add_parameter_line(parameter_data)


            print("Configuration loaded successfully.")
            self.log_message_signal.emit(f"Configuration loaded from {os.path.basename(file_path)}")
            self.add_recent_file(file_path) # Add to recent list after successful load

            # *** ADDED: Adjust window height after loading all components ***
            self._adjust_window_height()

        except configparser.Error as e:
            QMessageBox.critical(self, "Config Error", f"Error parsing configuration file:\n{file_path}\n\n{e}")
            self.logger.error(f"Error parsing configuration file {file_path}: {e}")
        except Exception as e:
            QMessageBox.critical(self, "Load Error", f"An unexpected error occurred while loading configuration:\n{e}")
            self.logger.error(f"Unexpected error loading configuration {file_path}: {e}", exc_info=True)

    def OnSave(self, file_path):
        """Saves the current UI settings to the specified .ini file."""
        print(f"Saving configuration to: {file_path}")
        self.logger.info(f"Saving configuration to: {file_path}")
        config = configparser.ConfigParser(interpolation=None) # Prevent % interpolation issues

        # --- Save Server, Paths, General (Keep as before) ---
        config['Server'] = {
            'url': self.server_url_text,
            'token_name': self.token_name_text,
            'token_secret': self.token_secret_text, # Consider security implications
            'site_id': self.site_id_text,
            'workbook_name': self.workbook_name.text()
        }
        config['Paths'] = {
            'excel_file': self.excel_file.text(),
            'sheet_name': self.sheet_name.currentText(),
            'output_folder': self.output_folder.text(),
            'tableau_filter_field': self.tableau_filter_field_dropdown.currentText()
        }
        config['General'] = {
            'file_naming': self.file_naming_option.currentText(),
            'export_mode': self.mode_selection_dropdown.currentText(),
            'export_format': 'PDF' if self.pdf_radio.isChecked() else 'PNG',
            'organize_by': self.organize_by_dropdown.currentText(),
            'organize_by2': self.organize_by2_dropdown.currentText(),
            'numbering_enabled': str(self.numbering_checkbox.isChecked()),
            'trim_pdf_enabled': str(self.trim_pdf_checkbox.isChecked()), # Save trim checkbox state
            'merge_pdfs_enabled': str(self.merge_pdfs_checkbox.isChecked()), # Save merge checkbox state
            'excluded_views_for_export' : ','.join(self.excluded_views_for_export)
        }

        # --- Save Filters (MODIFIED) ---
        config['Filters'] = {'count': str(len(self.filters))}
        for idx, filt in enumerate(self.filters):
            section = f'Filter_{idx}'
            field = "N/A"
            checked_values = []
            apply_param_state = False # Default state
            try:
                # Get field name safely
                if filt.get('field_combo'):
                    field = filt['field_combo'].currentText()

                # Get checked items text from the values_combo (CheckableComboBox)
                if filt.get('values_combo'):
                    values_combo_widget = filt['values_combo']
                    checked_values = values_combo_widget.getCheckedItemsText()
                else:
                    self.logger.warning(f"Filter {idx} missing 'values_combo' widget during save.")

                # *** ADDED: Get the state of the 'apply as parameter' checkbox ***
                if filt.get('apply_as_param_checkbox'):
                    apply_param_state = filt['apply_as_param_checkbox'].isChecked()
                else:
                    self.logger.warning(f"Filter {idx} missing 'apply_as_param_checkbox' widget during save.")

            except RuntimeError:
                self.logger.error(f"RuntimeError accessing widgets for Filter {idx} during save. Skipping.")
                continue # Skip saving this filter if widgets are gone

            config[section] = {
                'field': field,
                'values': ','.join(checked_values), # Save as comma-separated string
                'apply_as_param': str(apply_param_state) # Save checkbox state as string ('True'/'False')
            }

        # --- Save Conditions (Keep as before) ---
        config['Conditions'] = {'count': str(len(self.conditions))}
        for idx, cond in enumerate(self.conditions):
            section = f'Condition_{idx}'
            column = "N/A"; type_val = "Equals"; value = ""; excluded = ""
            try:
                if cond.get('column_txt'): column = cond['column_txt'].currentText()
                if cond.get('type_choice'): type_val = cond['type_choice'].currentText()
                if cond.get('value_txt'): value = cond['value_txt'].text()
                if cond.get('excluded_views'): excluded = ','.join(cond['excluded_views'])
            except RuntimeError:
                self.logger.error(f"RuntimeError accessing widgets for Condition {idx} during save. Skipping.")
                continue
            config[section] = {
                'column': column,
                'type': type_val,
                'value': value,
                'excluded_views': excluded
            }

        # --- Save Parameters (Keep as before) ---
        config['Parameters'] = {'count': str(len(self.parameters))}
        for idx, param in enumerate(self.parameters):
            section = f'Parameter_{idx}'
            name = ""; value = ""
            try:
                if param.get('param_name_txt'): name = param['param_name_txt'].text()
                if param.get('param_value_txt'): value = param['param_value_txt'].text()
            except RuntimeError:
                self.logger.error(f"RuntimeError accessing widgets for Parameter {idx} during save. Skipping.")
                continue
            config[section] = {
                'name': name,
                'value': value
            }

        # --- Write to File (Keep as before) ---
        try:
            with open(file_path, 'w') as configfile:
                config.write(configfile)
            QMessageBox.information(self, "Success", f"Configuration saved successfully to:\n{file_path}");
            self.log_message_signal.emit(f"Configuration saved to {os.path.basename(file_path)}");
            self.logger.info(f"Configuration successfully saved to {file_path}.")
        except IOError as e:
            QMessageBox.critical(self, "Save Error", f"Could not write configuration file:\n{file_path}\n\n{e}");
            self.logger.error(f"Error writing configuration file {file_path}: {e}")
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"An unexpected error occurred while saving configuration:\n{e}");
            self.logger.error(f"Unexpected error saving configuration {file_path}: {e}", exc_info=True)


    def load_columns_for_conditions(self, column_names, specific_condition=None):
        """ Load columns into the dropdown for condition lines. """
        conditions_to_update = [specific_condition] if specific_condition else self.conditions
        self.logger.debug(f"Loading columns {column_names} for {len(conditions_to_update)} condition(s).")
        for cond_info in conditions_to_update:
             if cond_info and 'column_txt' in cond_info:
                try:
                    combo = cond_info['column_txt']; current_text = combo.currentText()
                    combo.clear(); combo.addItems(column_names); combo.setEnabled(True)
                    if current_text in column_names: combo.setCurrentText(current_text)
                    else: combo.setCurrentIndex(-1); combo.setPlaceholderText("Select Excel Field")
                except RuntimeError:
                    self.logger.warning("RuntimeError accessing condition combo box during column load.")


    def load_tableau_views(self):
        """Connects to Tableau Server, finds the workbook, and loads its view names."""
        print("Loading Tableau views...")
        self.log_message_signal.emit("Attempting to connect to Tableau Server and load views..."); self.logger.info("Attempting to load Tableau views.")
        tableau_server_url = self.server_url_text; token_name = self.token_name_text; token_secret = self.token_secret_text
        tableau_site_id = self.site_id_text or ""; workbook_name_to_find = self.workbook_name.text().strip()
        if not all([tableau_server_url, token_name, token_secret, workbook_name_to_find]):
             QMessageBox.warning(self, "Missing Info", "Please configure Server Connection first (use 'Server Config' button) and provide a Workbook Name."); self.log_message_signal.emit("Connection cancelled: Missing server/auth/workbook info."); self.logger.warning("Load views cancelled: Missing server/auth/workbook info."); return
        if not tableau_server_url.startswith(('http://', 'https://')): tableau_server_url = 'https://' + tableau_server_url; self.logger.info(f"Prepended 'https://' to server URL: {tableau_server_url}")
        server = None
        try:
            self.log_message_signal.emit(f"Authenticating with {tableau_server_url} (Site: '{tableau_site_id or 'Default'}')...")
            auth = PersonalAccessTokenAuth(token_name, token_secret, site_id=tableau_site_id); server = Server(tableau_server_url, use_server_version=True)
            server.add_http_options({'timeout': 120}); server.auth.sign_in_with_personal_access_token(auth)
            self.log_message_signal.emit("✔ Connected to Tableau Server.")
            self.logger.info("Tableau connection successful.")
            self.log_message_signal.emit(f"Searching for workbook '{workbook_name_to_find}'...")
            req_option = RequestOptions(pagesize=1); req_option.filter.add(Filter(RequestOptions.Field.Name, RequestOptions.Operator.Equals, workbook_name_to_find))
            all_matching_workbooks, _ = server.workbooks.get(req_option)
            if not all_matching_workbooks:
                self.logger.warning(f"Workbook '{workbook_name_to_find}' not found with exact name match. Trying case-insensitive search...")
                req_option_all = RequestOptions(pagesize=1000); all_workbooks, _ = server.workbooks.get(req_option_all)
                target_workbook = next((wb for wb in all_workbooks if wb.name.lower() == workbook_name_to_find.lower()), None)
                if not target_workbook: raise Exception(f"Workbook '{workbook_name_to_find}' not found on site '{tableau_site_id or 'Default'}'.")
            else: target_workbook = all_matching_workbooks[0]
            self.log_message_signal.emit(f"Found workbook '{target_workbook.name}' (ID: {target_workbook.id}). Populating views...")
            self.logger.info(f"Found workbook '{target_workbook.name}' (ID: {target_workbook.id}).")
            server.workbooks.populate_views(target_workbook)
            self.tableau_views = sorted([view.name for view in target_workbook.views])
            self.logger.info(f"Loaded {len(self.tableau_views)} views: {self.tableau_views}")
            QMessageBox.information(self, "Success", f"Successfully loaded {len(self.tableau_views)} views from workbook:\n'{target_workbook.name}'.")
            self.log_message_signal.emit(f"Successfully loaded {len(self.tableau_views)} views.")
        except Exception as e: error_msg = f"Failed to load views: {str(e)}"; QMessageBox.critical(self, "Connection/Load Error", error_msg); self.logger.error(error_msg, exc_info=True); self.log_message_signal.emit(f"Error loading views: {str(e)}"); self.tableau_views = []
        finally:
            if server and server.auth_token:
                try: server.auth.sign_out(); self.logger.info("Signed out from Tableau Server.")
                except Exception as sign_out_e: self.logger.error(f"Error during Tableau sign out: {sign_out_e}")

    def OnStart(self):
        """Initiates the export process after validation, folder check, and auto-collapsing logic."""
        print("Starting process...")

        # --- Auto-contract logic section ---
        # If the content is visible, hide it and update the button/state
        if hasattr(self, 'logic_content_widget') and self.logic_content_widget.isVisible():
            self.logger.info("Auto-contracting logic section for export.")
            self.logic_section_visible = False # Update state variable
            self.logic_content_widget.setVisible(False) # Hide the content widget
            if hasattr(self, 'logic_toggle_button'):
                self.logic_toggle_button.setChecked(False) # Update button check state
                self._update_logic_toggle_icon() # Update button icon
        # --- End auto-contract ---

        self.log_text.clear()
        self.log_message_signal.emit(f"--- Starting Export Process ({time.strftime('%Y-%m-%d %H:%M:%S')}) ---")
        self.logger.info("Starting export process initiated by user.")

        # --- Validations ---

        # 1. Output Folder Check & Creation
        output_folder = self.output_folder.text().strip()
        if not output_folder:
            QMessageBox.warning(self, "Input Error", "Output folder is required.")
            self.log_message_signal.emit("❌ Error: Output folder is empty. Process stopped.")
            self.logger.error("Start cancelled: Output folder empty.")
            return

        if not os.path.isdir(output_folder):
            reply = QMessageBox.question(self, "Create Folder?",
                                        f"The base output folder does not exist:\n{output_folder}\n\nDo you want to create it?",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                try:
                    os.makedirs(output_folder, exist_ok=True) # Use makedirs to create intermediate dirs if needed
                    self.logger.info(f"Created base output folder: {output_folder}")
                    self.log_message_signal.emit(f"ℹ Created output folder: {output_folder}")
                except OSError as e:
                    QMessageBox.critical(self, "Folder Error", f"Could not create base output folder:\n{e}")
                    self.log_message_signal.emit(f"❌ Error: Could not create base output folder. Process stopped.")
                    self.logger.error(f"Start cancelled: Failed to create base output folder: {e}")
                    return
            else:
                self.log_message_signal.emit("❌ Error: Output folder does not exist. Process stopped.")
                self.logger.error(f"Start cancelled: Output folder does not exist and user chose not to create it: {output_folder}")
                return
        elif not os.access(output_folder, os.W_OK): # Check write permissions if folder exists
            QMessageBox.warning(self, "Permission Error", f"Cannot write to the selected output folder:\n{output_folder}\n\nPlease check permissions.")
            self.log_message_signal.emit("❌ Error: Cannot write to output folder. Process stopped.")
            self.logger.error(f"Start cancelled: No write permissions for output folder: {output_folder}")
            return


        # 2. Mode-Specific Validations
        if self.mode_selection_dropdown.currentText() == "Automate for a list":
            if not self.excel_path:
                QMessageBox.warning(self, "Input Error", "Excel file is required for 'Automate for a list' mode.")
                self.log_message_signal.emit("❌ Error: Excel file missing. Process stopped.")
                self.logger.error("Start cancelled: Excel file missing for automate mode.")
                return
            if not os.path.exists(self.excel_path):
                QMessageBox.warning(self, "Input Error", f"Excel file not found:\n{self.excel_path}")
                self.log_message_signal.emit("❌ Error: Excel file not found. Process stopped.")
                self.logger.error(f"Start cancelled: Excel file not found: {self.excel_path}")
                return
            if not self.sheet_name.currentText():
                QMessageBox.warning(self, "Input Error", "Sheet name is required for 'Automate for a list' mode.")
                self.log_message_signal.emit("❌ Error: Sheet name missing. Process stopped.")
                self.logger.error("Start cancelled: Sheet name missing for automate mode.")
                return

            # Validate Filters, Conditions, Parameters (Ensure widgets exist before access)
            for i, filt in enumerate(self.filters):
                try:
                    if not filt['field_combo'].currentText():
                        QMessageBox.warning(self, "Input Error", f"Filter #{i+1}: Field cannot be empty.")
                        self.log_message_signal.emit(f"❌ Error: Filter #{i+1} field empty. Process stopped.")
                        self.logger.error(f"Start cancelled: Filter #{i+1} field empty.")
                        return
                except (RuntimeError, KeyError) as e:
                    self.logger.warning(f"Could not validate Filter #{i+1} due to error: {e}. Widgets might be deleted.")
                    # Decide if this is critical - returning False might be safer
                    # return False

            for i, cond in enumerate(self.conditions):
                try:
                    if not cond['column_txt'].currentText():
                        QMessageBox.warning(self, "Input Error", f"Condition #{i+1}: Field cannot be empty.")
                        self.log_message_signal.emit(f"❌ Error: Condition #{i+1} field empty. Process stopped.")
                        self.logger.error(f"Start cancelled: Condition #{i+1} field empty.")
                        return
                    cond_type = cond['type_choice'].currentText()
                    if cond_type not in ['Is Blank', 'Is Not Blank'] and not cond['value_txt'].text().strip():
                        QMessageBox.warning(self, "Input Error", f"Condition #{i+1}: Value cannot be empty for type '{cond_type}'.")
                        self.log_message_signal.emit(f"❌ Error: Condition #{i+1} value empty. Process stopped.")
                        self.logger.error(f"Start cancelled: Condition #{i+1} value empty.")
                        return
                except (RuntimeError, KeyError) as e:
                    self.logger.warning(f"Could not validate Condition #{i+1} due to error: {e}. Widgets might be deleted.")
                    # return False # Potentially safer

            for i, param in enumerate(self.parameters):
                try:
                    if not param['param_name_txt'].text().strip():
                        QMessageBox.warning(self, "Input Error", f"Parameter #{i+1}: Parameter Name cannot be empty.")
                        self.log_message_signal.emit(f"❌ Error: Parameter #{i+1} name empty. Process stopped.")
                        self.logger.error(f"Start cancelled: Parameter #{i+1} name empty.")
                        return
                except (RuntimeError, KeyError) as e:
                    self.logger.warning(f"Could not validate Parameter #{i+1} due to error: {e}. Widgets might be deleted.")
                    # return False # Potentially safer

        # 3. General Validation (Applies to both modes)
        if not all([self.server_url_text, self.token_name_text, self.token_secret_text, self.workbook_name.text().strip()]):
            QMessageBox.warning(self, "Input Error", "Server URL, Token Name, Token Secret (via Server Config button), and Workbook Name are required.")
            self.log_message_signal.emit("❌ Error: Missing server/auth/workbook info. Process stopped.")
            self.logger.error("Start cancelled: Missing server/auth/workbook info.")
            return
        
        # 4. Merge PDF specific validation
        if self.merge_pdfs_checkbox.isChecked() and not self.pdf_radio.isChecked():
            QMessageBox.warning(self, "Input Error", "PDF merging is only available for PDF export format.")
            self.log_message_signal.emit("❌ Error: PDF merging selected but format is not PDF. Process stopped.")
            self.logger.error("Start cancelled: PDF merging selected but format is not PDF.")
            return

        # --- All Validations Passed - Start the process ---
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Starting...")
        self.stop_event.clear()
        self.set_ui_enabled(False) # Disable UI
        self.worker = Thread(target=self.run_task, daemon=True)
        self.worker.start()
        self.logger.info("Worker thread started for export task.")
        
    def set_ui_enabled(self, enabled):
         """ Enables or disables relevant UI elements during processing. """
         self.logger.debug(f"Setting UI enabled state to: {enabled}")
         # Find group boxes by object name if possible, otherwise by title (less reliable)
         export_settings_box = self.findChild(QGroupBox, "Export Settings") # Assumes no objectName set
         combined_logic_box = getattr(self, 'combined_logic_box', None) # Use combined box reference

         if export_settings_box: export_settings_box.setEnabled(enabled)
         if combined_logic_box: combined_logic_box.setEnabled(enabled) # Enable/disable combined box

         # Enable/disable individual controls
         if hasattr(self, 'start_btn'): self.start_btn.setEnabled(enabled)
         if hasattr(self, 'load_btn'): self.load_btn.setEnabled(enabled)
         if hasattr(self, 'save_btn'): self.save_btn.setEnabled(enabled)
         if hasattr(self, 'server_config_btn'): self.server_config_btn.setEnabled(enabled)
         if hasattr(self, 'stop_btn'): self.stop_btn.setEnabled(not enabled)

         # Ensure progress bar and log remain enabled for feedback
         if hasattr(self, 'progress_bar'): self.progress_bar.setEnabled(True)
         if hasattr(self, 'log_text'): self.log_text.setEnabled(True)

         if hasattr(self, 'export_settings_toggle_button'): self.export_settings_toggle_button.setEnabled(enabled)

         # Mode dropdown should also be disabled during run
         if hasattr(self, 'mode_selection_dropdown'): self.mode_selection_dropdown.setEnabled(enabled)

         QApplication.processEvents() # Process events to reflect UI changes


    def run_task(self):
        """The main worker function executed in a separate thread."""
        server = None
        start_time = time.time()
        self.logger.info("run_task started.")
        task_successful = False # Flag to track success
        duration = 0 # Initialize duration

        try:
            # --- Get Settings ---
            tableau_server_url = self.server_url_text
            if not tableau_server_url.startswith(('http://', 'https://')):
                tableau_server_url = 'https://' + tableau_server_url
            token_name = self.token_name_text
            token_secret = self.token_secret_text
            tableau_site_id = self.site_id_text or ""
            excel_file = self.excel_path
            sheet_name = self.sheet_name.currentText()
            output_folder = self.output_folder.text().strip()
            workbook_name_to_find = self.workbook_name.text().strip()
            export_format = 'PDF' if self.pdf_radio.isChecked() else 'PNG'
            tableau_filter_field = self.tableau_filter_field_dropdown.currentText() if self.tableau_filter_field_dropdown.isEnabled() else ""
            current_mode = self.mode_selection_dropdown.currentText()
            merge_pdfs_enabled = self.merge_pdfs_checkbox.isChecked() # Get merge state

            self.logger.info(f"Export settings: Mode={current_mode}, Format={export_format}, Workbook='{workbook_name_to_find}', Output='{output_folder}', Merge PDFs={merge_pdfs_enabled}")
            self.log_message_signal.emit(f"▶ Connecting to Tableau Server: {tableau_server_url} (Site: '{tableau_site_id or 'Default'}')")

            # --- Tableau Connection ---
            auth = PersonalAccessTokenAuth(token_name, token_secret, site_id=tableau_site_id)
            server = Server(tableau_server_url, use_server_version=True)
            server.add_http_options({'timeout': 120}) # Increased timeout
            server.auth.sign_in_with_personal_access_token(auth)
            self.log_message_signal.emit("✔ Connected to Tableau Server.")
            self.logger.info("Tableau connection successful.")

            # --- Find Workbook ---
            self.log_message_signal.emit(f"🔍 Finding workbook '{workbook_name_to_find}'...")
            req_option = RequestOptions(pagesize=1)
            req_option.filter.add(Filter(RequestOptions.Field.Name, RequestOptions.Operator.Equals, workbook_name_to_find))
            all_matching_workbooks, _ = server.workbooks.get(req_option)

            if not all_matching_workbooks:
                self.logger.warning(f"Workbook '{workbook_name_to_find}' not found with exact name match. Trying case-insensitive search...")
                req_option_all = RequestOptions(pagesize=1000) # Get all workbooks for case-insensitive check
                all_workbooks, _ = server.workbooks.get(req_option_all)
                target_workbook = next((wb for wb in all_workbooks if wb.name.lower() == workbook_name_to_find.lower()), None)
                if not target_workbook:
                    raise Exception(f"Workbook '{workbook_name_to_find}' not found on site '{tableau_site_id or 'Default'}'. Check name and site.")
            else:
                target_workbook = all_matching_workbooks[0]

            self.log_message_signal.emit(f"✔ Found workbook '{target_workbook.name}' (ID: {target_workbook.id}).")
            self.logger.info(f"Found workbook '{target_workbook.name}' (ID: {target_workbook.id}).")

            # --- Get Views ---
            self.log_message_signal.emit(" Retrieving view list from workbook...")
            server.workbooks.populate_views(target_workbook)
            all_views_in_workbook = target_workbook.views
            self.log_message_signal.emit(f"✔ Found {len(all_views_in_workbook)} views in workbook.")
            self.logger.info(f"{len(all_views_in_workbook)} views found in workbook.")

            # --- Process Based on Mode ---
            if current_mode == "Automate for a list":
                self.log_message_signal.emit(f"📖 Loading data from Excel: '{os.path.basename(excel_file)}' Sheet: '{sheet_name}'")
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                self.log_message_signal.emit(f" Original list size: {len(df)} rows")
                self.logger.info(f"Loaded Excel data: {len(df)} rows.")

                filtered_df = self.apply_filters(df) # Apply UI filters
                total_items = len(filtered_df)
                self.log_message_signal.emit(f" Filtered list size: {total_items} rows")
                self.logger.info(f"List size after filtering: {total_items} rows.")

                if total_items == 0:
                    self.log_message_signal.emit("ℹ No items remaining after filtering. Nothing to export.")
                    self.logger.info("No items to process after filtering.")
                else:
                    # Dictionary to hold PDFs to merge for each lowest organization level
                    # Key: tuple (org1_val, org2_val) -> List of PDF paths
                    pdfs_to_merge_by_level = {} 

                    processed_items = 0
                    self.progress_bar.setFormat("Processing 0%")
                    for index, item_row in filtered_df.iterrows():
                        if self.stop_event.is_set():
                            self.log_message_signal.emit("⏹ Process stopped by user.")
                            self.logger.info("Stop event detected during item processing loop.")
                            break

                        processed_items += 1
                        self.log_message_signal.emit(f"--- Processing Item {processed_items}/{total_items} (Excel Index: {index}) ---")
                        self.logger.info(f"Processing item {processed_items}/{total_items} (Excel Index: {index}).")

                        # Determine Tableau Filter Value for this row
                        tableau_filter_value = None
                        if tableau_filter_field and tableau_filter_field in item_row:
                            tableau_filter_value = str(item_row[tableau_filter_field])
                            self.logger.debug(f"Tableau filter value for item: {tableau_filter_field} = {tableau_filter_value}")
                        elif tableau_filter_field:
                            self.logger.warning(f"Tableau filter field '{tableau_filter_field}' not found in Excel row {index}.")

                        # Determine Output Subfolder based on Organize By
                        organize_by1_col = self.organize_by_dropdown.currentText()
                        organize_by2_col = self.organize_by2_dropdown.currentText()
                        org1_val = str(item_row[organize_by1_col]) if organize_by1_col != "None" and organize_by1_col in item_row else ""
                        org2_val = str(item_row[organize_by2_col]) if organize_by2_col != "None" and organize_by2_col in item_row else ""
                        # Sanitize folder names
                        org1_val_clean = "".join(c for c in org1_val if c.isalnum() or c in (' ', '_', '-')).strip()
                        org2_val_clean = "".join(c for c in org2_val if c.isalnum() or c in (' ', '_', '-')).strip()

                        item_folder = output_folder
                        if org1_val_clean: item_folder = os.path.join(item_folder, org1_val_clean)
                        if org2_val_clean and not merge_pdfs_enabled: # Only create org2 folder if not merging at this level
                            item_folder = os.path.join(item_folder, org2_val_clean)

                        # Create directory if it doesn't exist
                        try:
                            if not os.path.exists(item_folder):
                                os.makedirs(item_folder, exist_ok=True)
                                self.logger.info(f"Created output directory: {item_folder}")
                        except OSError as e:
                            self.log_message_signal.emit(f"❌ Error creating directory '{item_folder}': {e}. Skipping item {processed_items}.")
                            self.logger.error(f"Error creating directory '{item_folder}': {e}. Skipping item.")
                            progress_percentage = int((processed_items / total_items) * 100)
                            self.progress_signal.emit(progress_percentage)
                            self.progress_bar.setFormat(f"Processing {progress_percentage}%")
                            continue # Skip to next item

                        # Process views for this specific item/row
                        # Pass the pdfs_to_merge_by_level dictionary
                        self.process_views_for_item(server, all_views_in_workbook, item_folder, tableau_filter_field, tableau_filter_value, item_row, export_format, merge_pdfs_enabled, pdfs_to_merge_by_level)

                        # Update progress
                        progress_percentage = int((processed_items / total_items) * 100)
                        self.progress_signal.emit(progress_percentage)
                        self.progress_bar.setFormat(f"Processing {progress_percentage}%")
                    
                    # --- After all items are processed, perform merges for "Automate for a list" mode ---
                    if merge_pdfs_enabled and export_format == "PDF":
                        self.log_message_signal.emit("Merging PDFs at lowest organization level...")
                        self.logger.info("Initiating PDF merge for 'Automate for a list' mode.")
                        for (org1_val, org2_val), pdf_paths in pdfs_to_merge_by_level.items():
                            if self.stop_event.is_set():
                                self.log_message_signal.emit("⏹ PDF merging stopped by user.")
                                self.logger.info("Stop event detected during PDF merging.")
                                break

                            if not pdf_paths: continue

                            # Determine the merge output path
                            merge_folder = output_folder
                            if org1_val: merge_folder = os.path.join(merge_folder, org1_val)
                            
                            # The merged file name will be based on the lowest organize by level
                            # If org2_val exists, use it. Otherwise, use org1_val. If neither, use a generic name.
                            if org2_val:
                                merged_filename = f"{org2_val}_merged.pdf"
                            elif org1_val:
                                merged_filename = f"{org1_val}_merged.pdf"
                            else:
                                merged_filename = "all_views_merged.pdf" # Fallback if no organize by selected

                            merged_filepath = os.path.join(merge_folder, merged_filename)
                            
                            self.log_message_signal.emit(f"  Merging {len(pdf_paths)} PDFs into: {os.path.basename(merged_filepath)}")
                            self.logger.info(f"Merging PDFs for ({org1_val}, {org2_val}) into {merged_filepath}")
                            self._merge_pdfs(pdf_paths, merged_filepath)
                            
                            # Clean up individual PDFs after merging
                            for pdf_path in pdf_paths:
                                try:
                                    os.remove(pdf_path)
                                    self.logger.debug(f"Removed temporary PDF after merge: {os.path.basename(pdf_path)}")
                                except Exception as e:
                                    self.logger.warning(f"Could not remove temporary PDF '{os.path.basename(pdf_path)}': {e}")
                        self.log_message_signal.emit("✔ PDF merging complete.")
                        self.logger.info("PDF merging for 'Automate for a list' mode finished.")

            else: # Export All Views Once Mode
                self.log_message_signal.emit("--- Exporting Selected Views Once ---")
                item_folder = output_folder # Export directly to output folder
                # Pass a list to collect PDF paths for "Export All Views Once" mode
                pdfs_for_single_merge = []
                self.export_selected_views_once(server, all_views_in_workbook, item_folder, export_format, merge_pdfs_enabled, pdfs_for_single_merge)

                # --- After all views are exported, perform single merge for "Export All Views Once" mode ---
                if merge_pdfs_enabled and export_format == "PDF" and pdfs_for_single_merge:
                    self.log_message_signal.emit("Merging all exported PDFs into one file...")
                    self.logger.info("Initiating single PDF merge for 'Export All Views Once' mode.")
                    
                    # Determine the merged file name
                    merged_filename = "All_Views_Merged.pdf"
                    merged_filepath = os.path.join(output_folder, merged_filename)

                    self.log_message_signal.emit(f"  Merging {len(pdfs_for_single_merge)} PDFs into: {os.path.basename(merged_filepath)}")
                    self.logger.info(f"Merging all PDFs into {merged_filepath}")
                    self._merge_pdfs(pdfs_for_single_merge, merged_filepath)

                    # Clean up individual PDFs after merging
                    for pdf_path in pdfs_for_single_merge:
                        try:
                            os.remove(pdf_path)
                            self.logger.debug(f"Removed temporary PDF after merge: {os.path.basename(pdf_path)}")
                        except Exception as e:
                            self.logger.warning(f"Could not remove temporary PDF '{os.path.basename(pdf_path)}': {e}")
                    self.log_message_signal.emit("✔ All PDFs merged into one file.")
                    self.logger.info("Single PDF merge for 'Export All Views Once' mode finished.")


            # --- Final Message ---
            if not self.stop_event.is_set():
                final_message = "✔ Task completed successfully."
                task_successful = True # Set flag on success
                self.log_message_signal.emit(final_message)
                self.progress_bar.setFormat("Completed")
                self.logger.info(final_message)
            else:
                final_message = "⏹ Task stopped by user."
                self.log_message_signal.emit(final_message)
                self.progress_bar.setFormat("Stopped")
                self.logger.info(final_message)

        except Exception as e:
            error_msg = f"❌ An error occurred during the task: {str(e)}"
            self.log_message_signal.emit(error_msg)
            self.logger.error(f"Error during run_task: {e}", exc_info=True)
            self.progress_bar.setFormat("Error")
            # Display the error in a message box as well
            try:
                QMessageBox.critical(self, "Runtime Error", f"An unexpected error occurred:\n{e}\n\nPlease check the log file (app.log) for details.")
            except RuntimeError:
                self.logger.error("Failed to show error QMessageBox, likely due to thread context.")

        finally:
            # --- Sign Out ---
            if server and server.auth_token:
                try:
                    server.auth.sign_out()
                    self.log_message_signal.emit("🔌 Disconnected from Tableau Server.")
                    self.logger.info("Disconnected from Tableau Server.")
                except Exception as sign_out_e:
                    self.logger.error(f"Error during Tableau sign out: {sign_out_e}")

            # --- Re-enable UI ---
            self.set_ui_enabled(True)
            self.logger.info("UI re-enabled.")
            end_time = time.time()
            duration = end_time - start_time # Calculate duration here

            # --- Format the duration ---
            if duration >= 60:
                minutes = int(duration // 60)
                seconds = int(duration % 60)
                formatted_duration = f"{minutes} m {seconds} s"
            else:
                formatted_duration = f"{duration:.2f}s" # Keep precision for shorter durations
            # --- End Formatting ---

            self.logger.info(f"run_task finished. Duration: {formatted_duration}.")
            self.log_message_signal.emit(f"--- Export Process Finished (Duration: {formatted_duration}) ---")

            # *** ADDED: Desktop Notification Logic ***
            try:
                # Import locally within the try block in case plyer is not installed
                from plyer import notification

                notif_title = ""
                notif_message = ""
                if self.stop_event.is_set():
                    notif_title = "Export Stopped"
                    notif_message = f"Tableau export process was stopped by the user."
                elif task_successful:
                    notif_title = "Export Complete"
                    notif_message = f"Tableau export finished successfully in {formatted_duration}."
                else: # Task finished but with an error
                    notif_title = "Export Error"
                    notif_message = "Tableau export finished with an error after {formatted_duration}. Check log."

                # Only show notification if there's a title (i.e., process actually ran)
                if notif_title:
                    # Attempt to find app icon for notification (optional)
                    app_icon_path = ""
                    icon_path_png = os.path.join(self.current_dir(), 'icons', 'icon.png')
                    icon_path_ico = os.path.join(self.current_dir(), 'icons', 'icon.ico')

                    if sys.platform == "win32" and os.path.exists(icon_path_ico):
                        app_icon_path = icon_path_ico # Use .ico for Windows if available
                    elif os.path.exists(icon_path_png):
                        app_icon_path = icon_path_png # Use .png otherwise (or for other platforms)

                    self.logger.info(f"Sending desktop notification: Title='{notif_title}', Message='{notif_message}'")
                    notification.notify(
                        title=notif_title,
                        message=notif_message,
                        app_name='Tableau PDF Export Tool',
                        app_icon=app_icon_path, # Path to icon (best effort)
                        timeout=5 # Notification stays for 5 seconds
                    )
            except ImportError:
                self.logger.error("Could not send desktop notification: 'plyer' library not found. Please install it (`pip install plyer`).")
                # Optionally inform the user via the log text area if plyer is missing
                self.log_message_signal.emit("⚠️ Desktop notifications require 'plyer'. Run: pip install plyer")
            except Exception as notif_e:
                # Catch potential errors from plyer itself (e.g., platform backend issues)
                self.logger.error(f"Failed to send desktop notification: {notif_e}", exc_info=True)
            # *** END ADDED NOTIFICATION LOGIC ***

    # --- MODIFY THIS FUNCTION ---
    def apply_filters(self, df):
        """Applies the filters defined in the UI's 'Filters' group to the DataFrame."""
        if not self.filters:
            self.log_message_signal.emit("ℹ No Excel filters defined, using original list.")
            self.logger.info("No Excel filters applied.")
            return df

        self.log_message_signal.emit("Applying Excel list filters...")
        self.logger.info("Applying Excel list filters.")
        filtered_df = df.copy()

        for i, filt in enumerate(self.filters):
            try:
                field = filt['field_combo'].currentText()
                values_combo = filt['values_combo']
                selected_values = values_combo.getCheckedItemsText()
                # --- End getting selected values ---

                self.logger.debug(f"Applying Filter #{i+1}: Field='{field}', Selected Values='{selected_values}'")

                if not field:
                    self.log_message_signal.emit(f"⚠️ Skipping Filter #{i+1}: No field selected.")
                    self.logger.warning(f"Skipping Filter #{i+1}: No field selected.")
                    continue
                if field not in filtered_df.columns:
                    self.log_message_signal.emit(f"⚠️ Skipping Filter #{i+1}: Field '{field}' not found in Excel sheet.")
                    self.logger.warning(f"Skipping Filter #{i+1}: Field '{field}' not found in Excel sheet.")
                    continue
                if not selected_values:
                    # If nothing is checked (or visible and checked), filter removes all rows matching this field.
                    self.log_message_signal.emit(f"ℹ Filter #{i+1} (Field='{field}'): No values checked/visible.")
                    self.logger.info(f"Filter #{i+1} for field '{field}': No values checked/visible.")
                    # Apply filter even if empty, to remove all rows unless skipped
                    # If you want to SKIP the filter if nothing is checked, uncomment below:
                    # continue

                original_len = len(filtered_df)
                # Ensure comparison is done with strings if selected_values are strings
                mask = filtered_df[field].astype(str).isin([str(v) for v in selected_values])
                filtered_df = filtered_df[mask]
                rows_removed = original_len - len(filtered_df)
                self.log_message_signal.emit(f" Filter #{i+1} (Field='{field}', {len(selected_values)} values checked) applied. Rows remaining: {len(filtered_df)} (-{rows_removed})")
                self.logger.debug(f"Filter #{i+1} applied. Rows remaining: {len(filtered_df)}")

            except RuntimeError:
                self.logger.error(f"RuntimeError applying Filter #{i+1}. Widgets might be deleted. Skipping filter.")
                continue
            except KeyError as e:
                self.log_message_signal.emit(f"❌ Error applying Filter #{i+1} (Field='{field}'): Column key error {e}. Skipping filter.")
                self.logger.error(f"KeyError applying Filter #{i+1} (Field='{field}'): {e}")
                continue
            except Exception as e:
                error_details = f"Filter #{i+1} (Field='{field}')"
                self.log_message_signal.emit(f"❌ Error applying {error_details}: {e}. Skipping filter.")
                self.logger.error(f"Error applying {error_details}: {e}", exc_info=True)

            if filtered_df.empty:
                self.log_message_signal.emit("ℹ Filter resulted in 0 rows. Stopping further filtering.")
                self.logger.info("DataFrame empty after applying filter, stopping further filtering.")
                break
        return filtered_df


    # --- Inside PDFExportApp class ---

# REPLACE this function
    def process_views_for_item(self, server, all_views_in_workbook, item_folder_base, tableau_filter_field, tableau_filter_value, item_row, export_format, merge_pdfs_enabled, pdfs_to_merge_by_level):
        """Processes and exports the relevant views for a single item (row) from the filtered Excel list."""
        try:
            self.logger.info(f"Processing views for item. Output folder: {item_folder_base}")

            # --- Determine Excluded Views for this Item (Keep Condition logic as before) ---
            views_to_exclude_for_this_item = set(self.excluded_views_for_export) # Start with global exclusions
            # ... (Keep the existing condition evaluation loop exactly as it is) ...
            for i, condition in enumerate(self.conditions):
                try:
                    cond_col = condition['column_txt'].currentText()
                    cond_type = condition['type_choice'].currentText()
                    cond_val_str_ui = condition['value_txt'].text().strip()
                    cond_excluded_views = condition['excluded_views']
                    if not cond_col or cond_col not in item_row.index: continue
                    actual_value_raw = item_row[cond_col]
                    actual_value_str_stripped = str(actual_value_raw).strip()
                    condition_met = False
                    comparison_type_logged = "N/A"
                    # --- (Keep existing condition evaluation logic here) ---
                    if cond_type == 'Is Blank':
                        condition_met = pd.isna(actual_value_raw) or actual_value_str_stripped == ""
                        comparison_type_logged = "Is Blank"
                    elif cond_type == 'Is Not Blank':
                        condition_met = not pd.isna(actual_value_raw) and actual_value_str_stripped != ""
                        comparison_type_logged = "Is Not Blank"
                    elif cond_type in ['Equals', 'Not Equals']:
                        try:
                            actual_num = float(actual_value_raw); cond_num = float(cond_val_str_ui)
                            if cond_type == 'Equals': condition_met = actual_num == cond_num
                            else: condition_met = actual_num != cond_num
                            comparison_type_logged = "Numeric"
                        except (ValueError, TypeError):
                            actual_check_str = actual_value_str_stripped.lower(); cond_check_str = cond_val_str_ui.lower()
                            if cond_type == 'Equals': condition_met = actual_check_str == cond_check_str
                            else: condition_met = actual_check_str != cond_check_str
                            comparison_type_logged = "String (Case-Insensitive)"
                    elif cond_type in ['Greater Than', 'Less Than']:
                        try:
                            actual_num = pd.to_numeric(actual_value_raw, errors='coerce'); cond_num = pd.to_numeric(cond_val_str_ui, errors='coerce')
                            if pd.notna(actual_num) and pd.notna(cond_num):
                                if cond_type == 'Greater Than': condition_met = actual_num > cond_num
                                elif cond_type == 'Less Than': condition_met = actual_num < cond_num
                                comparison_type_logged = "Numeric Comparison (GT/LT)"
                            else: condition_met = False; comparison_type_logged = "Numeric Comparison Failed (GT/LT)"
                        except Exception as num_e: condition_met = False; comparison_type_logged = "Numeric Comparison Error (GT/LT)"; self.logger.error(f"Condition #{i+1}: Error during numeric comparison: {num_e}")
                    # --- (End of condition evaluation logic) ---
                    log_detail = (f"Condition #{i+1}: Field='{cond_col}', Type='{cond_type}', UI_Value='{cond_val_str_ui}', Actual='{actual_value_str_stripped}', Comparison='{comparison_type_logged}', Met={condition_met}")
                    if condition_met:
                        self.logger.info(f"{log_detail} -> Excluding views: {cond_excluded_views}")
                        views_to_exclude_for_this_item.update(cond_excluded_views)
                    else:
                        self.logger.debug(log_detail)
                except RuntimeError: self.logger.error(f"RuntimeError accessing widgets for Condition #{i+1}. Skipping condition."); continue
                except KeyError as e: self.logger.error(f"KeyError accessing data for Condition #{i+1}: {e}. Skipping condition."); continue
                except Exception as e: self.logger.error(f"Unexpected error evaluating Condition #{i+1}: {e}", exc_info=True); continue
            # --- End Condition Evaluation ---

            # --- Filter Views to Export (Keep as before) ---
            # ... (Keep the view filtering logic exactly as it is) ...
            included_views_for_item = [view for view in all_views_in_workbook if view.name not in views_to_exclude_for_this_item]
            if not included_views_for_item:
                message = "ℹ No views left to export for this item after applying conditions."
                self.log_message_signal.emit(message); self.logger.info(message); return
            else:
                included_names = [v.name for v in included_views_for_item]
                self.logger.info(f"Views to be exported for this item ({len(included_names)}): {included_names}")

            item_folder_final = item_folder_base
            try:
                os.makedirs(item_folder_final, exist_ok=True)
            except OSError as e:
                self.log_message_signal.emit(f"❌ Error ensuring directory '{item_folder_final}' exists: {e}. Skipping item.")
                self.logger.error(f"Error ensuring directory '{item_folder_final}' exists: {e}. Skipping item.", exc_info=True)
                return

            # --- Prepare Export Options ---
            if export_format == "PDF":
                export_options = PDFRequestOptions(page_type=PDFRequestOptions.PageType.Unspecified, maxage=0)
            else:
                export_options = ImageRequestOptions(imageresolution=ImageRequestOptions.Resolution.High, maxage=0)

            # Add Tableau Filter (if applicable) (Keep as before)
            # ... (Keep the Tableau filter logic exactly as it is) ...
            if tableau_filter_field and tableau_filter_value is not None:
                self.logger.debug(f"Applying Tableau filter: {tableau_filter_field} = {tableau_filter_value}")
                export_options.vf(tableau_filter_field, tableau_filter_value)
            elif tableau_filter_field:
                self.logger.warning(f"Tableau filter field '{tableau_filter_field}' specified but no value found/provided for this item.")


            # --- Apply Parameters from Filters (DEBUG LOGGING ADDED) ---
            applied_param_filter_fields = set() # Track fields already applied as params from filters
            self.logger.debug("--- Checking Filters for Parameter Application ---") # ADDED
            for i, filt in enumerate(self.filters):
                try:
                    self.logger.debug(f"Filter #{i+1}: Checking...") # ADDED
                    checkbox = filt.get('apply_as_param_checkbox')
                    field_combo = filt.get('field_combo')
                    values_combo = filt.get('values_combo')

                    # ADDED Debug checks
                    if not checkbox: self.logger.debug(f"Filter #{i+1}: Checkbox widget not found."); continue
                    if not field_combo: self.logger.debug(f"Filter #{i+1}: Field combo widget not found."); continue
                    if not values_combo: self.logger.debug(f"Filter #{i+1}: Values combo widget not found."); continue

                    is_checked = checkbox.isChecked() # Get state
                    self.logger.debug(f"Filter #{i+1}: 'Apply as Param' checkbox state: {is_checked}") # ADDED

                    if is_checked:
                        filter_field_name = field_combo.currentText()
                        selected_values = values_combo.getCheckedItemsText() # Get currently checked values
                        self.logger.debug(f"Filter #{i+1}: Field='{filter_field_name}', Selected Values='{selected_values}'") # ADDED

                        if not filter_field_name:
                            self.logger.warning(f"Filter #{i+1}: Cannot apply as parameter - field name is empty.")
                            continue

                        if not selected_values:
                            self.logger.info(f"Filter #{i+1} ('{filter_field_name}'): 'Apply as Param' checked, but no values selected. Skipping parameter application.")
                            continue

                        # *** Use the FIRST selected value as the parameter value ***
                        param_value = selected_values[0]
                        param_name = filter_field_name + "Param"

                        # Apply the parameter override using export_options.vf()
                        self.logger.debug(f"Filter #{i+1}: Attempting to apply vf('{param_name}', '{param_value}')") # ADDED
                        export_options.vf(param_name, param_value)
                        applied_param_filter_fields.add(param_name) # Mark this field as handled
                        # This INFO log should appear if the vf() call is reached
                        ##self.log_message_signal.emit(f"ℹ Applied Parameter from Filter: {param_name} = {param_value}") # INFO level for UI log
                        self.logger.info(f"Applied Parameter from Filter #{i+1}: Name='{param_name}', Value='{param_value}' (used first selected filter value).")

                except RuntimeError:
                    self.logger.error(f"RuntimeError accessing widgets for Filter #{i+1} during parameter application. Skipping.")
                    continue
                except Exception as e:
                    self.logger.error(f"Unexpected error processing Filter #{i+1} for parameter application: {e}", exc_info=True)
                    continue
            self.logger.debug("--- Finished Checking Filters for Parameter Application ---") # ADDED
            # --- End Apply Parameters from Filters ---

            # --- Apply Explicit Parameters (MODIFIED to prevent overwrite) ---
            # ... (Keep the Explicit Parameter logic exactly as it is, including the check for applied_param_filter_fields) ...
            self.logger.debug("--- Checking Explicit Parameters ---") # ADDED
            for i, param in enumerate(self.parameters):
                try:
                    param_name = param['param_name_txt'].text().strip()
                    param_value_source = param['param_value_txt'].text().strip()

                    if not param_name:
                        self.logger.warning(f"Skipping Explicit Parameter #{i+1}: Name is empty.")
                        continue

                    # *** CHECK: Only apply if not already handled by a filter checkbox ***
                    if param_name in applied_param_filter_fields:
                        self.logger.warning(f"Skipping Explicit Parameter #{i+1} ('{param_name}'): Parameter was already set by a Filter with 'Apply as Param' checked.")
                        continue

                    final_param_value = param_value_source # Default to static value
                    if param_value_source in item_row.index:
                        cell_value = item_row.get(param_value_source)
                        if pd.isna(cell_value): final_param_value = ""
                        else: final_param_value = str(cell_value)
                        self.logger.debug(f"Explicit Parameter '{param_name}': Using value from column '{param_value_source}' -> '{final_param_value}'.")
                    else:
                        self.logger.debug(f"Explicit Parameter '{param_name}': Using static value '{final_param_value}'.")

                    # Apply the parameter override
                    self.logger.debug(f"Explicit Param #{i+1}: Attempting to apply vf('{param_name}', '{final_param_value}')") # ADDED
                    export_options.vf(param_name, final_param_value)
                    self.log_message_signal.emit(f"ℹ Applied Explicit Parameter: {param_name} = {final_param_value}") # INFO level for UI log
                    self.logger.info(f"Applied Explicit Parameter #{i+1}: Name='{param_name}', Value='{final_param_value}'.") # ADDED Log

                except RuntimeError:
                    self.logger.error(f"RuntimeError accessing widgets for Explicit Parameter #{i+1}. Skipping parameter.")
                    continue
                except KeyError as e:
                    self.logger.error(f"KeyError accessing data for Explicit Parameter #{i+1}: {e}. Skipping parameter.")
                    continue
                except Exception as e:
                    self.logger.error(f"Unexpected error processing Explicit Parameter #{i+1}: {e}", exc_info=True)
                    continue
            self.logger.debug("--- Finished Checking Explicit Parameters ---") # ADDED
            # --- End Apply Explicit Parameters ---


            # --- Export Each Included View (Keep remaining logic as before) ---
            # ... (Keep the view export loop exactly as it is) ...
            file_naming_col = self.file_naming_option.currentText()
            numbering_enabled = self.numbering_checkbox.isChecked()
            counter = 1
            organize_by1_col = self.organize_by_dropdown.currentText(); organize_by2_col = self.organize_by2_dropdown.currentText()
            org1_val = str(item_row.get(organize_by1_col, "")).strip() if organize_by1_col != "None" else ""
            org2_val = str(item_row.get(organize_by2_col, "")).strip() if organize_by2_col != "None" else ""
            org1_val_clean = "".join(c for c in org1_val if c.isalnum() or c in (' ', '_', '-')).strip()
            org2_val_clean = "".join(c for c in org2_val if c.isalnum() or c in (' ', '_', '-')).strip()

            for view in included_views_for_item:
                if self.stop_event.is_set(): self.logger.info("Stop event detected during view export loop for item."); break
                prefix = f"{counter:02d}_" if numbering_enabled else ""
                base_name = ""; use_view_name = False
                try:
                    if file_naming_col == "By view": use_view_name = True
                    elif file_naming_col not in item_row.index: use_view_name = True; self.logger.warning(f"File naming column '{file_naming_col}' not found, using view name '{view.name}'.")
                    else:
                        cell_value = item_row.get(file_naming_col)
                        if pd.isna(cell_value) or str(cell_value).strip() == "": use_view_name = True; self.logger.warning(f"File naming column '{file_naming_col}' empty, using view name '{view.name}'.")
                        else: base_name = str(cell_value).strip()
                    if use_view_name: base_name = view.name
                    base_name_clean = "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in base_name).strip()
                    max_len = 100
                    if len(base_name_clean) > max_len: base_name_clean = base_name_clean[:max_len]
                    if not base_name_clean: base_name_clean = f"view_{view.id}_invalid_name"
                except Exception as name_e: base_name_clean = f"view_{view.id}_naming_error"; self.logger.error(f"Error determining base filename for view '{view.name}': {name_e}. Using fallback.", exc_info=True)

                file_name_final = f"{prefix}{base_name_clean}.{export_format.lower()}"
                
                # Determine the actual export path based on merge setting
                current_export_folder = item_folder_final
                if merge_pdfs_enabled and export_format == "PDF":
                    # If merging, export to a temporary directory inside HIDDEN_TEMP_DIR
                    # This ensures individual PDFs don't clutter the main output folder
                    current_export_folder = os.path.join(HIDDEN_TEMP_DIR, "temp_pdfs_for_merge")
                    os.makedirs(current_export_folder, exist_ok=True) # Ensure temp dir exists
                    # Append a unique identifier to the filename to avoid conflicts if views have same base_name
                    # This is crucial when merging, as all individual PDFs need unique names before merge
                    file_name_final = f"{prefix}{base_name_clean}_{view.id}.{export_format.lower()}"
                    self.logger.debug(f"Exporting to temp folder for merge: {current_export_folder}/{file_name_final}")

                export_file_path = os.path.join(current_export_folder, file_name_final)

                self.export_single_view_with_retry(server, view, export_options, export_file_path, export_format, org1_val_clean, org2_val_clean, base_name_clean)
                
                # If merging PDFs, add the path to the list for the corresponding merge level
                if merge_pdfs_enabled and export_format == "PDF":
                    # Determine the key for the merge dictionary
                    # If org2 is selected, merge at org2 level. Else, org1 level. Else, root.
                    merge_key = (org1_val_clean, org2_val_clean) # Tuple (org1, org2)
                    if merge_key not in pdfs_to_merge_by_level:
                        pdfs_to_merge_by_level[merge_key] = []
                    pdfs_to_merge_by_level[merge_key].append(export_file_path)

                if numbering_enabled: counter += 1
            # --- End Exporting Views for Item ---

            self.logger.info(f"Finished processing views for item (Folder: {os.path.basename(item_folder_final)}).")

        except Exception as e:
            error_msg = f"❌ Error processing views for item: {str(e)}"
            self.log_message_signal.emit(error_msg)
            self.logger.error(f"Error in process_views_for_item: {e}", exc_info=True)
            
    def export_selected_views_once(self, server, all_views_in_workbook, output_folder, export_format, merge_pdfs_enabled, pdfs_for_single_merge):
        """Exports views selected in the UI directly, without using Excel data ('Export All Views Once' mode)."""
        try:
            self.log_message_signal.emit(f"--- Exporting Selected Views ({export_format}) ---")
            self.logger.info(f"Starting 'Export All Views Once' mode. Format={export_format}, Output='{output_folder}', Merge PDFs={merge_pdfs_enabled}")

            # Filter views based on global exclusion list
            selected_views_to_export = [view for view in all_views_in_workbook if view.name not in self.excluded_views_for_export]
            total_views = len(selected_views_to_export)

            if total_views == 0:
                message = "ℹ No views selected or remaining after global exclusion. Nothing to export."
                self.log_message_signal.emit(message)
                self.logger.info(message)
                self.progress_bar.setFormat("No views selected")
                return

            self.log_message_signal.emit(f"Attempting to export {total_views} selected view(s) to: {output_folder}")
            self.logger.info(f"Attempting to export {total_views} views: {[v.name for v in selected_views_to_export]}")
            processed_views = 0
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("Exporting 0%")

            numbering_enabled = self.numbering_checkbox.isChecked()
            counter = 1

            # Initialize export_options once before the loop
            if export_format == "PDF":
                export_options = PDFRequestOptions(page_type=PDFRequestOptions.PageType.Unspecified, maxage=0)
            else: # PNG
                export_options = ImageRequestOptions(imageresolution=ImageRequestOptions.Resolution.High, maxage=0)

            for view in selected_views_to_export:
                if self.stop_event.is_set():
                    self.log_message_signal.emit("⏹ Process stopped by user.")
                    self.logger.info("Stop event detected during 'Export All Views Once'.")
                    break

                # Determine filename (using view name)
                prefix = f"{counter:02d}_" if numbering_enabled else ""
                base_name_clean = "" # Define before try
                try:
                    # Sanitize base name directly from view name
                    base_name = view.name
                    base_name_clean = "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in base_name).strip()
                    max_len = 100
                    if len(base_name_clean) > max_len:
                        # ... (truncation logic - keep as before) ...
                        base_name_clean = base_name_clean[:max_len] # Simple truncate fallback
                    if not base_name_clean:
                        base_name_clean = f"view_{view.id}_invalid_name"
                except Exception as name_e:
                    self.logger.error(f"Error determining base filename for view '{view.name}': {name_e}. Using fallback.", exc_info=True)
                    base_name_clean = f"view_{view.id}_naming_error"

                file_name_final = f"{prefix}{base_name_clean}.{export_format.lower()}"
                
                # Determine the actual export path based on merge setting
                current_export_folder = output_folder
                if merge_pdfs_enabled and export_format == "PDF":
                    # If merging, export to a temporary directory inside HIDDEN_TEMP_DIR
                    current_export_folder = os.path.join(HIDDEN_TEMP_DIR, "temp_pdfs_for_single_merge")
                    os.makedirs(current_export_folder, exist_ok=True) # Ensure temp dir exists
                    # Append a unique identifier to the filename to avoid conflicts if views have same base_name
                    file_name_final = f"{prefix}{base_name_clean}_{view.id}.{export_format.lower()}"
                    self.logger.debug(f"Exporting to temp folder for single merge: {current_export_folder}/{file_name_final}")

                export_file_path = os.path.join(current_export_folder, file_name_final)

                # Export the single view - Pass None for org parts, pass cleaned base name
                self.export_single_view_with_retry(
                    server, view, export_options, export_file_path, export_format,
                    org1_val_clean=None, org2_val_clean=None, base_name_clean=base_name_clean
                )
                
                # If merging PDFs, add the path to the list
                if merge_pdfs_enabled and export_format == "PDF":
                    pdfs_for_single_merge.append(export_file_path)

                processed_views += 1
                progress_percentage = int((processed_views / total_views) * 100)
                self.progress_signal.emit(progress_percentage)
                self.progress_bar.setFormat(f"Exporting {progress_percentage}%")

                if numbering_enabled: counter += 1

            if not self.stop_event.is_set():
                self.log_message_signal.emit("✔ Finished exporting selected views.")
                self.logger.info("Finished 'Export All Views Once'.")

        except Exception as e:
            error_message = f"❌ An error occurred during 'Export All Views Once': {str(e)}"
            # print(error_message) # Less verbose print
            self.log_message_signal.emit(error_message)
            self.logger.error(f"Error in export_selected_views_once: {e}", exc_info=True)
            self.progress_bar.setFormat("Error")
            QMessageBox.critical(self, "Export Error", f"An error occurred during export:\n{e}")

    def _trim_pdf_whitespace(self, pdf_path):
        """
        Trims whitespace from the BOTTOM ONLY of a PDF's first page using PyMuPDF (fitz),
        leaving specified padding below the detected content.
        Saves to a temporary file first, then replaces the original.

        Args:
            pdf_path (str): The path to the PDF file to trim.
        """

        self.logger.info(f"Attempting to trim bottom whitespace for {pdf_path} using PyMuPDF.")
        temp_pdf_path = pdf_path + ".trimmed_tmp.pdf"
        doc = None

        try:
            # Ensure fitz is available (rely on top-level import)
            if 'fitz' not in sys.modules:
                 raise ImportError("PyMuPDF (fitz) is not imported. Please ensure 'import fitz' is at the top of the script.")

            try:
                doc = fitz.open(pdf_path)
            except Exception as open_error:
                 self.log_message_signal.emit(f"  Error opening PDF for trimming: {os.path.basename(pdf_path)}")
                 self.logger.error(f"Failed to open PDF '{os.path.basename(pdf_path)}' with PyMuPDF: {open_error}")
                 return # Cannot proceed if file can't be opened

            if doc.page_count == 0:
                self.logger.warning(f"PDF has no pages, cannot trim: {pdf_path}")
                if doc: doc.close()
                return

            page = doc[0] # Process the first page
            content_bbox = fitz.Rect() # Initialize empty bounding box

            # --- Find content boundary (only need bottom edge, but find full box for robustness) ---
            blocks = page.get_text("blocks")
            if blocks:
                for b in blocks: content_bbox.include_rect(fitz.Rect(b[:4]))
                self.logger.debug(f"Found text blocks for trimming {os.path.basename(pdf_path)}. Content bbox: {content_bbox}")
            else:
                # If no text, try drawings
                self.logger.debug(f"No text blocks found in {os.path.basename(pdf_path)}, checking drawings...")
                drawings = page.get_drawings()
                if drawings:
                    for d in drawings:
                         # Consider rectangles that seem like content areas
                         if d['rect'].width > 1 and d['rect'].height > 1: content_bbox.include_rect(d['rect'])
                    self.logger.debug(f"Found drawings for trimming {os.path.basename(pdf_path)}. Content bbox: {content_bbox}")
                else:
                    # If still no content found, cannot determine bottom edge
                    self.logger.warning(f"No text blocks or drawings found on page 1 of {os.path.basename(pdf_path)}. Cannot determine bottom boundary. Skipping trim.")
                    if doc: doc.close()
                    return

            # Check if a valid bounding box (specifically with height) was found
            if content_bbox.is_empty or content_bbox.height <= 0:
                self.logger.warning(f"Could not determine valid content bottom boundary for {os.path.basename(pdf_path)}. Skipping trim.")
                if doc: doc.close()
                return
            # ---

            # --- Define Bottom Padding (Points) ---
            # <<< --- ADJUST THIS VALUE AS NEEDED --- >>> (72 points = 1 inch)
            bottom_padding = 20.0
            # ------------------------------------

            original_rect = page.rect # Get the original page dimensions (MediaBox or CropBox)

            # Calculate the new bottom edge (y1 in fitz Rect) based on content bottom + padding
            # Ensure it doesn't exceed the original page bottom
            new_bottom_edge = min(original_rect.y1, content_bbox.y1 + bottom_padding)

            # Create the final crop box using original left, top, right and ONLY the new bottom
            final_crop_box = fitz.Rect(
                original_rect.x0,       # Original Left
                original_rect.y0,       # Original Top
                original_rect.x1,       # Original Right
                new_bottom_edge         # New Bottom edge
            )

            # --- Check if trimming is actually needed and the box is valid ---
            if final_crop_box.height <= 0:
                self.logger.warning(f"Calculated final crop box has invalid height ({final_crop_box}) for {os.path.basename(pdf_path)}. Skipping trim.")
                if doc: doc.close()
                return
            # Use a small tolerance for floating point comparison
            elif abs(final_crop_box.height - original_rect.height) < 0.1:
                self.logger.info(f"Calculated final height is effectively the same as original for {os.path.basename(pdf_path)}. No bottom trim needed.")
                if doc: doc.close()
                return # No need to save if no change
            # ---

            padding_info = f"Bottom Pad:{bottom_padding}pt"
            self.logger.info(f"Trimming Bottom PDF {os.path.basename(pdf_path)}: Original Rect={original_rect}, Content Bottom={content_bbox.y1:.2f}, Final Crop Box={final_crop_box} (Padding: {padding_info})")

            # Apply the new crop box to the page
            page.set_cropbox(final_crop_box)

            # --- Save to TEMPORARY file first ---
            temp_pdf_path = pdf_path + ".trimmed_tmp.pdf"
            self.logger.debug(f"Saving bottom-trimmed PDF to temporary file: {temp_pdf_path}")
            # Use garbage collection options for potentially smaller file size
            doc.save(temp_pdf_path, garbage=4, deflate=True, linear=True)
            self.logger.debug(f"Successfully saved temporary file: {temp_pdf_path}")
            # ---

            # --- Close the document handle BEFORE replacing the file ---
            doc.close()
            doc = None # Prevent closing again in finally if successful
            self.logger.debug(f"Closed original document handle for {pdf_path}")
            # ---

            # --- Replace original file with temporary file ---
            try:
                os.replace(temp_pdf_path, pdf_path)
                self.log_message_signal.emit(f"  PDF Bottom Trim applied ({padding_info}): {os.path.basename(pdf_path)}")
                self.logger.info(f"Successfully replaced original with bottom-trimmed PDF: {os.path.basename(pdf_path)}")
            except Exception as replace_error:
                 # Log error, temp file cleanup happens in finally
                 self.log_message_signal.emit(f"  Error replacing original PDF with trimmed version: {os.path.basename(pdf_path)}")
                 self.logger.error(f"Error replacing '{pdf_path}' with '{temp_pdf_path}': {replace_error}")
            # ---

        except ImportError as e:
            self.log_message_signal.emit(f"  Error trimming PDF: PyMuPDF library not found. Please run 'pip install PyMuPDF'.")
            self.logger.error(f"Failed to trim PDF '{os.path.basename(pdf_path)}'. PyMuPDF not installed? Error: {e}")
        except Exception as e:
            # Catch any other errors during the process
            error_details = traceback.format_exc()
            self.log_message_signal.emit(f"  Error during PDF bottom trim for {os.path.basename(pdf_path)}: {type(e).__name__}")
            self.logger.error(f"Error trimming PDF bottom '{os.path.basename(pdf_path)}': {e}\n{error_details}")
        finally:
            # Ensure the document is closed if it's still open (e.g., if error happened before close)
            if doc is not None:
                try:
                    doc.close()
                    self.logger.debug(f"Closed document handle in finally block for {pdf_path}")
                except Exception as close_e:
                    self.logger.error(f"Error closing PDF document '{os.path.basename(pdf_path)}' in finally block: {close_e}")

            # --- Clean up the temporary file ---
            if os.path.exists(temp_pdf_path):
                try:
                    os.remove(temp_pdf_path)
                    self.logger.debug(f"Removed temporary file: {temp_pdf_path}")
                except Exception as remove_error:
                    self.logger.error(f"Error removing temporary file '{temp_pdf_path}': {remove_error}")
            # ---

    def _trim_png_whitespace(self, image_path):
        """
        Trims whitespace from the BOTTOM ONLY of a PNG image using Pillow
        by scanning pixels, leaving specified padding below the detected content.
        Relies on top-level PIL imports.

        Args:
            image_path (str): The path to the PNG file to trim.
        """
        # Ensure necessary Pillow modules are imported at the top of the script file:
        # from PIL import Image
        # Also ensure standard libraries are imported: import os, sys, traceback

        self.logger.info(f"Attempting to trim bottom whitespace for {image_path} using Pillow (Pixel Scan method).")

        try:
            # --- Define Bottom Padding (Pixels) ---
            # <<< --- ADJUST THIS VALUE AS NEEDED --- >>>
            bottom_padding = 20
            # -------------------------------------------

            # Use the globally imported Image object
            im = Image.open(image_path)
            width, height = im.size
            self.logger.debug(f"Image mode for {os.path.basename(image_path)}: {im.mode}")

            # --- Determine Background Color & Check Function ---
            bg_pixel = im.getpixel((0, 0))
            self.logger.debug(f"Sampled background pixel (0,0): {bg_pixel} (Mode: {im.mode})")

            im_for_scan = im # Assume we scan the original unless converted
            bg_color_rgb_default = (255, 255, 255)
            bg_color_l_default = 255

            if im.mode == "RGBA":
                bg_color_to_match = bg_color_rgb_default
                self.logger.debug(f"Assuming RGBA background is transparent OR RGB={bg_color_to_match}")
                def is_background(pixel):
                    return pixel[3] == 0 or pixel[:3] == bg_color_to_match
            elif im.mode == "RGB":
                bg_color_to_match = bg_pixel
                if bg_color_to_match != bg_color_rgb_default:
                     self.logger.debug(f"Corner pixel wasn't white, assuming {bg_color_rgb_default} background for RGB.")
                     bg_color_to_match = bg_color_rgb_default
                def is_background(pixel):
                    return pixel == bg_color_to_match
            elif im.mode == "L":
                bg_color_to_match = bg_pixel
                if bg_color_to_match != bg_color_l_default:
                    self.logger.debug(f"Corner pixel wasn't white, assuming {bg_color_l_default} background for L.")
                    bg_color_to_match = bg_color_l_default
                def is_background(pixel):
                    return pixel == bg_color_to_match
            else:
                 self.logger.warning(f"Image mode {im.mode} not directly handled, converting to RGBA for scan.")
                 try:
                     im_for_scan = im.convert("RGBA")
                     bg_color_to_match = bg_color_rgb_default
                     self.logger.debug(f"Assuming RGBA background is transparent OR RGB={bg_color_to_match} after conversion.")
                     def is_background(pixel):
                        return pixel[3] == 0 or pixel[:3] == bg_color_to_match
                 except Exception as convert_err:
                     self.logger.error(f"Failed to convert image {os.path.basename(image_path)} from {im.mode} to RGBA: {convert_err}. Skipping trim.")
                     im.close()
                     return
            # ---

            # --- Find Bottom Content Boundary by Scanning ---
            scan_width, scan_height = im_for_scan.size
            bottom = scan_height - 1 # Default to original bottom

            found_bottom = False
            # Scan from bottom up
            for y in range(scan_height - 1, -1, -1):
                for x in range(scan_width):
                     # Check if pixel is NOT background
                     if not is_background(im_for_scan.getpixel((x, y))):
                        bottom = y # Found the last row with content
                        found_bottom = True
                        break # Stop checking this row
                if found_bottom:
                    break # Stop scanning rows

            # Close converted image if it exists
            if im_for_scan != im:
                im_for_scan.close()

            if not found_bottom:
                 self.logger.warning(f"Could not find non-background pixels (bottom scan). Image might be empty. Skipping trim for {os.path.basename(image_path)}")
                 im.close()
                 return
            # ---

            self.logger.debug(f"Pixel scan determined content bottom edge: y={bottom}")

            # Calculate the new lower bound for cropping, adding padding
            # (+1 because crop coordinates are (left, upper, right, lower) where lower is exclusive)
            new_lower_bound = min(height, bottom + 1 + bottom_padding)

            # Check if trimming is needed and the new bound is valid
            if new_lower_bound <= 0:
                self.logger.warning(f"Calculated new lower bound is invalid ({new_lower_bound}) for PNG {os.path.basename(image_path)}. Skipping crop.")
                final_bbox_for_crop = None
            elif new_lower_bound >= height: # Check if change is actually needed
                 self.logger.info(f"Calculated final height ({new_lower_bound}) is same or greater than original ({height}) for {os.path.basename(image_path)}. No bottom trim needed.")
                 final_bbox_for_crop = None
            else:
                 # Define the crop box: (left, upper, right, lower)
                 # Use original left (0), top (0), width, and the new lower bound
                 final_bbox_for_crop = (0, 0, width, new_lower_bound)

            if final_bbox_for_crop:
                padding_info = f"Bottom Pad:{bottom_padding}px"
                self.logger.info(f"Trimming Bottom PNG {os.path.basename(image_path)}: Original Size=({width},{height}), Content Bottom={bottom}, Final Box={final_bbox_for_crop} (Padding: {padding_info})")

                im_cropped = im.crop(final_bbox_for_crop)
                try:
                    save_kwargs = {}
                    if 'transparency' in im.info: save_kwargs['transparency'] = im.info['transparency']
                    im_cropped.save(image_path, **save_kwargs) # Overwrite original
                    self.log_message_signal.emit(f"  PNG Bottom Trim applied ({padding_info}): {os.path.basename(image_path)}")
                    self.logger.info(f"Successfully bottom-trimmed and saved {os.path.basename(image_path)} using Pillow.")
                except Exception as save_error:
                    self.log_message_signal.emit(f"  Error saving trimmed PNG {os.path.basename(image_path)}.")
                    self.logger.error(f"Error saving cropped PNG '{image_path}': {save_error}")
                finally:
                     im_cropped.close()

            im.close() # Close the original image object

        except FileNotFoundError:
             self.log_message_signal.emit(f"  Error trimming PNG: File not found {os.path.basename(image_path)}.")
             self.logger.error(f"Error trimming PNG: File not found at path '{image_path}'")
        except Exception as e:
            # Catch any other errors (including potential Pillow import errors if top-level failed)
            error_details = traceback.format_exc()
            self.log_message_signal.emit(f"  Error during PNG bottom trim for {os.path.basename(image_path)}: {type(e).__name__}")
            self.logger.error(f"Error trimming PNG bottom '{os.path.basename(image_path)}': {e}\n{error_details}")

    def _merge_pdfs(self, pdf_paths, output_path):
        """
        Merges a list of PDF files into a single output PDF file.
        Uses PyMuPDF (fitz).
        """
        if not pdf_paths:
            self.logger.warning("No PDF paths provided for merging.")
            return

        self.logger.info(f"Merging {len(pdf_paths)} PDFs into {output_path}")
        try:
            # Create a new empty PDF document
            merged_pdf = fitz.open()

            for pdf_path in pdf_paths:
                if not os.path.exists(pdf_path):
                    self.logger.warning(f"PDF file not found for merging, skipping: {pdf_path}")
                    continue
                try:
                    # Open each PDF and insert its pages into the merged document
                    with fitz.open(pdf_path) as doc:
                        merged_pdf.insert_pdf(doc)
                    self.logger.debug(f"Added {os.path.basename(pdf_path)} to merge.")
                except Exception as e:
                    self.logger.error(f"Error adding PDF '{os.path.basename(pdf_path)}' to merge: {e}")
                    # Continue with other PDFs even if one fails

            if merged_pdf.page_count > 0:
                merged_pdf.save(output_path)
                self.logger.info(f"Successfully merged PDFs to: {output_path}")
                self.log_message_signal.emit(f"✔ Merged PDFs to: {os.path.basename(output_path)}")
            else:
                self.logger.warning(f"No pages were merged into {output_path}. Output file not created.")
                self.log_message_signal.emit(f"ℹ No PDFs were successfully merged to {os.path.basename(output_path)}.")

        except Exception as e:
            self.logger.error(f"Error during PDF merging to {output_path}: {e}", exc_info=True)
            self.log_message_signal.emit(f"❌ Error merging PDFs to {os.path.basename(output_path)}: {e}")
        finally:
            if 'merged_pdf' in locals() and merged_pdf:
                merged_pdf.close()
                self.logger.debug("Closed merged PDF document handle.")

    def export_selected_views_once(self, server, all_views_in_workbook, output_folder, export_format, merge_pdfs_enabled, pdfs_for_single_merge):
        """Exports views selected in the UI directly, without using Excel data ('Export All Views Once' mode)."""
        try:
            self.log_message_signal.emit(f"--- Exporting Selected Views ({export_format}) ---")
            self.logger.info(f"Starting 'Export All Views Once' mode. Format={export_format}, Output='{output_folder}', Merge PDFs={merge_pdfs_enabled}")

            # Filter views based on global exclusion list
            selected_views_to_export = [view for view in all_views_in_workbook if view.name not in self.excluded_views_for_export]
            total_views = len(selected_views_to_export)

            if total_views == 0:
                message = "ℹ No views selected or remaining after global exclusion. Nothing to export."
                self.log_message_signal.emit(message)
                self.logger.info(message)
                self.progress_bar.setFormat("No views selected")
                return

            self.log_message_signal.emit(f"Attempting to export {total_views} selected view(s) to: {output_folder}")
            self.logger.info(f"Attempting to export {total_views} views: {[v.name for v in selected_views_to_export]}")
            processed_views = 0
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("Exporting 0%")

            numbering_enabled = self.numbering_checkbox.isChecked()
            counter = 1

            # Initialize export_options once before the loop
            if export_format == "PDF":
                export_options = PDFRequestOptions(page_type=PDFRequestOptions.PageType.Unspecified, maxage=0)
            else: # PNG
                export_options = ImageRequestOptions(imageresolution=ImageRequestOptions.Resolution.High, maxage=0)

            for view in selected_views_to_export:
                if self.stop_event.is_set():
                    self.log_message_signal.emit("⏹ Process stopped by user.")
                    self.logger.info("Stop event detected during 'Export All Views Once'.")
                    break

                # Determine filename (using view name)
                prefix = f"{counter:02d}_" if numbering_enabled else ""
                base_name_clean = "" # Define before try
                try:
                    # Sanitize base name directly from view name
                    base_name = view.name
                    base_name_clean = "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in base_name).strip()
                    max_len = 100
                    if len(base_name_clean) > max_len:
                        # ... (truncation logic - keep as before) ...
                        base_name_clean = base_name_clean[:max_len] # Simple truncate fallback
                    if not base_name_clean:
                        base_name_clean = f"view_{view.id}_invalid_name"
                except Exception as name_e:
                    self.logger.error(f"Error determining base filename for view '{view.name}': {name_e}. Using fallback.", exc_info=True)
                    base_name_clean = f"view_{view.id}_naming_error"

                file_name_final = f"{prefix}{base_name_clean}.{export_format.lower()}"
                
                # Determine the actual export path based on merge setting
                current_export_folder = output_folder
                if merge_pdfs_enabled and export_format == "PDF":
                    # If merging, export to a temporary directory inside HIDDEN_TEMP_DIR
                    current_export_folder = os.path.join(HIDDEN_TEMP_DIR, "temp_pdfs_for_single_merge")
                    os.makedirs(current_export_folder, exist_ok=True) # Ensure temp dir exists
                    # Append a unique identifier to the filename to avoid conflicts if views have same base_name
                    file_name_final = f"{prefix}{base_name_clean}_{view.id}.{export_format.lower()}"
                    self.logger.debug(f"Exporting to temp folder for single merge: {current_export_folder}/{file_name_final}")

                export_file_path = os.path.join(current_export_folder, file_name_final)

                # Export the single view - Pass None for org parts, pass cleaned base name
                self.export_single_view_with_retry(
                    server, view, export_options, export_file_path, export_format,
                    org1_val_clean=None, org2_val_clean=None, base_name_clean=base_name_clean
                )
                
                # If merging PDFs, add the path to the list
                if merge_pdfs_enabled and export_format == "PDF":
                    pdfs_for_single_merge.append(export_file_path)

                processed_views += 1
                progress_percentage = int((processed_views / total_views) * 100)
                self.progress_signal.emit(progress_percentage)
                self.progress_bar.setFormat(f"Exporting {progress_percentage}%")

                if numbering_enabled: counter += 1

            if not self.stop_event.is_set():
                self.log_message_signal.emit("✔ Finished exporting selected views.")
                self.logger.info("Finished 'Export All Views Once'.")

        except Exception as e:
            error_message = f"❌ An error occurred during 'Export All Views Once': {str(e)}"
            # print(error_message) # Less verbose print
            self.log_message_signal.emit(error_message)
            self.logger.error(f"Error in export_selected_views_once: {e}", exc_info=True)
            self.progress_bar.setFormat("Error")
            QMessageBox.critical(self, "Export Error", f"An error occurred during export:\n{e}")

    def _trim_pdf_whitespace(self, pdf_path):
        """
        Trims whitespace from the BOTTOM ONLY of a PDF's first page using PyMuPDF (fitz),
        leaving specified padding below the detected content.
        Saves to a temporary file first, then replaces the original.

        Args:
            pdf_path (str): The path to the PDF file to trim.
        """

        self.logger.info(f"Attempting to trim bottom whitespace for {pdf_path} using PyMuPDF.")
        temp_pdf_path = pdf_path + ".trimmed_tmp.pdf"
        doc = None

        try:
            # Ensure fitz is available (rely on top-level import)
            if 'fitz' not in sys.modules:
                 raise ImportError("PyMuPDF (fitz) is not imported. Please ensure 'import fitz' is at the top of the script.")

            try:
                doc = fitz.open(pdf_path)
            except Exception as open_error:
                 self.log_message_signal.emit(f"  Error opening PDF for trimming: {os.path.basename(pdf_path)}")
                 self.logger.error(f"Failed to open PDF '{os.path.basename(pdf_path)}' with PyMuPDF: {open_error}")
                 return # Cannot proceed if file can't be opened

            if doc.page_count == 0:
                self.logger.warning(f"PDF has no pages, cannot trim: {pdf_path}")
                if doc: doc.close()
                return

            page = doc[0] # Process the first page
            content_bbox = fitz.Rect() # Initialize empty bounding box

            # --- Find content boundary (only need bottom edge, but find full box for robustness) ---
            blocks = page.get_text("blocks")
            if blocks:
                for b in blocks: content_bbox.include_rect(fitz.Rect(b[:4]))
                self.logger.debug(f"Found text blocks for trimming {os.path.basename(pdf_path)}. Content bbox: {content_bbox}")
            else:
                # If no text, try drawings
                self.logger.debug(f"No text blocks found in {os.path.basename(pdf_path)}, checking drawings...")
                drawings = page.get_drawings()
                if drawings:
                    for d in drawings:
                         # Consider rectangles that seem like content areas
                         if d['rect'].width > 1 and d['rect'].height > 1: content_bbox.include_rect(d['rect'])
                    self.logger.debug(f"Found drawings for trimming {os.path.basename(pdf_path)}. Content bbox: {content_bbox}")
                else:
                    # If still no content found, cannot determine bottom edge
                    self.logger.warning(f"No text blocks or drawings found on page 1 of {os.path.basename(pdf_path)}. Cannot determine bottom boundary. Skipping trim.")
                    if doc: doc.close()
                    return

            # Check if a valid bounding box (specifically with height) was found
            if content_bbox.is_empty or content_bbox.height <= 0:
                self.logger.warning(f"Could not determine valid content bottom boundary for {os.path.basename(pdf_path)}. Skipping trim.")
                if doc: doc.close()
                return
            # ---

            # --- Define Bottom Padding (Points) ---
            # <<< --- ADJUST THIS VALUE AS NEEDED --- >>> (72 points = 1 inch)
            bottom_padding = 20.0
            # ------------------------------------

            original_rect = page.rect # Get the original page dimensions (MediaBox or CropBox)

            # Calculate the new bottom edge (y1 in fitz Rect) based on content bottom + padding
            # Ensure it doesn't exceed the original page bottom
            new_bottom_edge = min(original_rect.y1, content_bbox.y1 + bottom_padding)

            # Create the final crop box using original left, top, right and ONLY the new bottom
            final_crop_box = fitz.Rect(
                original_rect.x0,       # Original Left
                original_rect.y0,       # Original Top
                original_rect.x1,       # Original Right
                new_bottom_edge         # New Bottom edge
            )

            # --- Check if trimming is actually needed and the box is valid ---
            if final_crop_box.height <= 0:
                self.logger.warning(f"Calculated final crop box has invalid height ({final_crop_box}) for {os.path.basename(pdf_path)}. Skipping trim.")
                if doc: doc.close()
                return
            # Use a small tolerance for floating point comparison
            elif abs(final_crop_box.height - original_rect.height) < 0.1:
                self.logger.info(f"Calculated final height is effectively the same as original for {os.path.basename(pdf_path)}. No bottom trim needed.")
                if doc: doc.close()
                return # No need to save if no change
            # ---

            padding_info = f"Bottom Pad:{bottom_padding}pt"
            self.logger.info(f"Trimming Bottom PDF {os.path.basename(pdf_path)}: Original Rect={original_rect}, Content Bottom={content_bbox.y1:.2f}, Final Crop Box={final_crop_box} (Padding: {padding_info})")

            # Apply the new crop box to the page
            page.set_cropbox(final_crop_box)

            # --- Save to TEMPORARY file first ---
            temp_pdf_path = pdf_path + ".trimmed_tmp.pdf"
            self.logger.debug(f"Saving bottom-trimmed PDF to temporary file: {temp_pdf_path}")
            # Use garbage collection options for potentially smaller file size
            doc.save(temp_pdf_path, garbage=4, deflate=True, linear=True)
            self.logger.debug(f"Successfully saved temporary file: {temp_pdf_path}")
            # ---

            # --- Close the document handle BEFORE replacing the file ---
            doc.close()
            doc = None # Prevent closing again in finally if successful
            self.logger.debug(f"Closed original document handle for {pdf_path}")
            # ---

            # --- Replace original file with temporary file ---
            try:
                os.replace(temp_pdf_path, pdf_path)
                self.log_message_signal.emit(f"  PDF Bottom Trim applied ({padding_info}): {os.path.basename(pdf_path)}")
                self.logger.info(f"Successfully replaced original with bottom-trimmed PDF: {os.path.basename(pdf_path)}")
            except Exception as replace_error:
                 # Log error, temp file cleanup happens in finally
                 self.log_message_signal.emit(f"  Error replacing original PDF with trimmed version: {os.path.basename(pdf_path)}")
                 self.logger.error(f"Error replacing '{pdf_path}' with '{temp_pdf_path}': {replace_error}")
            # ---

        except ImportError as e:
            self.log_message_signal.emit(f"  Error trimming PDF: PyMuPDF library not found. Please run 'pip install PyMuPDF'.")
            self.logger.error(f"Failed to trim PDF '{os.path.basename(pdf_path)}'. PyMuPDF not installed? Error: {e}")
        except Exception as e:
            # Catch any other errors during the process
            error_details = traceback.format_exc()
            self.log_message_signal.emit(f"  Error during PDF bottom trim for {os.path.basename(pdf_path)}: {type(e).__name__}")
            self.logger.error(f"Error trimming PDF bottom '{os.path.basename(pdf_path)}': {e}\n{error_details}")
        finally:
            # Ensure the document is closed if it's still open (e.g., if error happened before close)
            if doc is not None:
                try:
                    doc.close()
                    self.logger.debug(f"Closed document handle in finally block for {pdf_path}")
                except Exception as close_e:
                    self.logger.error(f"Error closing PDF document '{os.path.basename(pdf_path)}' in finally block: {close_e}")

            # --- Clean up the temporary file ---
            if os.path.exists(temp_pdf_path):
                try:
                    os.remove(temp_pdf_path)
                    self.logger.debug(f"Removed temporary file: {temp_pdf_path}")
                except Exception as remove_error:
                    self.logger.error(f"Error removing temporary file '{temp_pdf_path}': {remove_error}")
            # ---

    def _trim_png_whitespace(self, image_path):
        """
        Trims whitespace from the BOTTOM ONLY of a PNG image using Pillow
        by scanning pixels, leaving specified padding below the detected content.
        Relies on top-level PIL imports.

        Args:
            image_path (str): The path to the PNG file to trim.
        """
        # Ensure necessary Pillow modules are imported at the top of the script file:
        # from PIL import Image
        # Also ensure standard libraries are imported: import os, sys, traceback

        self.logger.info(f"Attempting to trim bottom whitespace for {image_path} using Pillow (Pixel Scan method).")

        try:
            # --- Define Bottom Padding (Pixels) ---
            # <<< --- ADJUST THIS VALUE AS NEEDED --- >>>
            bottom_padding = 20
            # -------------------------------------------

            # Use the globally imported Image object
            im = Image.open(image_path)
            width, height = im.size
            self.logger.debug(f"Image mode for {os.path.basename(image_path)}: {im.mode}")

            # --- Determine Background Color & Check Function ---
            bg_pixel = im.getpixel((0, 0))
            self.logger.debug(f"Sampled background pixel (0,0): {bg_pixel} (Mode: {im.mode})")

            im_for_scan = im # Assume we scan the original unless converted
            bg_color_rgb_default = (255, 255, 255)
            bg_color_l_default = 255

            if im.mode == "RGBA":
                bg_color_to_match = bg_color_rgb_default
                self.logger.debug(f"Assuming RGBA background is transparent OR RGB={bg_color_to_match}")
                def is_background(pixel):
                    return pixel[3] == 0 or pixel[:3] == bg_color_to_match
            elif im.mode == "RGB":
                bg_color_to_match = bg_pixel
                if bg_color_to_match != bg_color_rgb_default:
                     self.logger.debug(f"Corner pixel wasn't white, assuming {bg_color_rgb_default} background for RGB.")
                     bg_color_to_match = bg_color_rgb_default
                def is_background(pixel):
                    return pixel == bg_color_to_match
            elif im.mode == "L":
                bg_color_to_match = bg_pixel
                if bg_color_to_match != bg_color_l_default:
                    self.logger.debug(f"Corner pixel wasn't white, assuming {bg_color_l_default} background for L.")
                    bg_color_to_match = bg_color_l_default
                def is_background(pixel):
                    return pixel == bg_color_to_match
            else:
                 self.logger.warning(f"Image mode {im.mode} not directly handled, converting to RGBA for scan.")
                 try:
                     im_for_scan = im.convert("RGBA")
                     bg_color_to_match = bg_color_rgb_default
                     self.logger.debug(f"Assuming RGBA background is transparent OR RGB={bg_color_to_match} after conversion.")
                     def is_background(pixel):
                        return pixel[3] == 0 or pixel[:3] == bg_color_to_match
                 except Exception as convert_err:
                     self.logger.error(f"Failed to convert image {os.path.basename(image_path)} from {im.mode} to RGBA: {convert_err}. Skipping trim.")
                     im.close()
                     return
            # ---

            # --- Find Bottom Content Boundary by Scanning ---
            scan_width, scan_height = im_for_scan.size
            bottom = scan_height - 1 # Default to original bottom

            found_bottom = False
            # Scan from bottom up
            for y in range(scan_height - 1, -1, -1):
                for x in range(scan_width):
                     # Check if pixel is NOT background
                     if not is_background(im_for_scan.getpixel((x, y))):
                        bottom = y # Found the last row with content
                        found_bottom = True
                        break # Stop checking this row
                if found_bottom:
                    break # Stop scanning rows

            # Close converted image if it exists
            if im_for_scan != im:
                im_for_scan.close()

            if not found_bottom:
                 self.logger.warning(f"Could not find non-background pixels (bottom scan). Image might be empty. Skipping trim for {os.path.basename(image_path)}")
                 im.close()
                 return
            # ---

            self.logger.debug(f"Pixel scan determined content bottom edge: y={bottom}")

            # Calculate the new lower bound for cropping, adding padding
            # (+1 because crop coordinates are (left, upper, right, lower) where lower is exclusive)
            new_lower_bound = min(height, bottom + 1 + bottom_padding)

            # Check if trimming is needed and the new bound is valid
            if new_lower_bound <= 0:
                self.logger.warning(f"Calculated new lower bound is invalid ({new_lower_bound}) for PNG {os.path.basename(image_path)}. Skipping crop.")
                final_bbox_for_crop = None
            elif new_lower_bound >= height: # Check if change is actually needed
                 self.logger.info(f"Calculated final height ({new_lower_bound}) is same or greater than original ({height}) for {os.path.basename(image_path)}. No bottom trim needed.")
                 final_bbox_for_crop = None
            else:
                 # Define the crop box: (left, upper, right, lower)
                 # Use original left (0), top (0), width, and the new lower bound
                 final_bbox_for_crop = (0, 0, width, new_lower_bound)

            if final_bbox_for_crop:
                padding_info = f"Bottom Pad:{bottom_padding}px"
                self.logger.info(f"Trimming Bottom PNG {os.path.basename(image_path)}: Original Size=({width},{height}), Content Bottom={bottom}, Final Box={final_bbox_for_crop} (Padding: {padding_info})")

                im_cropped = im.crop(final_bbox_for_crop)
                try:
                    save_kwargs = {}
                    if 'transparency' in im.info: save_kwargs['transparency'] = im.info['transparency']
                    im_cropped.save(image_path, **save_kwargs) # Overwrite original
                    self.log_message_signal.emit(f"  PNG Bottom Trim applied ({padding_info}): {os.path.basename(image_path)}")
                    self.logger.info(f"Successfully bottom-trimmed and saved {os.path.basename(image_path)} using Pillow.")
                except Exception as save_error:
                    self.log_message_signal.emit(f"  Error saving trimmed PNG {os.path.basename(image_path)}.")
                    self.logger.error(f"Error saving cropped PNG '{image_path}': {save_error}")
                finally:
                     im_cropped.close()

            im.close() # Close the original image object

        except FileNotFoundError:
             self.log_message_signal.emit(f"  Error trimming PNG: File not found {os.path.basename(image_path)}.")
             self.logger.error(f"Error trimming PNG: File not found at path '{image_path}'")
        except Exception as e:
            # Catch any other errors (including potential Pillow import errors if top-level failed)
            error_details = traceback.format_exc()
            self.log_message_signal.emit(f"  Error during PNG bottom trim for {os.path.basename(image_path)}: {type(e).__name__}")
            self.logger.error(f"Error trimming PNG bottom '{os.path.basename(image_path)}': {e}\n{error_details}")

    def _merge_pdfs(self, pdf_paths, output_path):
        """
        Merges a list of PDF files into a single output PDF file.
        Uses PyMuPDF (fitz).
        """
        if not pdf_paths:
            self.logger.warning("No PDF paths provided for merging.")
            return

        self.logger.info(f"Merging {len(pdf_paths)} PDFs into {output_path}")
        try:
            # Create a new empty PDF document
            merged_pdf = fitz.open()

            for pdf_path in pdf_paths:
                if not os.path.exists(pdf_path):
                    self.logger.warning(f"PDF file not found for merging, skipping: {pdf_path}")
                    continue
                try:
                    # Open each PDF and insert its pages into the merged document
                    with fitz.open(pdf_path) as doc:
                        merged_pdf.insert_pdf(doc)
                    self.logger.debug(f"Added {os.path.basename(pdf_path)} to merge.")
                except Exception as e:
                    self.logger.error(f"Error adding PDF '{os.path.basename(pdf_path)}' to merge: {e}")
                    # Continue with other PDFs even if one fails

            if merged_pdf.page_count > 0:
                merged_pdf.save(output_path)
                self.logger.info(f"Successfully merged PDFs to: {output_path}")
                self.log_message_signal.emit(f"✔ Merged PDFs to: {os.path.basename(output_path)}")
            else:
                self.logger.warning(f"No pages were merged into {output_path}. Output file not created.")
                self.log_message_signal.emit(f"ℹ No PDFs were successfully merged to {os.path.basename(output_path)}.")

        except Exception as e:
            self.logger.error(f"Error during PDF merging to {output_path}: {e}", exc_info=True)
            self.log_message_signal.emit(f"❌ Error merging PDFs to {os.path.basename(output_path)}: {e}")
        finally:
            if 'merged_pdf' in locals() and merged_pdf:
                merged_pdf.close()
                self.logger.debug("Closed merged PDF document handle.")

    def export_single_view_with_retry(self, server, view, export_options, export_file_path, export_format,
                                    org1_val_clean=None, org2_val_clean=None, base_name_clean=""):
        """Exports a single Tableau view with retry logic and improved logging."""
        max_retries = 3
        retries = 0
        success = False
        while retries < max_retries and not success and not self.stop_event.is_set():
            try:
                self.logger.debug(f"Attempt {retries + 1} to export view '{view.name}' to {export_file_path}")
                start_export_time = time.time()

                # Populate data (PDF or Image)
                if export_format == "PDF":
                    server.views.populate_pdf(view, export_options)
                    export_data = view.pdf
                else: # PNG
                    server.views.populate_image(view, export_options)
                    export_data = view.image

                export_duration = time.time() - start_export_time
                self.logger.debug(f"View '{view.name}' data populated in {export_duration:.2f}s.")

                # Write data to file
                with open(export_file_path, 'wb') as f:
                    f.write(export_data)

                # *** Construct Relative Path for Logging ***
                log_path_parts = []
                if org1_val_clean: log_path_parts.append(f"[{org1_val_clean}]") # Add brackets for clarity
                if org2_val_clean: log_path_parts.append(f"[{org2_val_clean}]")
                # Use base_name (cleaned, before prefix/extension) + extension for the file part
                file_part = base_name_clean + f".{export_format.lower()}"
                log_path_parts.append(file_part)
                relative_log_path = " - ".join(log_path_parts)
                # *** End Construct Relative Path ***

                success_msg = f"✔ Exported: {relative_log_path}" # Use new path string
                self.log_message_signal.emit(success_msg)
                self.logger.info(f"Successfully exported view '{view.name}' as '{os.path.basename(export_file_path)}' (Log Ref: {relative_log_path}) (Attempt {retries + 1}).")
                success = True

                # Perform trimming if requested and successful export of PDF
                if success and self.trim_pdf_checkbox.isChecked(): # Checkbox now enables trimming for both
                    if export_format == "PDF":
                        self.log_message_signal.emit(f"  Attempting PDF trim for: {os.path.basename(export_file_path)}")
                        self._trim_pdf_whitespace(export_file_path) # Existing PDF trim
                    elif export_format == "PNG":
                        self.log_message_signal.emit(f"  Attempting PNG trim for: {os.path.basename(export_file_path)}")
                        self._trim_png_whitespace(export_file_path) 

            except Exception as e:
                retries += 1
                error_msg = f"❌ Error exporting '{view.name}' (Attempt {retries}/{max_retries}): {str(e)}"
                # print(f"      {error_msg}") # Less verbose print
                self.log_message_signal.emit(error_msg)
                self.logger.warning(f"Export attempt {retries} failed for view '{view.name}': {e}", exc_info=(retries == max_retries)) # Log full traceback on last retry

                # Wait before retrying, unless stop requested
                if retries < max_retries and not self.stop_event.is_set():
                    wait_time = 5 * retries
                    wait_msg = f"Waiting {wait_time}s before retrying..."
                    # print(f"      {wait_msg}") # Less verbose print
                    self.logger.info(f"Waiting {wait_time}s before retrying view '{view.name}'.")
                    stopped_during_wait = self.stop_event.wait(timeout=wait_time)
                    if stopped_during_wait:
                        self.logger.info(f"Stop event detected during retry wait for view '{view.name}'. Aborting retries.")
                        break # Exit retry loop
                elif not self.stop_event.is_set():
                    # Max retries reached, log final failure
                    fail_msg = f"‼ Max retries reached for '{view.name}'. Skipping this view."
                    self.log_message_signal.emit(fail_msg)
                    self.logger.error(f"Max retries reached for view '{view.name}'. Skipping.")
                else:
                    # Stopped during wait or before first attempt
                    self.logger.info(f"Stop requested, skipping further export attempts for '{view.name}'.")

    def OnStop(self):
        """Handles the Stop button click, signaling the worker thread to terminate."""
        print("Stopping process..."); self.logger.info("Stop requested by user."); self.log_message_signal.emit("⏹ Stop requested. Finishing current operation (if any)..."); self.stop_event.set(); self.stop_btn.setEnabled(False)


    @pyqtSlot(str)
    def update_log(self, message):
        """Appends a message to the log QTextEdit in a thread-safe manner."""
        try:
            self.log_text.append(message);
            self.log_text.ensureCursorVisible() # Scroll to the bottom

            # --- Log Trimming Logic ---
            # Limit the number of lines in the log text edit for performance
            max_lines = 2000 # Adjust as needed
            current_lines = self.log_text.document().blockCount()

            if current_lines > max_lines:
                cursor = self.log_text.textCursor()
                cursor.movePosition(QTextCursor.Start)
                # Keep roughly half the max lines
                lines_to_remove = current_lines - (max_lines // 2)
                # Move cursor down, keeping anchor at the start
                cursor.movePosition(QTextCursor.Down, QTextCursor.KeepAnchor, lines_to_remove)
                cursor.removeSelectedText()
                # Move cursor back to the end
                cursor.movePosition(QTextCursor.End)
                self.log_text.setTextCursor(cursor)
                # Log trimming occasionally
                if current_lines % 100 == 0:
                    self.logger.debug(f"Log widget trimmed to approx {max_lines // 2} lines.")
            # --- End Log Trimming ---
        except RuntimeError:
             # This can happen if the log widget is deleted while a signal is pending
             self.logger.warning("RuntimeError updating log widget, it might have been deleted.")
        except Exception as e:
             self.logger.error(f"Unexpected error updating log: {e}")


    @pyqtSlot(int)
    def update_progress(self, value):
        """Updates the progress bar value in a thread-safe manner."""
        try:
            self.progress_bar.setValue(value)
        except RuntimeError:
             self.logger.warning("RuntimeError updating progress bar, it might have been deleted.")
        except Exception as e:
             self.logger.error(f"Unexpected error updating progress bar: {e}")

    def check_for_updates(self):
        """Checks GitHub for the latest release version."""
        update_thread = Thread(target=self._perform_update_check, daemon=True); update_thread.start()

    def _perform_update_check(self):
        """Worker function for checking updates."""
        self.logger.info("Checking for updates..."); repo_url = "https://api.github.com/repos/haythamsoufi/TableauPDF/releases/latest"
        try:
            self.logger.debug(f"Requesting latest release info from {repo_url}"); response = requests.get(repo_url, timeout=10); response.raise_for_status()
            latest_release = response.json(); latest_version = latest_release.get("tag_name"); release_notes = latest_release.get("body", "No release notes provided."); download_url = latest_release.get("html_url")
            if not latest_version or not download_url: self.logger.warning("Could not parse latest release information from GitHub API response."); return
            self.logger.info(f"Current version: {CURRENT_VERSION}, Latest version found: {latest_version}")
            try:
                from packaging.version import parse as parse_version;
                update_available = parse_version(latest_version.lstrip('v')) > parse_version(CURRENT_VERSION.lstrip('v'))
            except ImportError:
                self.logger.warning("`packaging` library not found, using simple string comparison for version check.");
                update_available = latest_version != CURRENT_VERSION
            except Exception as parse_e:
                 self.logger.error(f"Error parsing versions for comparison: {parse_e}. Assuming no update.")
                 update_available = False

            if update_available:
                self.logger.info(f"Update available: {latest_version}");
                # Use signal to inform user on main thread
                self.log_message_signal.emit(f"ℹ Update Available: Version {latest_version} released. See GitHub for details: <a href='{download_url}'>{download_url}</a>")
                print(f"--- UPDATE AVAILABLE: Version {latest_version} ---");
                # print(f"--- Release Notes: ---\n{release_notes}"); # Can be long
                print(f"--- Download at: {download_url} ---")
            else:
                self.logger.info("Application is up-to-date.")
        except requests.exceptions.Timeout: self.logger.warning(f"Timeout occurred while checking for updates at {repo_url}.")
        except requests.exceptions.RequestException as e: self.logger.error(f"Error checking for updates: {e}")
        except Exception as e: self.logger.error(f"An unexpected error occurred during update check: {e}", exc_info=True)


    def closeEvent(self, event):
         """Handles the window close event."""
         self.logger.info("Close event triggered.")
         if hasattr(self, 'worker') and self.worker and self.worker.is_alive():
              reply = QMessageBox.question(self, 'Confirm Exit', "An export task is currently running.\nStopping the task might result in incomplete exports.\n\nAre you sure you want to exit?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
              if reply == QMessageBox.Yes:
                  self.logger.info("User confirmed exit while task running. Signaling stop.");
                  self.stop_event.set(); # Signal the thread to stop
                  # Optionally wait a short time for thread to potentially finish current step
                  # self.worker.join(timeout=1.0)
                  event.accept() # Close the window
              else:
                  self.logger.info("User cancelled exit.");
                  event.ignore() # Keep window open
         else:
             self.logger.info("Accepting close event (no task running).");
             event.accept()

    # --- ADD HELPER RESETS (if not already present) ---
    # Add these if you don't have them, adapt reset_combined_logic if needed

    def reset_filters(self):
        """Removes all filter lines from the UI and clears the internal list."""
        print("Resetting all filters...")
        self.logger.info("Resetting all Excel list filters.")
        # Iterate backwards to safely remove while iterating
        for i in range(len(self.filters) - 1, -1, -1):
             try:
                 filter_info = self.filters[i]
                 hbox = filter_info['hbox']
                 # Check if hbox might already be deleted or detached
                 if hbox and (hbox.parentWidget() or hbox.parentLayout()):
                     self.remove_filter_line(hbox) # Reuse the corrected remove logic
                 elif filter_info in self.filters:
                     # If layout seems gone but info is in list, just remove info
                     self.logger.warning(f"Filter hbox for index {i} seems detached but info exists. Removing info only.")
                     self.filters.pop(i)
                 else:
                      # Should not happen if logic is correct, but log if it does
                      self.logger.warning(f"Filter info for index {i} not found in list during reset.")

             except IndexError:
                 self.logger.error("IndexError during filter reset, list might have changed unexpectedly.")
             except Exception as e:
                 self.logger.error(f"Error removing filter during reset_filters for index {i}: {e}", exc_info=True)
        # Ensure list is cleared even if errors occurred
        self.filters = []


    def reset_conditions(self):
        """Removes all condition lines."""
        self.logger.info("Resetting all conditions.")
        for i in range(len(self.conditions) - 1, -1, -1):
            try:
                condition_info = self.conditions[i]
                hbox = condition_info['hbox']
                if hbox and (hbox.parentWidget() or hbox.parentLayout()):
                    self.remove_condition_line(hbox) # Assumes this method exists and is correct
                elif condition_info in self.conditions:
                    self.logger.warning(f"Condition hbox for index {i} seems detached but info exists. Removing info only.")
                    self.conditions.pop(i)
                else:
                    self.logger.warning(f"Condition info for index {i} not found in list during reset.")

            except IndexError:
                 self.logger.error("IndexError during condition reset.")
            except Exception as e:
                 self.logger.error(f"Error removing condition during reset_conditions for index {i}: {e}", exc_info=True)
        self.conditions = []


    def reset_parameters(self):
        """Removes all parameter lines."""
        self.logger.info("Resetting all parameters.")
        for i in range(len(self.parameters) - 1, -1, -1):
            try:
                parameter_info = self.parameters[i]
                hbox = parameter_info['hbox']
                if hbox and (hbox.parentWidget() or hbox.parentLayout()):
                    self.remove_parameter_line(hbox) # Assumes this method exists and is correct
                elif parameter_info in self.parameters:
                    self.logger.warning(f"Parameter hbox for index {i} seems detached but info exists. Removing info only.")
                    self.parameters.pop(i)
                else:
                    self.logger.warning(f"Parameter info for index {i} not found in list during reset.")

            except IndexError:
                 self.logger.error("IndexError during parameter reset.")
            except Exception as e:
                 self.logger.error(f"Error removing parameter during reset_parameters for index {i}: {e}", exc_info=True)
        self.parameters = []


    def reset_combined_logic(self):
        """Resets all filters, conditions, and parameters by calling individual reset functions."""
        self.logger.info("Resetting combined logic section (Filters, Conditions, Parameters).")
        self.reset_filters()
        self.reset_conditions()
        self.reset_parameters()
        self.logger.info("Cleared all filters, conditions, and parameters.")
        # Optional: Force layout update if needed
        # self.combined_logic_layout.update()
        # self.layout().activate()

def current_dir():
    """Returns the directory of the currently running script or executable."""
    if getattr(sys, 'frozen', False):
        # If running as a bundled executable (e.g., PyInstaller)
        return os.path.dirname(sys.executable)
    else:
        # If running as a script
        try:
            # __file__ might not be defined in some environments
            return os.path.dirname(os.path.abspath(__file__))
        except NameError:
            return os.getcwd() # Fallback to current working directory
          
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    ex = PDFExportApp()

    # *** Apply Custom Theme INITIALLY ***
    print(f"Attempting to apply initial custom theme...")
    if ex.apply_custom_theme(): # Call the method on the instance
         ex.is_custom_theme_active = True # Set state if successful
         ex._update_theme_button_text() # Update button text right away
    else:
         # If loading fails on startup, ensure state is False
         ex.is_custom_theme_active = False
         ex._update_theme_button_text()
         print("Failed to load initial custom theme, using default.")

    ex.show()
    sys.exit(app.exec_())
    # --- End Application Setup ---