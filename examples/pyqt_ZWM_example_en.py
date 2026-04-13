import sys
import os

# Ensure pyzwcadmech library can be found
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

try:
    from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                                 QHBoxLayout, QPushButton, QLabel, QLineEdit, 
                                 QFileDialog, QMessageBox, QGroupBox, QGridLayout,
                                 QDialog, QTableWidget, QTableWidgetItem, QHeaderView,
                                 QComboBox)
except ImportError:
    print("PyQt5 is not installed. Please run: pip install PyQt5")
    sys.exit(1)

from pyzwcadmech import ZwCADMech

class TitleDialog(QDialog):
    def __init__(self, mech, parent=None):
        super().__init__(parent)
        self.mech = mech
        self.title_obj = mech.get_title()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("Title Block")
        self.resize(450, 400)
        layout = QVBoxLayout(self)
        
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Label", "Value"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.table)
        
        btn_layout = QHBoxLayout()
        self.btn_read = QPushButton("Read")
        self.btn_write = QPushButton("Write")
        self.btn_refresh = QPushButton("Refresh")
        
        self.btn_read.clicked.connect(self.on_read)
        self.btn_write.clicked.connect(self.on_write)
        self.btn_refresh.clicked.connect(self.on_refresh)
        
        btn_layout.addWidget(self.btn_read)
        btn_layout.addWidget(self.btn_write)
        btn_layout.addWidget(self.btn_refresh)
        layout.addLayout(btn_layout)
        
    def on_read(self):
        if not self.title_obj: return
        self.table.setRowCount(0)
        
        count = self.title_obj.get_item_count()
        self.table.setRowCount(count)
        
        for i in range(count):
            label, name, value = self.title_obj.get_item(i)
            self.table.setItem(i, 0, QTableWidgetItem(label))
            self.table.setItem(i, 1, QTableWidgetItem(value))

    def on_write(self):
        if not self.title_obj: return
        for i in range(self.table.rowCount()):
            label_item = self.table.item(i, 0)
            value_item = self.table.item(i, 1)
            if label_item and value_item:
                self.title_obj.set_item(label_item.text(), value_item.text())
        
        self.table.viewport().update()
        QMessageBox.information(self, "Info", "Write successful")

    def on_refresh(self):
        if self.mech and self.mech.zwm_db:
            self.mech.zwm_db.refresh_title()
            QMessageBox.information(self, "Info", "Refresh successful")


class BomDialog(QDialog):
    def __init__(self, mech, parent=None):
        super().__init__(parent)
        self.mech = mech
        self.bom_obj = mech.get_bom()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("BOM")
        self.resize(600, 400)
        layout = QVBoxLayout(self)
        
        self.table = QTableWidget()
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.table)
        
        btn_layout = QHBoxLayout()
        self.btn_read = QPushButton("Read")
        self.btn_write = QPushButton("Write")
        self.btn_refresh = QPushButton("Refresh")
        self.btn_add = QPushButton("Add")
        self.btn_delete = QPushButton("Delete")
        
        self.btn_read.clicked.connect(self.on_read)
        self.btn_write.clicked.connect(self.on_write)
        self.btn_refresh.clicked.connect(self.on_refresh)
        self.btn_add.clicked.connect(self.on_add)
        self.btn_delete.clicked.connect(self.on_delete)
        
        btn_layout.addWidget(self.btn_read)
        btn_layout.addWidget(self.btn_write)
        btn_layout.addWidget(self.btn_refresh)
        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_delete)
        layout.addLayout(btn_layout)
        
    def on_read(self):
        if not self.bom_obj: return
        self.table.clear()
        self.table.setColumnCount(0)
        self.table.setRowCount(0)
        
        row_count = self.bom_obj.get_item_count()
        self.table.setRowCount(row_count)
        
        for r in range(row_count):
            bom_row = self.bom_obj.get_item(r)
            
            if bom_row:
                col_count = bom_row.get_item_count()
                
                if r == 0:
                    self.table.setColumnCount(col_count)
                    headers = []
                    for c in range(col_count):
                        label, name, value = bom_row.get_item(c)
                        headers.append(label)
                    self.table.setHorizontalHeaderLabels(headers)
                
                for c in range(col_count):
                    label, name, value = bom_row.get_item(c)
                    self.table.setItem(r, c, QTableWidgetItem(value))

    def on_write(self):
        if not self.bom_obj: return
        
        row_count = self.bom_obj.get_item_count()
        
        for r in range(row_count):
            bom_row = self.bom_obj.get_item(r)
            
            if bom_row:
                col_count = bom_row.get_item_count()
                for c in range(col_count):
                    label, name, value = bom_row.get_item(c)
                    
                    item = self.table.item(r, c)
                    if item:
                        bom_row.set_item(label, item.text())
                
                self.bom_obj.set_item(r, bom_row)
        
        self.table.viewport().update()
        QMessageBox.information(self, "Info", "Write successful")

    def on_refresh(self):
        if self.mech and self.mech.zwm_db:
            self.mech.zwm_db.refresh_bom()
            QMessageBox.information(self, "Info", "Refresh successful")

    def on_add(self):
        if not self.bom_obj: return
        
        bom_row = self.bom_obj.create_bom_row()
        if bom_row:
            col_count = bom_row.get_item_count()
            for c in range(col_count):
                label, name, value = bom_row.get_item(c)
                new_val = label + name + value
                bom_row.set_item(label, new_val)
            
            self.bom_obj.add_item(bom_row)
            if self.mech and self.mech.zwm_db:
                self.mech.zwm_db.refresh_bom()
            self.on_read()

    def on_delete(self):
        if not self.bom_obj: return
        selected_rows = set(item.row() for item in self.table.selectedItems())
        if not selected_rows:
            QMessageBox.warning(self, "Warning", "Please select rows to delete first")
            return
            
        for r in sorted(selected_rows, reverse=True):
            self.bom_obj.delete_item(r)
            
        if self.mech and self.mech.zwm_db:
            self.mech.zwm_db.refresh_bom()
            
        count = self.bom_obj.get_item_count()
        if count == 0:
            if self.mech and self.mech.zwm_app:
                self.mech.zwm_app.send_command("_.ZwmPartlist\n")
                
        self.on_read()


class FrameDialog(QDialog):
    def __init__(self, mech, parent=None):
        super().__init__(parent)
        self.mech = mech
        self.frame_obj = None
        self.current_frame_name = ""
        self.initUI()
        self.load_initial_data()
        
    def initUI(self):
        self.setWindowTitle("Drawing Frame")
        self.resize(500, 600)
        layout = QVBoxLayout(self)
        
        # Frame selection
        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("Select Frame:"))
        self.combo_frame = QComboBox()
        self.combo_frame.currentTextChanged.connect(self.on_frame_changed)
        top_layout.addWidget(self.combo_frame)
        layout.addLayout(top_layout)
        
        # Property grid
        grid = QGridLayout()
        self.fields = {}
        labels = [
            ("std_name", "Standard"), ("frame_size_name", "Frame Size"), ("frame_style_name", "Frame Style"),
            ("orientation", "Orientation"), ("width", "Width"), ("height", "Height"),
            ("title_style_name", "Title Block"), ("bom_style_name", "BOM"), ("dhl_style_name", "Code Block"),
            ("fjl_style_name", "Additional Block"), ("csl_style_name", "Parameter Block"), ("ggl_style_name", "Modification Block"),
            ("have_dhl", "Has Code Block"), ("have_fjl", "Has Add. Block"), ("have_btl", "Has Title Block"),
            ("have_csl", "Has Param Block"), ("have_ggl", "Has Mod Block"), ("scale1", "Scale 1"), ("scale2", "Scale 2")
        ]
        
        for i, (key, label) in enumerate(labels):
            grid.addWidget(QLabel(label + ":"), i, 0)
            le = QLineEdit()
            self.fields[key] = le
            grid.addWidget(le, i, 1)
            
        layout.addLayout(grid)
        
        # Buttons
        btn_layout = QHBoxLayout()
        self.btn_read = QPushButton("Read")
        self.btn_refresh = QPushButton("Refresh Dict")
        self.btn_rebuild = QPushButton("Rebuild Frame")
        self.btn_create = QPushButton("New Frame")
        
        self.btn_read.clicked.connect(self.on_read)
        self.btn_refresh.clicked.connect(self.on_refresh)
        self.btn_rebuild.clicked.connect(self.on_rebuild)
        self.btn_create.clicked.connect(self.on_create)
        
        btn_layout.addWidget(self.btn_read)
        btn_layout.addWidget(self.btn_refresh)
        btn_layout.addWidget(self.btn_rebuild)
        btn_layout.addWidget(self.btn_create)
        layout.addLayout(btn_layout)
        
    def load_initial_data(self):
        if not self.mech or not self.mech.zwm_db: return
        
        count = self.mech.zwm_db.get_frame_count()
        
        self.combo_frame.blockSignals(True)
        self.combo_frame.clear()
        for i in range(count):
            name = self.mech.zwm_db.get_frame_name(i)
            self.combo_frame.addItem(name)
        self.combo_frame.blockSignals(False)
        
        if self.combo_frame.count() > 0:
            self.combo_frame.setCurrentIndex(0)
            self.current_frame_name = self.combo_frame.currentText()

    def on_frame_changed(self, text):
        self.current_frame_name = text

    def on_read(self):
        if not self.mech or not self.mech.zwm_db: return
        
        self.mech.zwm_db.switch_frame(self.current_frame_name)
        self.frame_obj = self.mech.zwm_db.get_frame()
        
        if self.frame_obj:
            for key, le in self.fields.items():
                try:
                    val = getattr(self.frame_obj, key, "")
                    le.setText(str(val))
                except Exception:
                    pass

    def write_to_frame(self):
        if not self.frame_obj: return
        for key, le in self.fields.items():
            val_str = le.text()
            if key in ["width", "height", "have_dhl", "have_fjl", "have_btl", "have_csl", "have_ggl"]:
                try:
                    setattr(self.frame_obj, key, int(val_str))
                except ValueError:
                    pass
            else:
                try:
                    setattr(self.frame_obj, key, val_str)
                except Exception:
                    pass

    def on_refresh(self):
        if not self.mech or not self.mech.zwm_db: return
        
        self.mech.zwm_db.switch_frame(self.current_frame_name)
        self.frame_obj = self.mech.zwm_db.get_frame()
        
        self.write_to_frame()
        self.mech.zwm_db.refresh_frame()
        QMessageBox.information(self, "Info", "Refresh successful")

    def on_rebuild(self):
        if not self.mech or not self.mech.zwm_db: return
        
        self.mech.zwm_db.switch_frame(self.current_frame_name)
        self.frame_obj = self.mech.zwm_db.get_frame()
        
        self.write_to_frame()
        self.mech.zwm_db.build_frame(511)
        QMessageBox.information(self, "Info", "Rebuild successful")

    def on_create(self):
        if not self.mech or not self.mech.zwm_db: return
        
        self.frame_obj, name = self.mech.zwm_db.get_next_frm_name()
        
        self.current_frame_name = name
        self.mech.zwm_db.switch_frame(self.current_frame_name)
        self.frame_obj = self.mech.zwm_db.get_frame()
        
        self.write_to_frame()
        
        self.combo_frame.addItem(self.current_frame_name)
        self.combo_frame.setCurrentText(self.current_frame_name)
        
        self.mech.zwm_db.build_frame(511)
        QMessageBox.information(self, "Info", "Create successful")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.mech = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('ZWCAD Mechanical PLM Test (PyQt5)')
        self.resize(600, 450)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 1. Connection Area
        conn_group = QGroupBox("Connect to ZWCAD MFG")
        conn_layout = QGridLayout()
        
        self.btn_start = QPushButton("Start/Connect ZWCAD MFG")
        self.btn_start.clicked.connect(self.on_start_clicked)
        
        self.txt_cad_path = QLineEdit()
        self.txt_cad_path.setReadOnly(True)
        self.txt_zwm_path = QLineEdit()
        self.txt_zwm_path.setReadOnly(True)
        self.txt_version = QLineEdit()
        self.txt_version.setReadOnly(True)

        conn_layout.addWidget(self.btn_start, 0, 0, 1, 2)
        conn_layout.addWidget(QLabel("CAD Path:"), 1, 0)
        conn_layout.addWidget(self.txt_cad_path, 1, 1)
        conn_layout.addWidget(QLabel("ZWM Path:"), 2, 0)
        conn_layout.addWidget(self.txt_zwm_path, 2, 1)
        conn_layout.addWidget(QLabel("Version:"), 3, 0)
        conn_layout.addWidget(self.txt_version, 3, 1)
        conn_group.setLayout(conn_layout)
        main_layout.addWidget(conn_group)

        # 2. File Operation Area
        file_group = QGroupBox("File Operations")
        file_layout = QHBoxLayout()
        
        self.txt_file = QLineEdit()
        self.btn_browse = QPushButton("Browse...")
        self.btn_browse.clicked.connect(self.on_browse_clicked)
        self.btn_open = QPushButton("Open Drawing")
        self.btn_open.clicked.connect(self.on_open_clicked)
        
        file_layout.addWidget(QLabel("DWG File:"))
        file_layout.addWidget(self.txt_file)
        file_layout.addWidget(self.btn_browse)
        file_layout.addWidget(self.btn_open)
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)

        # 3. Data Acquisition Area
        data_group = QGroupBox("Data Acquisition & Save")
        data_layout = QHBoxLayout()
        
        self.btn_title = QPushButton("Get Title Block")
        self.btn_bom = QPushButton("Get BOM")
        self.btn_frame = QPushButton("Get Frame")
        self.btn_save = QPushButton("Save")
        self.btn_close = QPushButton("Close Drawing")
        
        self.btn_title.clicked.connect(self.on_title_clicked)
        self.btn_bom.clicked.connect(self.on_bom_clicked)
        self.btn_frame.clicked.connect(self.on_frame_clicked)
        self.btn_save.clicked.connect(self.on_save_clicked)
        self.btn_close.clicked.connect(self.on_close_clicked)
        
        data_layout.addWidget(self.btn_title)
        data_layout.addWidget(self.btn_bom)
        data_layout.addWidget(self.btn_frame)
        data_layout.addWidget(self.btn_save)
        data_layout.addWidget(self.btn_close)
        data_group.setLayout(data_layout)
        main_layout.addWidget(data_group)

        # 4. Interactive Editing Area
        edit_group = QGroupBox("Interactive Editing")
        edit_layout = QHBoxLayout()
        
        self.btn_title_edit = QPushButton("Edit Title Block")
        self.btn_mxb_edit = QPushButton("Edit BOM")
        self.btn_frame_edit = QPushButton("Edit Frame")
        self.btn_fjl_edit = QPushButton("Edit Add. Block")
        
        self.btn_title_edit.clicked.connect(self.on_title_edit_clicked)
        self.btn_mxb_edit.clicked.connect(self.on_mxb_edit_clicked)
        self.btn_frame_edit.clicked.connect(self.on_frame_edit_clicked)
        self.btn_fjl_edit.clicked.connect(self.on_fjl_edit_clicked)
        
        edit_layout.addWidget(self.btn_title_edit)
        edit_layout.addWidget(self.btn_mxb_edit)
        edit_layout.addWidget(self.btn_frame_edit)
        edit_layout.addWidget(self.btn_fjl_edit)
        edit_group.setLayout(edit_layout)
        main_layout.addWidget(edit_group)

        main_layout.addStretch()

        # Initialize button states
        self.update_buttons_state(False)

    def update_buttons_state(self, enabled):
        self.btn_title.setEnabled(enabled)
        self.btn_bom.setEnabled(enabled)
        self.btn_frame.setEnabled(enabled)
        self.btn_save.setEnabled(enabled)
        self.btn_close.setEnabled(enabled)
        
        self.btn_title_edit.setEnabled(enabled)
        self.btn_mxb_edit.setEnabled(enabled)
        self.btn_frame_edit.setEnabled(enabled)
        self.btn_fjl_edit.setEnabled(enabled)

    def on_start_clicked(self):
        try:
            self.mech = ZwCADMech()
            
            # Trigger connection
            app = self.mech.app
            zwm_app = self.mech.zwm_app
            zwm_db = self.mech.zwm_db
            
            if zwm_app:
                try:
                    # Get path information of the mechanical module
                    self.txt_cad_path.setText(zwm_app.get_cad_path())
                    self.txt_zwm_path.setText(zwm_app.get_zwm_path())
                    self.txt_version.setText(zwm_app.get_version())
                    QMessageBox.information(self, "Success", "Successfully connected to ZWCAD MFG!")
                except Exception as e:
                    QMessageBox.warning(self, "Warning", f"Failed to get path info: {e}")
            else:
                QMessageBox.critical(self, "Error", "Failed to create ZwmToolKit.ZwmApp!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Connection failed: {e}")

    def on_browse_clicked(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Select DWG File", "", "DWG Files (*.dwg)")
        if filename:
            # Convert path separators to avoid CAD recognition errors
            filename = os.path.normpath(filename)
            self.txt_file.setText(filename)
            self.update_buttons_state(False)

    def on_open_clicked(self):
        if not self.mech or not self.mech.zwm_db:
            QMessageBox.warning(self, "Warning", "Please connect or start the ZWCAD MFG first")
            return
            
        filepath = self.txt_file.text()
        if not filepath:
            QMessageBox.information(self, "Info", "Input file name is empty, will open the current active document")
            
        try:
            self.mech.close() # Close any potentially open connection first
            self.mech.open_file(filepath)
            self.update_buttons_state(True)
            QMessageBox.information(self, "Success", "Successfully connected to the drawing!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open drawing: {e}")

    def on_title_clicked(self):
        if not self.mech or not self.mech.zwm_db:
            QMessageBox.warning(self, "Warning", "Please connect or start the ZWCAD MFG first")
            return
        try:
            title = self.mech.get_title()
            if title:
                dialog = TitleDialog(self.mech, self)
                dialog.exec_()
            else:
                QMessageBox.warning(self, "Warning", "Title block object not found")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to get title block: {e}")

    def on_bom_clicked(self):
        if not self.mech or not self.mech.zwm_db:
            QMessageBox.warning(self, "Warning", "Please connect or start the ZWCAD MFG first")
            return
        try:
            bom = self.mech.get_bom()
            if bom:
                dialog = BomDialog(self.mech, self)
                dialog.exec_()
            else:
                QMessageBox.warning(self, "Warning", "BOM object not found")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to get BOM: {e}")

    def on_frame_clicked(self):
        if not self.mech or not self.mech.zwm_db:
            QMessageBox.warning(self, "Warning", "Please connect or start the ZWCAD MFG first")
            return
        try:
            frame = self.mech.get_frame()
            if frame:
                dialog = FrameDialog(self.mech, self)
                dialog.exec_()
            else:
                QMessageBox.warning(self, "Warning", "Drawing frame object not found")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to get drawing frame: {e}")

    def on_save_clicked(self):
        if not self.mech or not self.mech.zwm_db:
            QMessageBox.warning(self, "Warning", "Please connect or start the ZWCAD MFG first")
            return
        try:
            self.mech.save()
            QMessageBox.information(self, "Success", "Saved successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Save failed: {e}")

    def on_close_clicked(self):
        if not self.mech or not self.mech.zwm_db:
            QMessageBox.warning(self, "Warning", "Please connect or start the ZWCAD MFG first")
            return
        try:
            self.mech.close()
            self.update_buttons_state(False)
            QMessageBox.information(self, "Success", "Drawing connection closed!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Close failed: {e}")

    def on_title_edit_clicked(self):
        if self.mech and self.mech.zwm_db:
            self.mech.title_edit()
        else:
            QMessageBox.warning(self, "Warning", "Please connect or start the ZWCAD MFG first")

    def on_mxb_edit_clicked(self):
        if self.mech and self.mech.zwm_db:
            self.mech.total_bom_edit()
        else:
            QMessageBox.warning(self, "Warning", "Please connect or start the ZWCAD MFG first")

    def on_frame_edit_clicked(self):
        if self.mech and self.mech.zwm_db:
            self.mech.frame_edit()
        else:
            QMessageBox.warning(self, "Warning", "Please connect or start the ZWCAD MFG first")

    def on_fjl_edit_clicked(self):
        if self.mech and self.mech.zwm_db:
            self.mech.fjl_edit()
        else:
            QMessageBox.warning(self, "Warning", "Please connect or start the ZWCAD MFG first")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
