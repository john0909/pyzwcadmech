import sys
import os

# 确保能找到 pyzwcadmech 库
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

try:
    from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                                 QHBoxLayout, QPushButton, QLabel, QLineEdit, 
                                 QFileDialog, QMessageBox, QGroupBox, QGridLayout,
                                 QDialog, QTableWidget, QTableWidgetItem, QHeaderView,
                                 QComboBox)
except ImportError:
    print("未安装 PyQt5，请先运行: pip install PyQt5")
    sys.exit(1)

from pyzwcadmech import ZwCADMech

import comtypes
import comtypes.gen.ZwmToolKitLib as ZWM

class TitleDialog(QDialog):
    def __init__(self, mech, parent=None):
        super().__init__(parent)
        self.mech = mech
        self.title_obj = mech.get_title()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("标题栏")
        self.resize(450, 400)
        layout = QVBoxLayout(self)
        
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Label", "Value"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.table)
        
        btn_layout = QHBoxLayout()
        self.btn_read = QPushButton("读取")
        self.btn_write = QPushButton("写入")
        self.btn_refresh = QPushButton("刷新")
        
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
        QMessageBox.information(self, "提示", "写入成功")

    def on_refresh(self):
        if self.mech and self.mech.zwm_db:
            self.mech.zwm_db.refresh_title()
            QMessageBox.information(self, "提示", "刷新成功")


class BomDialog(QDialog):
    def __init__(self, mech, parent=None):
        super().__init__(parent)
        self.mech = mech
        self.bom_obj = mech.get_bom()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("明细表")
        self.resize(600, 400)
        layout = QVBoxLayout(self)
        
        self.table = QTableWidget()
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.table)
        
        btn_layout = QHBoxLayout()
        self.btn_read = QPushButton("读取")
        self.btn_write = QPushButton("写入")
        self.btn_refresh = QPushButton("刷新")
        self.btn_add = QPushButton("添加")
        self.btn_delete = QPushButton("删除")
        
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
        QMessageBox.information(self, "提示", "写入成功")

    def on_refresh(self):
        if self.mech and self.mech.zwm_db:
            self.mech.zwm_db.refresh_bom()
            QMessageBox.information(self, "提示", "刷新成功")

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
            QMessageBox.warning(self, "提示", "请先选择要删除的行")
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
        self.setWindowTitle("图框")
        self.resize(500, 600)
        layout = QVBoxLayout(self)
        
        # 图幅选择
        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("选择图幅:"))
        self.combo_frame = QComboBox()
        self.combo_frame.currentTextChanged.connect(self.on_frame_changed)
        top_layout.addWidget(self.combo_frame)
        layout.addLayout(top_layout)
        
        # 属性网格
        grid = QGridLayout()
        self.fields = {}
        labels = [
            ("std_name", "标准"), ("frame_size_name", "图幅"), ("frame_style_name", "图框样式"),
            ("orientation", "方向"), ("width", "宽度"), ("height", "高度"),
            ("title_style_name", "标题栏"), ("bom_style_name", "明细表"), ("dhl_style_name", "代号栏"),
            ("fjl_style_name", "附加栏"), ("csl_style_name", "参数栏"), ("ggl_style_name", "更改栏"),
            ("have_dhl", "有代号栏"), ("have_fjl", "有附加栏"), ("have_btl", "有标题栏"),
            ("have_csl", "有参数栏"), ("have_ggl", "有更改栏"), ("scale1", "比例1"), ("scale2", "比例2")
        ]
        
        for i, (key, label) in enumerate(labels):
            grid.addWidget(QLabel(label + ":"), i, 0)
            le = QLineEdit()
            self.fields[key] = le
            grid.addWidget(le, i, 1)
            
        layout.addLayout(grid)
        
        # 按钮
        btn_layout = QHBoxLayout()
        self.btn_read = QPushButton("读取")
        self.btn_refresh = QPushButton("刷新字典")
        self.btn_rebuild = QPushButton("重建图幅")
        self.btn_create = QPushButton("新建图幅")
        
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
        QMessageBox.information(self, "提示", "刷新成功")

    def on_rebuild(self):
        if not self.mech or not self.mech.zwm_db: return
        
        self.mech.zwm_db.switch_frame(self.current_frame_name)
        self.frame_obj = self.mech.zwm_db.get_frame()
        
        self.write_to_frame()
        self.mech.zwm_db.build_frame(511)
        QMessageBox.information(self, "提示", "重建成功")

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
        QMessageBox.information(self, "提示", "新建成功")



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

        # 1. 启动/连接区域
        conn_group = QGroupBox("连接中望机械")
        conn_layout = QGridLayout()
        
        self.btn_start = QPushButton("启动/连接机械软件")
        self.btn_start.clicked.connect(self.on_start_clicked)
        
        self.txt_cad_path = QLineEdit()
        self.txt_cad_path.setReadOnly(True)
        self.txt_zwm_path = QLineEdit()
        self.txt_zwm_path.setReadOnly(True)
        self.txt_version = QLineEdit()
        self.txt_version.setReadOnly(True)

        conn_layout.addWidget(self.btn_start, 0, 0, 1, 2)
        conn_layout.addWidget(QLabel("CAD 路径:"), 1, 0)
        conn_layout.addWidget(self.txt_cad_path, 1, 1)
        conn_layout.addWidget(QLabel("ZWM 路径:"), 2, 0)
        conn_layout.addWidget(self.txt_zwm_path, 2, 1)
        conn_layout.addWidget(QLabel("版本号:"), 3, 0)
        conn_layout.addWidget(self.txt_version, 3, 1)
        conn_group.setLayout(conn_layout)
        main_layout.addWidget(conn_group)

        # 2. 文件操作区域
        file_group = QGroupBox("文件操作")
        file_layout = QHBoxLayout()
        
        self.txt_file = QLineEdit()
        self.btn_browse = QPushButton("浏览...")
        self.btn_browse.clicked.connect(self.on_browse_clicked)
        self.btn_open = QPushButton("打开图纸")
        self.btn_open.clicked.connect(self.on_open_clicked)
        
        file_layout.addWidget(QLabel("DWG文件:"))
        file_layout.addWidget(self.txt_file)
        file_layout.addWidget(self.btn_browse)
        file_layout.addWidget(self.btn_open)
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)

        # 3. 数据获取区域
        data_group = QGroupBox("数据获取与保存")
        data_layout = QHBoxLayout()
        
        self.btn_title = QPushButton("获取标题栏")
        self.btn_bom = QPushButton("获取明细表")
        self.btn_frame = QPushButton("获取图框")
        self.btn_save = QPushButton("保存")
        self.btn_close = QPushButton("关闭图纸")
        
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

        # 4. 交互编辑区域
        edit_group = QGroupBox("交互式编辑")
        edit_layout = QHBoxLayout()
        
        self.btn_title_edit = QPushButton("编辑标题栏")
        self.btn_mxb_edit = QPushButton("编辑明细表")
        self.btn_frame_edit = QPushButton("编辑图框")
        self.btn_fjl_edit = QPushButton("编辑附加栏")
        
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

        # 初始化按钮状态
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
            
            # 触发连接
            app = self.mech.app
            zwm_app = self.mech.zwm_app
            zwm_db = self.mech.zwm_db
            
            if zwm_app:
                try:
                    # 获取机械模块的路径信息
                    self.txt_cad_path.setText(zwm_app.get_cad_path())
                    self.txt_zwm_path.setText(zwm_app.get_zwm_path())
                    self.txt_version.setText(zwm_app.get_version())
                    QMessageBox.information(self, "成功", "成功连接到中望机械！")
                except Exception as e:
                    QMessageBox.warning(self, "警告", f"获取路径信息失败: {e}")
            else:
                QMessageBox.critical(self, "错误", "创建 ZwmToolKit.ZwmApp 失败！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"连接失败: {e}")

    def on_browse_clicked(self):
        filename, _ = QFileDialog.getOpenFileName(self, "选择 DWG 文件", "", "DWG Files (*.dwg)")
        if filename:
            # 转换路径分隔符，避免 CAD 识别错误
            filename = os.path.normpath(filename)
            self.txt_file.setText(filename)
            self.update_buttons_state(False)

    def on_open_clicked(self):
        if not self.mech or not self.mech.zwm_db:
            QMessageBox.warning(self, "提示", "请首先连接或者启动机械软件")
            return
            
        filepath = self.txt_file.text()
        if not filepath:
            QMessageBox.information(self, "提示", "输入文件名称为空，将打开当前打开的活动文档")
            
        try:
            self.mech.close() # 先关闭可能打开的连接
            self.mech.open_file(filepath)
            self.update_buttons_state(True)
            QMessageBox.information(self, "成功", "图纸连接成功！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"打开图纸失败: {e}")

    def on_title_clicked(self):
        if not self.mech or not self.mech.zwm_db:
            QMessageBox.warning(self, "提示", "请首先连接或者启动机械软件")
            return
        try:
            title = self.mech.get_title()
            if title:
                dialog = TitleDialog(self.mech, self)
                dialog.exec_()
            else:
                QMessageBox.warning(self, "提示", "未找到标题栏对象")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"获取标题栏失败: {e}")

    def on_bom_clicked(self):
        if not self.mech or not self.mech.zwm_db:
            QMessageBox.warning(self, "提示", "请首先连接或者启动机械软件")
            return
        try:
            bom = self.mech.get_bom()
            if bom:
                dialog = BomDialog(self.mech, self)
                dialog.exec_()
            else:
                QMessageBox.warning(self, "提示", "未找到明细表对象")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"获取明细表失败: {e}")

    def on_frame_clicked(self):
        if not self.mech or not self.mech.zwm_db:
            QMessageBox.warning(self, "提示", "请首先连接或者启动机械软件")
            return
        try:
            frame = self.mech.get_frame()
            if frame:
                dialog = FrameDialog(self.mech, self)
                dialog.exec_()
            else:
                QMessageBox.warning(self, "提示", "未找到图框对象")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"获取图框失败: {e}")

    def on_save_clicked(self):
        if not self.mech or not self.mech.zwm_db:
            QMessageBox.warning(self, "提示", "请首先连接或者启动机械软件")
            return
        try:
            self.mech.save()
            QMessageBox.information(self, "成功", "保存成功！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存失败: {e}")

    def on_close_clicked(self):
        if not self.mech or not self.mech.zwm_db:
            QMessageBox.warning(self, "提示", "请首先连接或者启动机械软件")
            return
        try:
            self.mech.close()
            self.update_buttons_state(False)
            QMessageBox.information(self, "成功", "图纸已关闭连接！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"关闭失败: {e}")

    def on_title_edit_clicked(self):
        if self.mech and self.mech.zwm_db:
            self.mech.title_edit()
        else:
            QMessageBox.warning(self, "提示", "请首先连接或者启动机械软件")

    def on_mxb_edit_clicked(self):
        if self.mech and self.mech.zwm_db:
            self.mech.total_bom_edit()
        else:
            QMessageBox.warning(self, "提示", "请首先连接或者启动机械软件")

    def on_frame_edit_clicked(self):
        if self.mech and self.mech.zwm_db:
            self.mech.frame_edit()
        else:
            QMessageBox.warning(self, "提示", "请首先连接或者启动机械软件")

    def on_fjl_edit_clicked(self):
        if self.mech and self.mech.zwm_db:
            self.mech.fjl_edit()
        else:
            QMessageBox.warning(self, "提示", "请首先连接或者启动机械软件")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
