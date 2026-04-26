import sys
import os

# 确保能找到 pyzwcadmech 库
sys.path.insert(0, os.path.abspath(os.path.dirname(os.path.dirname(__file__))))

try:
    from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, 
                                 QHBoxLayout, QPushButton, QLabel, QLineEdit, 
                                 QMessageBox, QGridLayout, QComboBox)
except ImportError:
    print("未安装 PyQt5，请先运行: pip install PyQt5")
    sys.exit(1)

from pyzwcadmech import ZwCADMech

class CreateFrameWidget(QWidget):
    def __init__(self, mech, parent=None):
        super().__init__(parent)
        self.mech = mech
        self.frame_obj = None
        self.current_frame_name = ""
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("创建图框示例")
        self.resize(500, 600)
        layout = QVBoxLayout(self)
        
        # 图幅选择 (只读，用于显示新建的图幅名称)
        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("当前图幅:"))
        self.combo_frame = QComboBox()
        self.combo_frame.setEnabled(False) # 只能显示新建的名称
        top_layout.addWidget(self.combo_frame)
        layout.addLayout(top_layout)
        
        # 属性网格
        grid = QGridLayout()
        self.fields = {}
        labels = [
            ("std_name", "标准", "GB"), ("frame_size_name", "图幅", "A4"), ("frame_style_name", "图框样式", "分区图框"),
            ("orientation", "方向", "landscape"), ("width", "宽度", "0"), ("height", "高度", "0"),
            ("title_style_name", "标题栏", "标题栏1"), ("bom_style_name", "明细表", "明细表1"), ("dhl_style_name", "代号栏", "图样代号"),
            ("fjl_style_name", "附加栏", "附加栏"), ("csl_style_name", "参数栏", "包络环面蜗杆"), ("ggl_style_name", "更改栏", ""),
            ("have_dhl", "有代号栏", "0"), ("have_fjl", "有附加栏", "0"), ("have_btl", "有标题栏", "1"),
            ("have_csl", "有参数栏", "0"), ("have_ggl", "有更改栏", "1"), ("scale1", "比例1", "1"), ("scale2", "比例2", "1")
        ]
        
        for i, (key, label, default_value) in enumerate(labels):
            grid.addWidget(QLabel(label + ":"), i, 0)
            le = QLineEdit()
            le.setText(default_value)
            self.fields[key] = le
            grid.addWidget(le, i, 1)
            
        layout.addLayout(grid)
        
        # 按钮
        btn_layout = QHBoxLayout()
        self.btn_create = QPushButton("新建图框")
        
        self.btn_create.clicked.connect(self.on_create)
        
        btn_layout.addWidget(self.btn_create)
        layout.addLayout(btn_layout)

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

    def on_create(self):
        if not self.mech or not self.mech.zwm_db: 
            QMessageBox.warning(self, "提示", "未连接到中望机械，请检查。")
            return
            
        # 1. 获取下一个新建图幅的对象和名称
        self.frame_obj, name = self.mech.zwm_db.get_next_frm_name()
        if not name:
            QMessageBox.warning(self, "错误", "获取新图框名称失败。")
            return
            
        self.current_frame_name = name
        
        # 2. 切换到该新图框名称
        self.mech.zwm_db.switch_frame(self.current_frame_name)
        
        # 3. 获取当前激活的图框对象以进行属性设置
        self.frame_obj = self.mech.zwm_db.get_frame()
        if not self.frame_obj:
            QMessageBox.warning(self, "错误", "无法获取图框对象以进行编辑。")
            return
        
        # 4. 将界面上的值写入图框对象
        self.write_to_frame()
        
        # 更新界面显示
        self.combo_frame.addItem(self.current_frame_name)
        self.combo_frame.setCurrentText(self.current_frame_name)
        
        # 5. 构建图框
        try:
            self.mech.zwm_db.build_frame(511)
            QMessageBox.information(self, "提示", f"成功新建图幅: {self.current_frame_name}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"构建图框失败: {e}")

def main():
    app = QApplication(sys.argv)
    
    print("正在连接中望机械...")
    try:
        mech = ZwCADMech()
        if not mech.zwm_app or not mech.zwm_db:
            print("无法获取 ZwmApp 或 ZwmDb，请检查中望机械状态。")
            sys.exit(1)
            
        # 绑定到当前活动文档 (这一步必不可少，否则会报 COMError 服务器意外情况)
        mech.open_file("")
        
        print("成功连接中望机械！")
    except Exception as e:
        print(f"连接中望机械失败: {e}")
        print("请确保中望机械已启动。")
        sys.exit(1)
        
    window = CreateFrameWidget(mech)
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
