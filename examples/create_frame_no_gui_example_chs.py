import sys
import os

# 确保能找到 pyzwcadmech 库
sys.path.insert(0, os.path.abspath(os.path.dirname(os.path.dirname(__file__))))

from pyzwcadmech import ZwCADMech

def create_frame_no_gui_example():
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

    print(f"CAD 路径: {mech.zwm_app.get_cad_path()}")
    print(f"ZWM 路径: {mech.zwm_app.get_zwm_path()}")
    print("-" * 40)

    print("开始从无到有创建新图框...")
    
    # 1. 获取下一个新建图幅的对象和名称
    frame_obj, new_frame_name = mech.zwm_db.get_next_frm_name()
    if not new_frame_name:
        print("获取新图框名称失败。")
        return
        
    print(f"获取到新图框名称: {new_frame_name}")

    # 2. 切换到该新图框名称
    mech.zwm_db.switch_frame(new_frame_name)

    # 3. 获取当前激活的图框对象以进行属性设置
    frame_obj = mech.zwm_db.get_frame()
    if not frame_obj:
        print("无法获取图框对象以进行编辑。")
        return

    # 4. 设置图框的相关属性 (参考 create_frame_pyqt_example_chs.py 里的默认值)
    print("正在设置图框属性...")
    
    # 属性字典 (完全匹配 GUI 版本的默认值)
    fields = {
        "std_name": "GB",
        "frame_size_name": "A4",
        "frame_style_name": "分区图框",
        "orientation": "landscape",
        "width": "0",
        "height": "0",
        "title_style_name": "标题栏1",
        "bom_style_name": "明细表1",
        "dhl_style_name": "图样代号",
        "fjl_style_name": "附加栏",
        "csl_style_name": "包络环面蜗杆",
        "ggl_style_name": "",
        "have_dhl": "0",
        "have_fjl": "0",
        "have_btl": "1",
        "have_csl": "0",
        "have_ggl": "1",
        "scale1": "1",
        "scale2": "1"
    }

    # 将属性值写入图框对象
    for key, val_str in fields.items():
        if key in ["width", "height", "have_dhl", "have_fjl", "have_btl", "have_csl", "have_ggl"]:
            try:
                setattr(frame_obj, key, int(val_str))
            except ValueError:
                pass
        else:
            try:
                setattr(frame_obj, key, val_str)
            except Exception:
                pass

    print("属性设置成功。")

    # 5. 构建图框
    print("正在构建图框...")
    try:
        mech.zwm_db.build_frame(511)
        print(f"成功新建图幅: {new_frame_name}")
        print("图框构建成功！可以在中望机械中查看新建的图框。")
    except Exception as e:
        print(f"构建图框失败: {e}")

if __name__ == '__main__':
    create_frame_no_gui_example()