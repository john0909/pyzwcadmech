import os
import logging
import comtypes
import comtypes.client

logger = logging.getLogger(__name__)

# Try to load ZwmToolKit.tlb
try:
    import glob
    import comtypes.client
    
    # Prioritize searching for ZwmToolKit.tlb in the standard installation directory
    found_tlb = False
    for pattern in ("ZwmToolKit*.tlb", "ZwmToolKit.tlb"):
        # Try to find it in common ZWSOFT installation paths
        search_pattern = os.path.join(
            r"C:\Program Files\ZWSOFT\ZWCAD Mechanical*",
            "Zwcadm",
            pattern
        )
        tlibs = glob.glob(search_pattern)
        if tlibs:
            # If multiple versions are found, default to the latest version (sorted by path)
            tlb_path = sorted(tlibs, reverse=True)[0]
            comtypes.client.GetModule(tlb_path)
            found_tlb = True
            break
            
    # If not found in standard paths, try the current working directory or script directory
    if not found_tlb:
        tlb_path = os.path.abspath("ZwmToolKit.tlb")
        if not os.path.exists(tlb_path):
            tlb_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "ZwmToolKit.tlb"))
        
        if os.path.exists(tlb_path):
            comtypes.client.GetModule(tlb_path)
            found_tlb = True
            
    if found_tlb:
        import comtypes.gen.ZwmToolKitLib as ZWM
    else:
        logger.warning("ZwmToolKit.tlb not found in standard paths or current directory.")
        ZWM = None
except Exception as e:
    logger.error(f"Failed to load ZwmToolKit.tlb: {e}")
    ZWM = None


# ==========================================
# Interface Wrapper Classes
# ==========================================

class ZwmBomRow:
    """Wrapper for IBomRow interface"""
    def __init__(self, com_obj):
        self._obj = com_obj

    def get_item_count(self):
        count = comtypes.c_long()
        self._obj.GetItemCount(comtypes.byref(count))
        return count.value

    def get_item(self, index):
        label = comtypes.BSTR()
        name = comtypes.BSTR()
        value = comtypes.BSTR()
        self._obj.GetItem(index, comtypes.byref(label), comtypes.byref(name), comtypes.byref(value))
        return label.value or "", name.value or "", value.value or ""

    def set_item(self, key, value):
        self._obj.SetItem(key, value)


class ZwmBom:
    """Wrapper for IBom interface"""
    def __init__(self, com_obj):
        self._obj = com_obj

    def get_item_count(self):
        count = comtypes.c_long()
        self._obj.GetItemCount(comtypes.byref(count))
        return count.value

    def get_item(self, index):
        if ZWM is not None:
            row = comtypes.POINTER(ZWM.IBomRow)()
            self._obj.GetItem(index, comtypes.byref(row))
            return ZwmBomRow(row) if row else None
        return None

    def set_item(self, index, bom_row):
        self._obj.SetItem(index, bom_row._obj)

    def add_item(self, bom_row):
        self._obj.AddItem(bom_row._obj)

    def insert_item(self, index, bom_row):
        self._obj.InsertItem(index, bom_row._obj)

    def delete_item(self, index):
        self._obj.DeleteItem(index)

    def create_bom_row(self):
        if ZWM is not None:
            row = comtypes.POINTER(ZWM.IBomRow)()
            self._obj.CreateBomRow(comtypes.byref(row))
            return ZwmBomRow(row) if row else None
        return None


class ZwmTitle:
    """Wrapper for ITitle interface"""
    def __init__(self, com_obj):
        self._obj = com_obj

    def get_item_count(self):
        count = comtypes.c_long()
        self._obj.GetItemCount(comtypes.byref(count))
        return count.value

    def get_item(self, index):
        label = comtypes.BSTR()
        name = comtypes.BSTR()
        value = comtypes.BSTR()
        self._obj.GetItem(index, comtypes.byref(label), comtypes.byref(name), comtypes.byref(value))
        return label.value or "", name.value or "", value.value or ""

    def set_item(self, key, value):
        self._obj.SetItem(key, value)


class ZwmFrame:
    """Wrapper for IFrame interface"""
    def __init__(self, com_obj):
        self._obj = com_obj

    @property
    def width(self): return self._obj.Width
    @width.setter
    def width(self, val): self._obj.Width = val

    @property
    def height(self): return self._obj.Height
    @height.setter
    def height(self, val): self._obj.Height = val

    @property
    def std_name(self): return self._obj.StdName
    @std_name.setter
    def std_name(self, val): self._obj.StdName = val

    @property
    def frame_size_name(self): return self._obj.FrameSizeName
    @frame_size_name.setter
    def frame_size_name(self, val): self._obj.FrameSizeName = val

    @property
    def frame_style_name(self): return self._obj.FrameStyleName
    @frame_style_name.setter
    def frame_style_name(self, val): self._obj.FrameStyleName = val

    @property
    def orientation(self): return self._obj.Orientation
    @orientation.setter
    def orientation(self, val): self._obj.Orientation = val

    @property
    def title_style_name(self): return self._obj.TitleStyleName
    @title_style_name.setter
    def title_style_name(self, val): self._obj.TitleStyleName = val

    @property
    def bom_style_name(self): return self._obj.BomStyleName
    @bom_style_name.setter
    def bom_style_name(self, val): self._obj.BomStyleName = val

    @property
    def dhl_style_name(self): return self._obj.DhlStyleName
    @dhl_style_name.setter
    def dhl_style_name(self, val): self._obj.DhlStyleName = val

    @property
    def fjl_style_name(self): return self._obj.FjlStyleName
    @fjl_style_name.setter
    def fjl_style_name(self, val): self._obj.FjlStyleName = val

    @property
    def csl_style_name(self): return self._obj.CslStyleName
    @csl_style_name.setter
    def csl_style_name(self, val): self._obj.CslStyleName = val

    @property
    def ggl_style_name(self): return self._obj.GglStyleName
    @ggl_style_name.setter
    def ggl_style_name(self, val): self._obj.GglStyleName = val

    @property
    def have_dhl(self): return self._obj.HaveDHL
    @have_dhl.setter
    def have_dhl(self, val): self._obj.HaveDHL = val

    @property
    def have_fjl(self): return self._obj.HaveFJL
    @have_fjl.setter
    def have_fjl(self, val): self._obj.HaveFJL = val

    @property
    def have_btl(self): return self._obj.HaveBTL
    @have_btl.setter
    def have_btl(self, val): self._obj.HaveBTL = val

    @property
    def have_csl(self): return self._obj.HaveCSL
    @have_csl.setter
    def have_csl(self, val): self._obj.HaveCSL = val

    @property
    def have_ggl(self): return self._obj.HaveGGL
    @have_ggl.setter
    def have_ggl(self, val): self._obj.HaveGGL = val

    @property
    def scale1(self): return self._obj.Scale1
    @scale1.setter
    def scale1(self, val): self._obj.Scale1 = val

    @property
    def scale2(self): return self._obj.Scale2
    @scale2.setter
    def scale2(self, val): self._obj.Scale2 = val


class ZwmDb:
    """Wrapper for IZwmDb interface"""
    def __init__(self, com_obj):
        self._obj = com_obj

    def open_file(self, filepath):
        self._obj.OpenFile(filepath)

    def save(self, flag=33):
        self._obj.Save(flag)

    def close(self):
        self._obj.Close()

    def switch_frame(self, frame_name):
        self._obj.SwitchFrame(frame_name)

    def get_title(self):
        if ZWM is not None:
            title = comtypes.POINTER(ZWM.ITitle)()
            self._obj.GetTitle(comtypes.byref(title))
            return ZwmTitle(title) if title else None
        return None

    def get_bom(self):
        if ZWM is not None:
            bom = comtypes.POINTER(ZWM.IBom)()
            self._obj.GetBom(comtypes.byref(bom))
            return ZwmBom(bom) if bom else None
        return None

    def get_frame(self):
        if ZWM is not None:
            # paramflags is (2, 'ppIFrame'), which indicates an out parameter.
            # comtypes automatically handles out parameters, so byref is not needed.
            frame = self._obj.GetFrame()
            return ZwmFrame(frame) if frame else None
        return None

    def get_next_frm_name(self):
        if ZWM is not None:
            # paramflags is ((2, 'ppIFrame'), (2, 'bstrFrameName'))
            frame, name = self._obj.GetNextFrmName()
            return ZwmFrame(frame) if frame else None, name or ""
        return None, ""

    def refresh_title(self):
        self._obj.RefreshTitle()

    def refresh_bom(self):
        self._obj.RefreshBom()

    def refresh_frame(self):
        self._obj.RefreshFrame()

    def build_frame(self, switch_type=511):
        self._obj.BuildFrame(switch_type)

    def get_frame_count(self):
        # paramflags is ((2, 'nCount'),)
        count = self._obj.GetFrameCount()
        return count

    def get_frame_name(self, index):
        # paramflags is ((1, 'nIndex'), (2, 'bstrFrameName'))
        name = self._obj.GetFrameName(index)
        return name or ""

    def get_frame_name2(self, point):
        # paramflags is ((1, 'point'), (2, 'bstrFrameName'))
        name = self._obj.GetFrameName2(point)
        return name or ""

    def cad_environment_init(self, std_name):
        self._obj.CadEnvironmentInit(std_name)

    def frame_init(self, std_name):
        self._obj.FrameInit(std_name)

    def frame_edit(self):
        self._obj.FrameEdit()

    def title_edit(self):
        self._obj.TitleEdit()

    def csl_edit(self):
        self._obj.CslEdit()

    def fjl_edit(self):
        self._obj.FjlEdit()

    def total_bom_edit(self):
        self._obj.TotalBomEdit()

    def get_balloon(self, text=""):
        # paramflags is ((2, 'varBalloon'), (17, 'strBalloonText', ''))
        balloon = self._obj.GetBalloon(text)
        return balloon


class ZwmApp:
    """Wrapper for IZwmApp interface"""
    def __init__(self, com_obj):
        self._obj = com_obj

    def get_db(self):
        return ZwmDb(self._obj.GetDb())

    def get_cad_path(self):
        return self._obj.GetCadPath()

    def get_zwm_path(self):
        return self._obj.GetZwmPath()

    def get_style_path(self):
        return self._obj.GetStylePath()

    def get_version(self):
        return self._obj.GetVersion()

    def send_command(self, cmd):
        self._obj.SendCommand(cmd)

    def get_about(self):
        return self._obj.GetAbout()

    @property
    def cmd_line(self):
        return self._obj.CmdLine

    def open_doc(self, file_path):
        self._obj.OpenDoc(file_path)

    def new_doc(self, file_path):
        self._obj.NewDoc(file_path)

    def new_named_doc(self, file_path, template):
        self._obj.NewNamedDoc(file_path, template)


# ==========================================
# Core Entry Class
# ==========================================

class ZwCADMech(object):
    """Main ZwCAD Mechanical Automation entry class"""
    
    def __init__(self, cad_app=None):
        self._cad_app = cad_app
        self._zwm_app = None
        self._zwm_db = None

    @property
    def app(self):
        """Returns the active ZWCAD.Application instance"""
        if self._cad_app is None:
            try:
                self._cad_app = comtypes.client.GetActiveObject('ZWCAD.Application', dynamic=True)
            except WindowsError:
                self._cad_app = comtypes.client.CreateObject('ZWCAD.Application', dynamic=True)
                self._cad_app.Visible = True
        return self._cad_app

    @property
    def zwm_app(self):
        """Returns the ZwmApp wrapper instance"""
        if self._zwm_app is None:
            try:
                app_dispatch = self.app.GetInterfaceObject("ZwmToolKit.ZwmApp")
                if ZWM is not None:
                    raw_obj = app_dispatch.QueryInterface(ZWM.IZwmApp)
                else:
                    raw_obj = app_dispatch
                self._zwm_app = ZwmApp(raw_obj)
            except Exception as e:
                logger.error(f"Failed to get ZwmToolKit.ZwmApp: {e}")
                raise RuntimeError("Failed to create ZwmToolKit.ZwmApp. Please check if ZWCAD Mechanical is installed and running.")
        return self._zwm_app

    @property
    def zwm_db(self):
        """Returns the ZwmDb wrapper instance"""
        if self._zwm_db is None:
            try:
                self._zwm_db = self.zwm_app.get_db()
            except Exception as e:
                logger.error(f"Failed to get ZwmDb: {e}")
                raise RuntimeError("Failed to get ZwmDb object.")
        return self._zwm_db

    # Shortcut methods provided for backward compatibility
    def open_file(self, filepath):
        self.zwm_db.open_file(filepath)

    def close(self):
        if self._zwm_db is not None:
            self._zwm_db.close()
            self._zwm_db = None

    def save(self, flag=33):
        self.zwm_db.save(flag)

    def get_title(self):
        return self.zwm_db.get_title()

    def get_bom(self):
        return self.zwm_db.get_bom()

    def get_frame(self):
        return self.zwm_db.get_frame()

    def title_edit(self):
        self.zwm_db.title_edit()

    def total_bom_edit(self):
        self.zwm_db.total_bom_edit()

    def frame_edit(self):
        self.zwm_db.frame_edit()

    def fjl_edit(self):
        self.zwm_db.fjl_edit()