# pyzwcadmech 使用说明文档

`pyzwcadmech` 是一个基于 Python 的中望机械（ZWCAD Mechanical）二次开发接口封装库。它通过 COM 接口（`ZwmToolKit.tlb`）与中望机械进行通信，并将底层的 COM 指针操作封装为符合 Python 习惯的面向对象 API，允许开发者使用 Python 脚本轻松读取和修改图纸中的标题栏、明细表、图框等机械专属对象。

## 1. 环境要求

- **操作系统**：Windows
- **软件依赖**：已安装并授权的 **中望CAD机械版**。
- **Python 环境**：Python 3.x
- **第三方库**：`comtypes`（用于 COM 组件通信）
  ```bash
  pip install comtypes
  ```
- **核心文件**：确保 `ZwmToolKit.tlb` 文件位于项目根目录或 `pyzwcadmech` 模块同级目录中。

## 2. 目录结构

确保你的项目目录包含以下结构：

```text
你的项目目录/
├── ZwmToolKit.tlb       # 中望机械提供的 COM 类型库文件
├── pyzwcadmech/         # 核心封装库
│   ├── __init__.py
│   └── api.py
└── hello_mech.py        # 你的调用脚本
```

## 3. 快速入门

以下是一个简单的示例，演示如何连接到中望机械并读取当前图纸的标题栏数据

```python
from pyzwcadmech import ZwCADMech

def main():
    # 1. 初始化并连接到中望机械
    mech = ZwCADMech()
    
    # 2. 连接当前活动图纸
    mech.open_file("")
    
    # 3. 获取并读取标题栏对象
    title = mech.get_title()
    if title:
        print("成功获取标题栏对象！")
        count = title.get_item_count()
        for i in range(count):
            label, name, value = title.get_item(i)
            print(f"{label}: {value}")
            
        # 修改标题栏属性
        # title.set_item("设计", "张三")
        # mech.zwm_db.refresh_title() # 刷新图纸显示

    # 4. 交互式编辑（会弹出中望机械的编辑对话框）
    # mech.title_edit()
    
    # 5. 保存并断开连接
    # mech.save()
    mech.close()

if __name__ == "__main__":
    main()
```

## 4. API 参考

### 4.1 核心入口类：`ZwCADMech`

`ZwCADMech` 是整个库的入口，负责管理 CAD 进程和机械模块数据库。

#### 初始化
```python
mech = ZwCADMech(cad_app=None)
```
- `cad_app`: 可选参数。如果已有 `ZWCAD.Application` 实例可传入，否则自动获取或启动。

#### 核心属性
- `mech.app`: 获取原生的 `ZWCAD.Application` COM 对象。
- `mech.zwm_app`: 返回 `ZwmApp` 封装实例。
- `mech.zwm_db`: 返回 `ZwmDb` 封装实例。

#### 快捷方法（代理到 `ZwmDb`）
- `open_file(filepath: str)`: 打开指定的 DWG 文件。传空字符串 `""` 表示当前活动文档。
- `save(flag: int = 33)`: 保存当前图纸。
- `close()`: 关闭数据库连接。
- `get_title() -> ZwmTitle`: 获取标题栏对象。
- `get_bom() -> ZwmBom`: 获取明细表对象。
- `get_frame() -> ZwmFrame`: 获取图框对象。
- 交互编辑：`title_edit()`, `total_bom_edit()`, `frame_edit()`, `fjl_edit()`

---

### 4.2 标题栏对象：`ZwmTitle`

- `get_item_count() -> int`: 获取标题栏属性的数量。
- `get_item(index: int) -> tuple(str, str, str)`: 根据索引获取属性，返回 `(label, name, value)`。
- `set_item(key: str, value: str)`: 根据标签名或属性名设置对应的值。

---

### 4.3 明细表对象：`ZwmBom` & `ZwmBomRow`

#### `ZwmBom` (明细表整体)
- `get_item_count() -> int`: 获取明细表的行数。
- `get_item(index: int) -> ZwmBomRow`: 获取指定索引的明细表行对象。
- `set_item(index: int, bom_row: ZwmBomRow)`: 更新指定行的数据。
- `add_item(bom_row: ZwmBomRow)`: 在末尾添加一行。
- `insert_item(index: int, bom_row: ZwmBomRow)`: 在指定位置插入一行。
- `delete_item(index: int)`: 删除指定行。
- `create_bom_row() -> ZwmBomRow`: 创建一个空白的明细表行对象，用于后续添加。

#### `ZwmBomRow` (明细表单行)
- `get_item_count() -> int`: 获取该行的列数（属性数）。
- `get_item(index: int) -> tuple(str, str, str)`: 获取指定列的数据，返回 `(label, name, value)`。
- `set_item(key: str, value: str)`: 设置指定列的值。

---

### 4.4 图框对象：`ZwmFrame`

图框对象的属性全部封装为 Python 的 `@property`，可直接读写：

- **尺寸与标准**：`width` (宽), `height` (高), `std_name` (标准名, 如 "GB"), `frame_size_name` (图幅, 如 "A3"), `orientation` (方向), `scale1`, `scale2` (比例)。
- **样式名称**：`frame_style_name`, `title_style_name`, `bom_style_name`, `dhl_style_name`, `fjl_style_name`, `csl_style_name`, `ggl_style_name`。
- **包含元素 (1=有, 0=无)**：`have_dhl` (代号栏), `have_fjl` (附加栏), `have_btl` (标题栏), `have_csl` (参数栏), `have_ggl` (更改栏)。

*修改属性后，需调用 `mech.zwm_db.build_frame(511)` 或 `refresh_frame()` 使更改在图纸中生效。*

---

### 4.5 数据库操作：`ZwmDb`

除了被 `ZwCADMech` 代理的基础方法外，`ZwmDb` 还提供：
- `switch_frame(frame_name: str)`: 切换当前操作的图幅。
- `get_frame_count() -> int`: 获取图纸中图幅的数量。
- `get_frame_name(index: int) -> str`: 获取指定索引的图幅名称。
- `get_next_frm_name() -> tuple(ZwmFrame, str)`: 获取下一个新建图幅的对象和名称。
- `refresh_title()`, `refresh_bom()`, `refresh_frame()`: 刷新对应对象在图纸中的显示。
- `build_frame(switch_type: int = 511)`: 重建图幅块及关联块。

## 5. 常见问题 (FAQ)
**Q:运行报错`ModuleNotFoundError: No module named 'pyzwcadmech'`**
A:原因为pyzwcadmech库未安装，可使用`pip install pyzwcadmech`来进行安装

**Q: 运行报错 `Failed to load ZwmToolKit.tlb` 或 `ModuleNotFoundError`**
A: 请确保 `ZwmToolKit.tlb` 文件存在于当前运行脚本的目录下，或者放在 `pyzwcadmech` 文件夹旁边。同时确保已安装 `comtypes` 库。

**Q: 运行报错 `无法创建 ZwmToolKit.ZwmApp，请检查中望机械是否正常安装并启动`**
A: 这通常是因为：
1. 电脑上安装的是普通版 ZWCAD，而不是**机械版**（ZWCAD Mechanical）。
2. 中望机械的 COM 组件未正确注册。可以尝试以管理员身份运行一次中望机械。

**Q: 获取标题栏或明细表返回空值或报错**
A: 请确保在调用 `get_title()` 或 `get_bom()` 之前，已经成功调用了 `mech.open_file("")` 来绑定图纸数据库。同时，确保当前图纸中确实存在中望机械的标题栏或明细表实体。
