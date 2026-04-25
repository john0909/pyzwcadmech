[English](README_EN.md) | [中文](README_CN.md)

# pyzwcadmech Documentation

`pyzwcadmech` is a Python wrapper for the ZWCAD Mechanical COM API (`ZwmToolKit.tlb`). It encapsulates the underlying COM pointer operations into a Pythonic object-oriented API, allowing developers to easily read and modify mechanical-specific objects such as title blocks, BOMs (Bill of Materials), and drawing frames using Python scripts.

## 1. Requirements

- **OS**: Windows
- **Software Dependency**: A licensed installation of **ZWCAD Mechanical**.
- **Python Environment**: Python 3.x
- **Third-party Libraries**: `comtypes` (for COM component communication)
  ```bash
  pip install comtypes
  ```
- **Core Files**: The `ZwmToolKit.tlb` file will be automatically searched in the standard ZWCAD Mechanical installation path. If not found, ensure it is placed in the project root or the same directory as the `pyzwcadmech` module.

## 2. Directory Structure

Ensure your project directory has the following structure:

```text
your_project_directory/
├── ZwmToolKit.tlb       # (Optional) COM type library file provided by ZWCAD Mechanical
├── pyzwcadmech/         # Core wrapper library
│   ├── __init__.py
│   └── api.py
└── hello_mech.py        # Your calling script
```

## 3. Quick Start

Below is a simple example demonstrating how to connect to ZWCAD Mechanical and read the title block data of the current drawing:

```python
from pyzwcadmech import ZwCADMech

def main():
    # 1. Initialize and connect to ZWCAD Mechanical
    mech = ZwCADMech()
    
    # 2. Connect to the current active drawing
    mech.open_file("")
    
    # 3. Get and read the title block object
    title = mech.get_title()
    if title:
        print("Successfully obtained the title block object!")
        count = title.get_item_count()
        for i in range(count):
            label, name, value = title.get_item(i)
            print(f"{label}: {value}")
            
        # Modify title block properties
        # title.set_item("Designer", "John Doe")
        # mech.zwm_db.refresh_title() # Refresh the drawing display

    # 4. Interactive editing (pops up the ZWCAD Mechanical editing dialog)
    # mech.title_edit()
    
    # 5. Save and disconnect
    # mech.save()
    mech.close()

if __name__ == "__main__":
    main()
```

## 4. API Reference

### 4.1 Core Entry Class: `ZwCADMech`

`ZwCADMech` is the entry point of the entire library, responsible for managing the CAD process and the mechanical module database.

#### Initialization
```python
mech = ZwCADMech(cad_app=None)
```
- `cad_app`: Optional. If you already have a `ZWCAD.Application` instance, you can pass it in. Otherwise, it will be automatically obtained or started.

#### Core Properties
- `mech.app`: Returns the native `ZWCAD.Application` COM object.
- `mech.zwm_app`: Returns the `ZwmApp` wrapper instance.
- `mech.zwm_db`: Returns the `ZwmDb` wrapper instance.

#### Shortcut Methods (Delegated to `ZwmDb`)
- `open_file(filepath: str)`: Opens the specified DWG file. Passing an empty string `""` means the current active document.
- `save(flag: int = 33)`: Saves the current drawing.
- `close()`: Closes the database connection.
- `get_title() -> ZwmTitle`: Gets the title block object.
- `get_bom() -> ZwmBom`: Gets the BOM object.
- `get_frame() -> ZwmFrame`: Gets the drawing frame object.
- Interactive editing: `title_edit()`, `total_bom_edit()`, `frame_edit()`, `fjl_edit()`

---

### 4.2 Title Block Object: `ZwmTitle`

- `get_item_count() -> int`: Gets the number of title block properties.
- `get_item(index: int) -> tuple(str, str, str)`: Gets properties by index, returns `(label, name, value)`.
- `set_item(key: str, value: str)`: Sets the corresponding value by label or property name.

---

### 4.3 BOM Object: `ZwmBom` & `ZwmBomRow`

#### `ZwmBom` (Entire BOM)
- `get_item_count() -> int`: Gets the number of rows in the BOM.
- `get_item(index: int) -> ZwmBomRow`: Gets the BOM row object at the specified index.
- `set_item(index: int, bom_row: ZwmBomRow)`: Updates the data of the specified row.
- `add_item(bom_row: ZwmBomRow)`: Adds a row to the end.
- `insert_item(index: int, bom_row: ZwmBomRow)`: Inserts a row at the specified position.
- `delete_item(index: int)`: Deletes the specified row.
- `create_bom_row() -> ZwmBomRow`: Creates a blank BOM row object for subsequent addition.

#### `ZwmBomRow` (Single BOM Row)
- `get_item_count() -> int`: Gets the number of columns (properties) in the row.
- `get_item(index: int) -> tuple(str, str, str)`: Gets the data of the specified column, returns `(label, name, value)`.
- `set_item(key: str, value: str)`: Sets the value of the specified column.

---

### 4.4 Drawing Frame Object: `ZwmFrame`

All properties of the drawing frame object are encapsulated as Python `@property` and can be read and written directly:

- **Dimensions and Standards**: `width`, `height`, `std_name` (e.g., "GB"), `frame_size_name` (e.g., "A3"), `orientation`, `scale1`, `scale2`.
- **Style Names**: `frame_style_name`, `title_style_name`, `bom_style_name`, `dhl_style_name`, `fjl_style_name`, `csl_style_name`, `ggl_style_name`.
- **Included Elements (1=Yes, 0=No)**: `have_dhl` (Code Block), `have_fjl` (Additional Block), `have_btl` (Title Block), `have_csl` (Parameter Block), `have_ggl` (Modification Block).

*After modifying properties, you must call `mech.zwm_db.build_frame(511)` or `refresh_frame()` to apply the changes to the drawing.*

---

### 4.5 Database Operations: `ZwmDb`

In addition to the basic methods delegated by `ZwCADMech`, `ZwmDb` also provides:
- `switch_frame(frame_name: str)`: Switches the current operating drawing frame.
- `get_frame_count() -> int`: Gets the number of drawing frames in the drawing.
- `get_frame_name(index: int) -> str`: Gets the name of the drawing frame at the specified index.
- `get_next_frm_name() -> tuple(ZwmFrame, str)`: Gets the object and name of the next newly created drawing frame.
- `refresh_title()`, `refresh_bom()`, `refresh_frame()`: Refreshes the display of the corresponding objects in the drawing.
- `build_frame(switch_type: int = 511)`: Rebuilds the drawing frame block and associated blocks.

## 5. FAQ

**Q: Running error `Failed to load ZwmToolKit.tlb` or `ModuleNotFoundError`**
A: Ensure that ZWCAD Mechanical is installed on your system. If it is installed in a non-standard path, please copy the `ZwmToolKit.tlb` file to the directory where the script is running, or place it next to the `pyzwcadmech` folder. Also, ensure the `comtypes` library is installed.

**Q: Running error `Failed to create ZwmToolKit.ZwmApp. Please check if ZWCAD Mechanical is installed and running.`**
A: This is usually because:
1. The standard version of ZWCAD is installed on the computer, not the **Mechanical version** (ZWCAD Mechanical).
2. The COM component of ZWCAD Mechanical is not registered correctly. Try running ZWCAD Mechanical once as an administrator.

**Q: Getting title block or BOM returns null or throws an error**
A: Ensure that `mech.open_file("")` has been successfully called to bind the drawing database before calling `get_title()` or `get_bom()`. Also, ensure that the title block or BOM entities of ZWCAD Mechanical actually exist in the current drawing.
