# LibreOffice + Python UNO Cheat Sheet

This cheat sheet covers essential LibreOffice automation tasks using Python and UNO (Universal Network Objects). Useful for working with `.ods` and `.xlsx` files on Linux.

---

## 🚀 Setup

### ✅ Start LibreOffice in Headless Mode

```bash
libreoffice --headless --accept="socket,host=localhost,port=2002;urp;" --nologo --nofirststartwizard &
```

### ✅ Python Dependencies

Install `python3-uno`:

```bash
sudo apt install python3-uno
```

---

## 🔌 Connect to LibreOffice from Python

```python
import uno

local_ctx = uno.getComponentContext()
resolver = local_ctx.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", local_ctx)
ctx = resolver.resolve(
    "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
desktop = ctx.ServiceManager.createInstanceWithContext(
    "com.sun.star.frame.Desktop", ctx)
```

---

## 📁 File Handling

### ✅ Open an Excel File

```python
def open_doc(path):
    file_url = uno.systemPathToFileUrl(path)
    return desktop.loadComponentFromURL(file_url, "_blank", 0, ())

doc = open_doc("/path/to/Book1.xlsx")
sheet = doc.Sheets.getByIndex(0)
```

### ✅ Save As New File

```python
from com.sun.star.beans import PropertyValue
save_props = (PropertyValue(Name="FilterName", Value="Calc MS Excel 2007 XML"),)
doc.storeToURL(uno.systemPathToFileUrl("/path/to/output.xlsx"), save_props)
```

### ✅ Close Document

```python
doc.close(True)
```

---

## 📄 Sheet & Cell Operations

### ✅ Access Sheet & Cell

```python
sheet = doc.Sheets.getByIndex(0)
cell = sheet.getCellByPosition(0, 0)  # A1
```

### ✅ Set Cell Values

#### 🔹 String/Text

```python
cell.String = "Hello World"
```

#### 🔹 Numeric Value

```python
cell.Value = 123.45
```

#### 🔹 Formula

```python
cell.Formula = "=SUM(B1:B5)"
```

#### 🔹 Boolean

```python
cell.Value = 1  # True
```

#### 🔹 Set Multiple Cells in a Loop

```python
for i in range(5):
    cell = sheet.getCellByPosition(0, i)  # A1 to A5
    cell.String = f"Row {i+1}"
```

---

## 🖌️ Formatting

### ✅ Format Part of Text in a Cell

```python
text = cell.Text
cursor = text.createTextCursor()
cursor.gotoEnd(False)
cursor.goLeft(4, True)  # Select last 4 characters (e.g., "UNO!")
cursor.CharColor = 0xFF0000  # Red

cursor.gotoStart(False)
cursor.goRight(len("Updated via Python "), True)
cursor.CharColor = 0x000000  # Black
```

### ✅ Set Font Weight and Size

```python
cursor.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
cursor.CharHeight = 14
```

---

## 🧩 Merge & Unmerge Cells

```python
sheet.getCellRangeByName("A1:A5").merge(True)   # Merge
sheet.getCellRangeByName("A1:A5").merge(False)  # Unmerge
```

---

## 📏 Borders

```python
from com.sun.star.table import BorderLine

line = BorderLine()
line.OuterLineWidth = 50
cell.TableBorder.TopLine = line
cell.TableBorder.BottomLine = line
```

---

## 📐 Diagonal Line (via Borders)

```python
from com.sun.star.table import BorderLine2
from com.sun.star.table.BorderLineStyle import SOLID

line = BorderLine2()
line.LineWidth = 50
line.LineStyle = SOLID

cell.DiagonalBLTR = line  # Bottom-left to top-right
cell.DiagonalTLBR = line  # Top-left to bottom-right
```

---

## 🧪 Check LibreOffice Status

```bash
ps aux | grep libreoffice
ss -ltnp | grep 2002
nc -zv localhost 2002
```

---

## 🧠 Tips

- UNO cell indexes are **0-based**: A1 → (0, 0), B2 → (1, 1), etc.
- Use `.String` for text, `.Value` for numbers/dates, `.Formula` for formulas.
- Use `uno.getConstantByName` for constants like font styles, weights, etc.

---

Need advanced features like inserting images, creating charts, or macro automation? Let me know!

---

Running container within ubuntu wsl for rhel 7
```
    1  sudo apt update
    2  sudo apt install podman
    3  podman login registry.redhat.io
    4  podman run -it registry.redhat.io/rhel7/rhel bash
```
