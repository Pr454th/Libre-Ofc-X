import uno
from com.sun.star.beans import PropertyValue

# Connect to LibreOffice
local_context = uno.getComponentContext()
resolver = local_context.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", local_context)
    
context = resolver.resolve(
    "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    
desktop = context.ServiceManager.createInstanceWithContext(
    "com.sun.star.frame.Desktop", context)

# Load the spreadsheet
def open_doc(path):
    file_url = uno.systemPathToFileUrl(path)
    props = tuple()
    return desktop.loadComponentFromURL(file_url, "_blank", 0, props)

doc = open_doc("/home/kai/Dev/Libre-test/Report.xlsx")
sheet = doc.getSheets().getByIndex(0)
cell = sheet.getCellByPosition(0, 0)  # A1

# Set entire string
cell.String = "Updated via Python UNO!"

# Apply formatting to "UNO"
text = cell.Text
cursor = text.createTextCursor()
cursor.gotoEnd(False)
cursor.goLeft(4, True)  # Select "UNO!"

# Apply red color to selected text
cursor.CharColor = 0xFF0000  # Red

# Move cursor to rest of the string and reset to black
cursor.gotoStart(False)
cursor.goRight(len("Updated via Python "), True)
cursor.CharColor = 0x000000  # Black

# Save to new file
save_props = (PropertyValue(Name="FilterName", Value="Calc MS Excel 2007 XML"),)
doc.storeToURL(uno.systemPathToFileUrl("/home/kai/Dev/Libre-test/Book1_colored.xlsx"), save_props)
doc.close(True)

print("Cell A1 updated with partial red-colored text and saved.")