import win32com.client
from PIL import ImageGrab

wb_file_name = 'C:/Memorias y servidor/Estabilizadores/Granshor_210527.xlsx'
outputPNGImage = 'C:/Memorias y servidor/Estabilizadores/test.jpg'

xls_file = win32com.client.gencache.EnsureDispatch("Excel.Application")

wb = xls_file.Workbooks.Open(Filename=wb_file_name)
xls_file.DisplayAlerts = False 
ws = wb.Worksheets("Estabilizador")
ws.Range(ws.Cells(46,1),ws.Cells(110,14)).CopyPicture(Format= win32com.client.constants.xlBitmap)  # example from cell (1,1) to cell (15,3)
img = ImageGrab.grabclipboard()
img.save(outputPNGImage)
wb.Close(SaveChanges=False, Filename=wb_file_name)