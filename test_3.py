import win32com.client

try:
    swApp = win32com.client.Dispatch("SldWorks.Application")
Exception
swApp.Visible = True
swApp.NewPart()
swApp.Close("peça3")