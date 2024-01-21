import win32com.client as win32

swApp = win32.Dispatch("SldWorks.Application")
swDoc = swApp.ActiveDoc
#swDoc = swApp.OpenDoc("C:\\Users\\RONAL\\Desktop\\solidworks integration\\test_1.SLDPRT", 1)

print(swDoc.getPathName)
#print(swDoc.getProperties())
