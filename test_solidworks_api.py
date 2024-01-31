
import subprocess as sb
import win32com.client
import pythoncom

def is_solidworks_running():
    wmi = win32com.client.GetObject("winmgmts:")
    processes = wmi.ExecQuery("Select * from Win32_Process")

    for process in processes:
        if "SLDWORKS.exe" in process.Name:
            return True
    return False

def startSW():
    ## Starts Solidworks
    SW_PROCESS_NAME = r'C:/Program Files/SOLIDWORKS Corp/SOLIDWORKS/SLDWORKS.exe'
    sb.Popen(SW_PROCESS_NAME)

def shutSW():
    ## Kills Solidworks
    sb.call('Taskkill /IM SLDWORKS.exe /F')

def connectToSW():
    ## With Solidworks window open, connects to application      
    sw = win32com.client.Dispatch("SLDWORKS.Application")
    return sw

def openFile(sw, Path):
    ## With connection established (sw), opens part, assembly, or drawing file            
    f = sw.getopendocspec(Path)
    model = sw.opendoc7(f)
    return model

def updatePrt(model):
    ## Rebuilds the active part, assembly, or drawing (model)
    model.EditRebuild3

def getProperties(model):
    ## Allows you to see a list of custom properties associated with the active part, assembly, or drawing (model)
    
    modelExt = model.Extension
    p = modelExt.CustomPropertyManager("")
    names = p.GetNames
    #properties = p.GetObject
    return names#, properties
    
def change_color(model):
    swApp = win32com.client.Dispatch("SldWorks.Application")
    swApp.Visible = True
    part_doc = swApp.OpenDoc(model, 1)
    config = part_doc.GetActiveConfiguration()
    display_state = config.GetDisplayState()
    part_body = display_state.GetVisibleEntities2(1, -1)
    part_body.Color = (255, 0, 0)

if __name__ == '__main__':
    if is_solidworks_running():
        print("SolidWorks is already open.")
    else:
        print("SolidWorks is not open.")
        startSW()
    sw = connectToSW()
    model = openFile(sw, "C:\\Users\\RONAL\\Desktop\\solidworks integration\\test_1.SLDPRT")
    updatePrt(model)
    #names, properties = getProperties(model)
    names = getProperties(model)
    print(f'Names: {names}')
    #change_color(sw, model)
    change_color(model)
    #print(f'Properties: {properties}')
    

    