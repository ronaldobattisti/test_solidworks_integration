
"""
functions to use with swApp
swApp.NewPart - Create a new part
swApp.CloseDoc('Peça1') - Closes the 'peça1' document
"""

import win32com.client

import win32com.client

def main():
    # Connect to SolidWorks
    swApp = win32com.client.Dispatch("SldWorks.Application")
    if not swApp:
        print("Failed to connect to SolidWorks.")
        return

    # Create a new part
    part = swApp.NewPart
    if not part:
        print("Failed to create a new part document.")
        return

    # Access the active model
    part = swApp.ActiveDoc

    # Create a sketch on the top plane
    sketch = part.SketchManager.AddSketch(part.GetPlane("Top Plane"))
    if not sketch:
        print("Failed to create a sketch.")
        return

    # Draw a rectangle in the sketch
    sketch.CreateLine(0, 0, 0, 0.1, 0)
    sketch.CreateLine(0.1, 0, 0, 0.1, 0.1)
    sketch.CreateLine(0.1, 0.1, 0, 0, 0.1)
    sketch.CreateLine(0, 0.1, 0, 0, 0)

    # Exit the sketch
    part.ShowNamedView2("", 7)  # Set to isometric view
    part.FeatureManager.FeatureExtrusion(True, False, False, 0, 0, 0.1, 0.01, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0)

    # Save the document
    part.SaveAs3("C:\\Temp\\AutomatedPart.SLDPRT", 0, 0)

    # Close SolidWorks
    swApp.ExitApp()

if __name__ == "__main__":
    main()




"""
import win32com.client

def open_solidworks():
    try:
        # Connect to SolidWorks application
        swApp = win32com.client.Dispatch("SldWorks.Application")
    except Exception as e:
        print("Error: Unable to connect to SolidWorks.")
        print(e)

        return None

    # Check if SolidWorks is already running
    if not swApp:
        print("Error: SolidWorks not available.")
        return None

    # Make SolidWorks visible
    swApp.Visible = True

    return swApp

if __name__ == "__main__":
    swApp = open_solidworks()
    swApp = swApp.NewPart
    print(swApp)
    swApp = swApp.ActiveDoc"""