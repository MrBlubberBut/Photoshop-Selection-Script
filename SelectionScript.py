import pygetwindow
from photoshop import Session
from pywintypes import com_error
from win32com.client import Dispatch, GetActiveObject
import win32com.client
import keyboard
from pynput.mouse import Listener
from pynput import mouse


def main():
    def start():
        # Defining necessary variables
        app = GetActiveObject("Photoshop.Application")
        psApp = win32com.client.Dispatch("Photoshop.Application")
        doc = psApp.Application.ActiveDocument
        docRef = app.ActiveDocument


        # Logs the current layer
        previousLayer = docRef.activeLayer.name


        # Logs the current tool
        with Session() as ps:
            tool = ps.app.currentTool


        # Goes to the top most layer
        docRef.ActiveLayer = docRef.Layers.Item(1)


        # Plays the action "Magic Wand for Selection Script" in group "Selection Script Group" in the Actions panel
        # Sets the tool to Magic Wand
        app.DoAction('Magic Wand for Selection Script', 'Selection Script Group')


        # Waits for the left click
        # If it a mouse click that is not the left mouse button, does nothing
        def on_click(x, y, button, pressed):
            while button == mouse.Button.left:
                return False

        with Listener(on_click=on_click) as listener:
            listener.join()


        # After left click, sets the tool as the the tool that was selected previously
        ps.app.currentTool = tool


        # Goes to the layer that was previously active
        docRef.ActiveLayer = doc.ArtLayers[previousLayer]


    # Checks whether or not Photoshop is minimzed or not every time the F key is pressed
    # If Photoshop is minimized, the script does nothing
    # If Photoshop is active, proceeds with the script
    while True:
        def check():
            try:
                psApp = win32com.client.Dispatch("Photoshop.Application")
                doc = psApp.Application.ActiveDocument
                photoshop = pygetwindow.getWindowsWithTitle(doc.Name)[0]
                if photoshop.isActive:
                    start()
                    pass
                else:
                    pass
            except com_error:
                pass
        keyboard.wait('f')
        check()


main()







