from pyautocad import Autocad, APoint
import win32com.client
import os

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

layouts = doc.Layouts
listLayouts = []
for layout in layouts:
      layoutName = layout.TabName
      doc.ActiveLayout = layoutName
      acad.app.ZoomExtents()
