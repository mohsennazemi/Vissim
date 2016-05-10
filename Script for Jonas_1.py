# COM-Server
import win32com.client as com
import os

# Connecting the COM Server
Vissim = com.Dispatch("Vissim.Vissim.800")

#filepath = 'C:\Program Files (x86)\PTV Vision\PTV Vissim 8\Exe\Vissim.exe'
#os.startfile(filepath)

# For this example, units are set to Metric.
# Note: PTV Vissim coordinates are always in meters [m]
UnitCurrent = Vissim.Net.NetPara.AttValue('UnitLenShort')
UnitAttributes = ('UnitAccel', 'UnitLenLong', 'UnitLenShort', 'UnitLenVeryShort', 'UnitSpeed', 'UnitSpeedSmall')
for UnitAttrCurr in UnitAttributes:
    Vissim.Net.NetPara.SetAttValue(UnitAttrCurr, 0)         # 0: Metric [m]

# Zoom Network Editor
Vissim.Graphics.CurrentNetworkWindow.ZoomTo(-400, -400, 400, 400)

#--------------------------------------
# Links
#--------------------------------------
# Input parameters of AddLink:
# 1: unsigned int Key   = attribute Number (No)         | example: 0 or 123 
# 2: BSTR WktLinestring = attribution Points3D          | example: 'LINESTRING(PosX1 PosY1 PosZ1, PosX2 PosY2 PosZ2, ..., PosXn PosYn PosZn)' with Pos as double : PosZ is optional
# 3: LaneWidths         = number and widths of lanes    | example: [WidthLane1, WiidthLane2, ... WidthLaneN] as double
LinksEnter = range(18)
LinksEnter[0] = Vissim.Net.Links.AddLink(0, 'LINESTRING(300 -23, 300 277)', [3.5, 3.5])
LinksEnter[1] = Vissim.Net.Links.AddLink(0, 'LINESTRING(300 -23, 300 -123)', [3.5, 3.5]) 
LinksEnter[2] = Vissim.Net.Links.AddLink(0, 'LINESTRING(300 -123, 300 -323)', [3.5, 3.5])
LinksEnter[3] = Vissim.Net.Links.AddLink(0, 'LINESTRING(100 -323, 100 -123)', [3.5, 3.5]) 
LinksEnter[4] = Vissim.Net.Links.AddLink(0, 'LINESTRING(-100 -23, -100 277)', [3.5, 3.5, 3.5, 3.5, 3.5, 3.5])
LinksEnter[5] = Vissim.Net.Links.AddLink(0, 'LINESTRING(-99	-123, -100 -23)', [3.5, 3.5, 3.5, 3.5, 3.5, 3.5])
LinksEnter[6] = Vissim.Net.Links.AddLink(0, 'LINESTRING(-100 -323, -99 -123)', [3.5, 3.5, 3.5, 3.5, 3.5, 3.5])
LinksEnter[7] = Vissim.Net.Links.AddLink(0, 'LINESTRING(-300 -23, -300 276)', [3.5, 3.5]) 
LinksEnter[8] = Vissim.Net.Links.AddLink(0, 'LINESTRING(-300 -23, -299 -322)', [3.5, 3.5])
LinksEnter[9] = Vissim.Net.Links.AddLink(0, 'LINESTRING(-100 277, 300 277)', [3.5, 3.5]) 
LinksEnter[10] = Vissim.Net.Links.AddLink(0, 'LINESTRING(-100 277, -300 276)', [3.5, 3.5])
LinksEnter[11] = Vissim.Net.Links.AddLink(0, 'LINESTRING(300 -23, -100 -23)', [3.5, 3.5, 3.5, 3.5])
LinksEnter[12] = Vissim.Net.Links.AddLink(0, 'LINESTRING(-100 -23, -300 -23)', [3.5, 3.5, 3.5, 3.5]) 
LinksEnter[13] = Vissim.Net.Links.AddLink(0, 'LINESTRING(300 -123, 100 -123)', [3.5, 3.5, 3.5, 3.5])
LinksEnter[14] = Vissim.Net.Links.AddLink(0, 'LINESTRING(100 -123, -99 -123)', [3.5, 3.5, 3.5, 3.5])
LinksEnter[15] = Vissim.Net.Links.AddLink(0, 'LINESTRING(100 -323, 300 -323)', [3.5, 3.5])
LinksEnter[16] = Vissim.Net.Links.AddLink(0, 'LINESTRING(-100 -323, 100 -323)', [3.5, 3.5]) 
LinksEnter[17] = Vissim.Net.Links.AddLink(0, 'LINESTRING(-100 -323, -299 -322)', [3.5, 3.5])

## ========================================================================
# Saving
#==========================================================================
Vissim.SaveNetAs('C:\Users\Sergio\Desktop\My Vissim Network\Pakour_1.inpx')
Vissim.SaveLayout('C:\Users\Sergio\Desktop\My Vissim Network\Parkour_1.layx')