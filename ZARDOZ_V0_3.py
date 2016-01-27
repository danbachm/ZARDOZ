##############################################################################
# ZARDOZ
# Date/Version: 20.11.2015  V0.3
#
# Written by    - Daniel Bachmann       (daniel.bachmann@arch.ethz.ch)
#
# Streamlined Foam Machine Cutting out of Rhino V5
# ZARDOZ converts Curves and translates them into G-Code, then 
# streams the Data directly to the CUT 1610S-3D Foam-Cutter.
# The Script is meant to be started from Rhino Command-Line but can
# also be assigned to an Icon.
#
# Prerequisites:
# The Script is depending on pyserial to send commands to the CUT 1610S-3D Foam-Cutting-Machine.
# Make sure you have copied the serial folder of pyserial to the same directory where the ZARDOZ script resides.
# ->   Rhinoscriptfolder|serial
#                       |ZARDOZ_V0_1.py
#
# To create a Rhino Command Line Alias:
# -Open Tools|Options|Aliases from the Rhino Menu
# -Press "New" to create a new Alias
# -Enter a name for the Alias. We prefer to use ZARDOZ as command-name ;)
# -Enter the following as Command Macro:  ! _-RunPythonScript "FilePath\ScriptFileName"
# -It should look like this: ! _-RunPythonScript "C:\Users\User\AppData\Roaming\McNeel\Rhinoceros\5.0\scripts\ZARDOZ_V0_1.py"
#
# Using the variable definitions it is possible to customize the default Workspace Area and Cutting Speeds/Depths
# Use the sendToMachine Variable-Option to save to a local file instead of sending the Data directly to the Foam-Cutter
#
# The script does some basic error-handling:
# -Only Curves will be selectable
# -A Bounding Box is used to check if all selected Curves are in the Workspace Area
#
# Changes/Bugfixes
# ----------------
# 3.12.2015 Cleaned code and comments
# 9.12.2015 Added Timestamp to Filename
# 9.12.2015 Saves Files to Desktop/PLT
#
#
# ToDo
# ----
# -Custom Filename option
# -2D/3D/Rotation Cut option
# -generate Mesh and Solid Outlines
# -generate automated in and out points for cutting
# -Automated Outlinegenerator for rotated objects
# -G-Code Accumulator for rotated objects
#
###############################################################################


import rhinoscriptsyntax as rs
import scriptcontext as sc
import Rhino
import serial
import time

Rhino.Runtime.HostUtils.DisplayOleAlerts(False)

#Switches from Machine to File Output
sendToMachine=False      #False for File Output

#default variables for Workspace settings
maxWSPx=1300           #Max Workspace Coordinate X-Axis in mm
maxWSPy=700            #Max Workspace Coordinate Y-Axis in mm
minWSPx=-1               #Min Workspace Coordinate X-Axis in mm
minWSPy=-1               #Min Workspace Coordinate Y-Axis in mm

#default variables for Commandline Options
minSpeed = 1            #Min Speed in mm/s
maxSpeed = 900           #Max Speed in mm/s
defaultSpeed = 550       #Default Speed in mm/s Ideal for EPS. Go slower for XPS
minDepth = 0.01         #Min Depth in mm #NOT USED 
maxDepth = 40.00        #Max Depth in mm #NOT USED
defaultDepth = 0.10     #Default Depth in mm #NOT USED

#default variables for HPGL
VS = "VS 10, 40;"       #Default Settings for PenUp/PenDown Velocity #NOT USED
ZP = "ZP 1000, 200;"    #Default Settings for Z Axis Position ZP 1000, 200 #NOT USED

################################################################################
#ZARDOZ handles the Commandline Input
def zardoz():
    # Curve Selection
    curve1 = rs.ObjectsByLayer("ckdr", True);
    objectList = rs.GetObjects("Select the curve to cut", rs.filter.curve)
    if rs.CurveDirectionsMatch(curve1, objectList):
        print "Curves are in the same direction"
    else:
        print "Curve are not in the same direction"
        rs.ReverseCurve(objectList);   
        print "Curve direction has been changed and is clockwise now"
    



    # Option Configuration
    go = Rhino.Input.Custom.GetOption()
    go.SetCommandPrompt("Select cutting options. Press Enter when Done:")
    rs.SelectObjects(objectList)

    #set options for CommandPrompt
    listValues = "Tool1", "Tool2", "Pen"
    DepthOpt   = Rhino.Input.Custom.OptionDouble(defaultDepth, minDepth, maxDepth)
    SpeedOpt   = Rhino.Input.Custom.OptionInteger(defaultSpeed, minSpeed, maxSpeed)
    listIndex = 0
    opList = go.AddOptionList("Tool", listValues, listIndex)
    go.AddOptionDouble("Depth", DepthOpt)
    go.AddOptionInteger("Speed", SpeedOpt)

    #Check Boundary for curves not fitting into workspace
    poschecker = checkBound(objectList)
    if poschecker:
        #Loop to set CommandPrompt variables 
        while True:
            get_rc = go.Get()
            if get_rc == Rhino.Input.GetResult.Option:
                #Refresh Option Variables
                if go.OptionIndex()==opList:
                    listIndex = go.Option().CurrentListOptionIndex
                continue
            break



        # Send the Curves to Machine or abort
        polyline = []
        if objectList:
            for id in objectList:
                polyline.append(rs.ConvertCurveToPolyline(id,5.0,0.01,True))
        else:
            print("no lines selected")
        checkUser = rs.GetString("Do you want to send this Job to the machine?","y","y / n")
        if checkUser == "y":
            sendEverything (polyline, listIndex, DepthOpt.CurrentValue, SpeedOpt.CurrentValue)
            print("ZARDOZ has finished your job successfully")
        else:
            print("Almighty ZARDOZ forgives you!")
        return Rhino.Commands.Result.Success
    else:
        return



################################################################################
#Send the job, handles Zuend-HPGL Conversion and streams the data to the machine
def sendEverything (lineObjects, tool, depth, speed):
    #compensate for pen -> equals tool 4
    tool = tool +1
    if tool == 3: tool = tool+1
    
    if sendToMachine:
        #-----------Export to Machine-START-----------------------------------------
        header = "PS 1,1;PB 2,1;\nDT 59;\nUR ZuendTest;\nSP "+str(tool)+"; "+"TR 1;PU;PA;\nSP "+str(tool)+";\n"+"ZP 1000, "+str(int(depth))+";\nSP "+str(tool)+";\nVS "+str(speed)+", 40;"
        footer = "PU;\nPU;\nPU;PA 160000,0;BP;PS 1,0;PB 2,0;NR;"
        commands = []
        tmpcommands = []
        for pline in lineObjects:
            curve = rs.CurvePoints(pline)
            i=0
            for p in curve:
                if i < 1:
                    tmpcommands.append("\n"+"PU "+str(int(p[0]*100))+", "+str(int(p[1]*100))+";")
                else:
                    tmpcommands.append("\n"+"PD "+str(int(p[0]*100))+", "+str(int(p[1]*100))+";")
                i = i+1
        commands.append(''.join(tmpcommands))
        serialSend(header+''.join(commands)+footer)
        #-----------Export to Machine-END---------------------------------------
    else:
        #-----------Export to File-START----------------------------------------
#        f = open("C:\PLT\dumpZardoz.plt","w")
        now = time.strftime("%Y%M%D_%H%M%S")
        f = open("C:\Users\danbachm\Desktop\PLT\cut_"+str(now)+".plt","w")
        





        header = ""
        footer = ""
        commands = []
        tmpcommands = []
        tmpcommands.append("\n"+"G01 F"+str(speed))
        for pline in lineObjects:
            curve = rs.CurvePoints(pline)
            i=0
            for p in curve:
                if i < 1:
                    tmpcommands.append("\n"+"G01 X"+str(int(p[1]))+" Y"+str(int(p[2]))+" A"+str(int(p[1]))+" B"+str(int(p[2])))
                else:
                    tmpcommands.append("\n"+"G01 X"+str(int(p[1]))+" Y"+str(int(p[2]))+" A"+str(int(p[1]))+" B"+str(int(p[2])))
                i = i+1
        commands.append(''.join(tmpcommands))
        f.write(header)
        f.writelines(''.join(commands))
        f.write("\n"+footer)
        f.close()
        print("job sent to file...")
        #-----------Export to File-END------------------------------------------


################################################################################
# Checks if Bounding Box fits in Workspace
def checkBound(selectedObj):
    box = rs.BoundingBox(selectedObj)
    if box:
        if ((min(box)[0]) < minWSPx) or ((min(box)[2]) < minWSPy) or ((max(box)[0]) > maxWSPx) or ((max(box)[2]) > maxWSPy):
            print ("ZARDOZ says: Curves are not in Workspace-Area. Make sure your Curves are inside the Workspace!")
            return False
        else:
            print ("ZARDOZ says: All Curves are happily located inside the Workspace")
            return True

if __name__ == "__main__":
    zardoz()