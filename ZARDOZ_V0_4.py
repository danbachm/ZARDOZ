##############################################################################
# ZARDOZ
# Date/Version: 28.01.2016  V0.4
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
#                       |ZARDOZ_V0_4.py
#
# To create a Rhino Command Line Alias:
# -Open Tools|Options|Aliases from the Rhino Menu
# -Press "New" to create a new Alias
# -Enter a name for the Alias. We prefer to use ZARDOZ as command-name ;)
# -Enter the following as Command Macro:  ! _-RunPythonScript "FilePath\ScriptFileName"
# -It should look like this: ! _-RunPythonScript "C:\Users\User\AppData\Roaming\McNeel\Rhinoceros\5.0\scripts\ZARDOZ_V0_4.py"
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
# 28.1.2016 Added variables for tolerance options in ConvertCurveToPolyline
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
minAngleTolerance = .01         #Min Depth in mm #NOT USED 
maxAngleTolerance = 100        #Max Depth in mm #NOT USED
defaultAngleTolerance = 5     #Default Depth in mm #NOT USED
minDistTolerance = 0.0001         #Min Depth in mm #NOT USED 
maxDistTolerance = 40        #Max Depth in mm #NOT USED
defaultDistTolerance = 0.01     #Default Depth in mm #NOT USED

#default variables for HPGL
VS = "VS 10, 40;"       #Default Settings for PenUp/PenDown Velocity #NOT USED
ZP = "ZP 1000, 200;"    #Default Settings for Z Axis Position ZP 1000, 200 #NOT USED

################################################################################
#ZARDOZ handles the Commandline Input
def zardoz():
    # make temporary layers
    rs.AddLayer("tmp_curve")
    rs.AddLayer("tmp_polyline")
    
    # Curve Selection
    curve1 = rs.ObjectsByLayer("ckdr", True);
    objectList = rs.GetObjects("Select the curve to cut", rs.filter.curve)
    
    # check if more than one curve was selected
    if len(objectList) > 1:
        print "More than one curve found in selection. Taking the first one..."
    curve = objectList[0]
        
    # move selected curves to tmp layer and set invisible
    restoreLayer = rs.ObjectLayer(curve)
    rs.ObjectLayer(objectList, "tmp_curve")
    
    # change the curve direction if needed
    if rs.CurveDirectionsMatch(curve1, curve):
        print "Curve has the right direction (clockwise)"
    else:
        print "Curve has the wrong direction (not clockwise)"
        rs.ReverseCurve(objectList);   
        print "Curve direction has been changed and is clockwise now"
    
    # Option Configuration
    go = Rhino.Input.Custom.GetOption()
    go.SetCommandPrompt("Select cutting options. Press Enter when Done:")

    #set options for CommandPrompt
    DistToleranceOpt   = Rhino.Input.Custom.OptionDouble(defaultDistTolerance, minDistTolerance, maxDistTolerance)
    AngleToleranceOpt   = Rhino.Input.Custom.OptionDouble(defaultAngleTolerance, minAngleTolerance, maxAngleTolerance)
    SpeedOpt   = Rhino.Input.Custom.OptionInteger(defaultSpeed, minSpeed, maxSpeed)
    go.AddOptionDouble("DistTolerance", DistToleranceOpt)
    go.AddOptionDouble("AngleTolerance", AngleToleranceOpt)
    go.AddOptionInteger("Speed", SpeedOpt)

    polyline = []
            
    #Check Boundary for curves not fitting into workspace
    poschecker = checkBound(curve)
    if poschecker:
        #Loop to set CommandPrompt variables 
        while True:
            # show res
            #print "Current Res: %s" % AngleToleranceOpt.CurrentValue
            if curve:
                if polyline:
                    #print "deleting old polyline"
                    rs.DeleteObject(polyline)
                #print "converting curve id %s to polyline with res %s" % (curve, AngleToleranceOpt.CurrentValue)
                polyline = rs.ConvertCurveToPolyline(curve,AngleToleranceOpt.CurrentValue,DistToleranceOpt.CurrentValue,False)
                #print "converted curve to polyline id %s" % polyline
                rs.ObjectLayer(polyline, "tmp_polyline")
                rs.SelectObjects(polyline)
                rs.Command("_PointsOn")
                rs.UnselectObject(polyline)
            else:
                print("no curve selected")
            
            
            get_rc = go.Get()
            if get_rc == Rhino.Input.GetResult.Option:
                continue
            break

        # Send the Curve to Machine or abort
        checkUser = rs.GetString("Do you want to write Gcode for this Job?","Y","Y / N")
        if checkUser == "Y" and polyline:
            #print "Sending polyline id %s to machine" % polyline
            sendEverything (polyline, SpeedOpt.CurrentValue)
            print("ZARDOZ has successfully created your code")
        else:
            print("Almighty ZARDOZ forgives you!")
        CleanUp (restoreLayer, curve, polyline)
        return Rhino.Commands.Result.Success
    else:
        CleanUp (restoreLayer, curve)
        return


def CleanUp (restoreLayer, curve, polyline):
    #print "cleaning up"
    rs.ObjectLayer(curve, restoreLayer)
    rs.DeleteLayer("tmp_curve")
    rs.DeleteObject(polyline)
    rs.DeleteLayer("tmp_polyline")
    
################################################################################
#Send the job, handles Zuend-HPGL Conversion and streams the data to the machine
def sendEverything (polyline, speed):

    #-----------Export to File-START----------------------------------------
    #        f = open("C:\PLT\dumpZardoz.plt","w")
    now = time.strftime("%Y%M%D_%H%M%S")
    f = open("C:\Users\YOURUSERNAME\Desktop\PLT\cut_"+str(now)+".plt","w")
    
    header = ""
    footer = ""
    commands = []
    tmpcommands = []
    tmpcommands.append("\n"+"G01 F"+str(speed))
    curve = rs.CurvePoints(polyline)
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