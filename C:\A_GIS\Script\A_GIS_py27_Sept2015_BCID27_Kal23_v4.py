# Created by Shea Hammond, USFWS A_GIS BCID27 and Kaleidoscope23, April 2015

# There are only three files this script utilizes. First is the gps.txt file created when exporting data from the anabat, second is the data output from BCID27, and lastly the output from Kal23.
# This script simply georefrences and maps the Kal output data. NO classification of bat data is performed by this script.
# The script will create a new folder within the folder containing the data, copy the files to the newly created folder, then modify the copied files.


from __future__ import with_statement
import csv
import os
from os import path
import xlwt
import sys
import arcpy
import math
import string
import win32com.client
import shutil
import glob
from Tkinter import Tk
from tkFileDialog import askopenfilename, askdirectory
import xlrd
from os import sys


# #### Select folder and file locations ####

# Tk() opens GUI to allow user to define path to file folders containing data
Tk().withdraw() # allows GIU to open for file selection when command is given
print ('Select file Folder containing BOTH gps.txt and Kal output file')
FileinputMain = askdirectory() # command to open GIU and select root directory folder contating gis.txt file
FileInputMain = os.path.normpath(FileinputMain)
print ('You selected ' + FileInputMain)
print ('')


path = FileInputMain
GPS_file = glob.glob(path + '/*gps.txt') # Identifies gps.txt file in directory selected
Kal_file = glob.glob(path + '/id.csv') # Identifies Kal file in directory selected
BCID_file = glob.glob(path + '/*_*.xls') # Identifies BCID file in directory selected

Fix_GPS = ''.join(GPS_file) # converts list to string
Fix_Kal = ''.join(Kal_file)
Fix_BCID = ''.join(BCID_file)

print "GPS FILE SELECTED - - -> "+ (Fix_GPS)
print ""
print "Kal OUTPUT FILE SELECTED - - -> " + (Fix_Kal)
print ""
print "BCID OUTPUT FILE SELECTED - - -> " + (Fix_BCID)
print ""

# normalize path
Orig_GPS_file = os.path.normpath(Fix_GPS)
Orig_Kal_file = os.path.normpath(Fix_Kal)
Orig_BCID_file = os.path.normpath(Fix_BCID)


# #### Create new folder and copy files to folder

try:
    Shapefiles = (FileInputMain + '/Shapefiles')
    os.makedirs(Shapefiles) # create new folder

    shutil.copyfile(Orig_GPS_file, Shapefiles +'/GPScopy.txt')
    FileInput = os.path.normpath(Shapefiles) # copy gps.txt file

    shutil.copyfile(Orig_Kal_file, Shapefiles + '/Kalcopy.csv')
    FileKalinput = os.path.normpath(Shapefiles +'/Kalcopy.csv') # copy and rename Kal file

    shutil.copyfile(Orig_BCID_file, Shapefiles + '/BCIDcopy.xls')
    FileBCIDinput = os.path.normpath(Shapefiles +'/BCIDcopy.xls') # copy and rename BCID file


except:
     print arcpy.GetMessages()
     print "Didn't Make New Folder"
     sys.exit


# #### General File Prep ####

print ('Prep files for Excel')

# preps gps txt file for import to cvs - This command replaces blank spaces with a commas within the gps.txt file and saves file as .cvs, thus allowing Excel to open file in a comma delemeted format
with open(FileInput + '/mod0gps.txt', 'w') as outfile:
    with open(FileInput + '/GPScopy.txt') as infile:
        outfile.write(infile.read().replace(" ", ", ")) # works

# Found that as altitude changes so does the number of blank spaces within the altitude column. This series of commands will standardize the altitude column.
# Fix data with gpx.txt file, random errors in output file
with open(FileInput + '/mod1gps.txt', 'w') as outfile:
    with open(FileInput + '/mod0gps.txt') as infile:
        outfile.write(infile.read().replace(", , , , , , ,", ", , , , ")) # works

# Fix data with gpx.txt file, random errors in output file
with open(FileInput + '/mod2gps.txt', 'w') as outfile:
    with open(FileInput + '/mod1gps.txt') as infile:
        outfile.write(infile.read().replace(", , , , , ,", ", , , , ")) # works

# Fix data with gpx.txt file, random errors in output file
with open(FileInput + '/mod3gps.txt', 'w') as outfile:
    with open(FileInput + '/mod2gps.txt') as infile:
        outfile.write(infile.read().replace(", , , , ,", ", , , , ")) # works

# renames gps.txt file to .csv file that is comma delimted
os.rename(FileInput + '/mod3gps.txt',FileInput + '/mod3gps.csv')

# Converts Allen .xls file to .cvs
changeWb = xlrd.open_workbook(FileBCIDinput)
sh = changeWb.sheet_by_name('File Level')
csv_file = open(FileInput + '/AllenBCID.csv', 'wb')
wr = csv.writer(csv_file, quoting=csv.QUOTE_ALL)
for rownum in xrange(sh.nrows):
      wr.writerow(sh.row_values(rownum))
csv_file.close()

print ('Prep Files for GIS')

# ##### Starts Excel #####
xl = win32com.client.Dispatch("Excel.Application") #works
xl.Visible = True #works

# Allows for Excel files to overwite - - for testing purposes - -
xl.DisplayAlerts = False

# Opens csv file in excel - Modify route
wb = xl.WorkBooks.open(FileInput + '/mod3gps.csv') #works

# imports and runs excel macro
xl.VBE.ActiveVBProject.VBComponents.Import('c:\A_GIS\ScriptFiles\DeleteTopRows.bas') # Import and Run Macro
xl.Run('DeleteTopRows') #works

# imports and runs excel macro
xl.VBE.ActiveVBProject.VBComponents.Import('c:\A_GIS\ScriptFiles\AdjustTopRows.bas') # Import and Run Macro
xl.Run('AdjustTopRows') #works

# imports and runs excel macro
xl.VBE.ActiveVBProject.VBComponents.Import('c:\A_GIS\ScriptFiles\ModifyLatLong.bas') # Import and Run Macro
xl.Run('ModifyLatLong') #works

# imports and runs excel macro
xl.VBE.ActiveVBProject.VBComponents.Import('c:\A_GIS\ScriptFiles\Call_ID.bas') # Import and Run Macro
xl.Run('Call_ID') # works

# imports and runs excel macro
xl.VBE.ActiveVBProject.VBComponents.Import('c:\A_GIS\ScriptFiles\Alt_m.bas') # Import and Run Macro
xl.Run('Alt_m') # works

# imports and runs excel macro
xl.VBE.ActiveVBProject.VBComponents.Import('c:\A_GIS\ScriptFiles\CleanRouteGIS1.bas') # Import and Run Macro
xl.Run('CleanRouteGIS') # works

# imports and runs excel macro
xl.VBE.ActiveVBProject.VBComponents.Import('c:\A_GIS\ScriptFiles\FixDate2.bas') # Import and Run Macro
xl.Run('FixDate2') # works

# imports and runs excel macro
xl.VBE.ActiveVBProject.VBComponents.Import('c:\A_GIS\ScriptFiles\DT.bas') # Import and Run Macro
xl.Run('DT') # works

# imports and runs excel macro
#xl.VBE.ActiveVBProject.VBComponents.Import('c:\A_GIS\ScriptFiles\FixTime2.bas') # Import and Run Macro
#xl.Run('FixTime2') # works

# saves as output csv file
wb.SaveAs(FileInput + '/GPS_GISinput.csv')
wb.Close(SaveChanges=0) #works

print ('Route File Prep Complete')

# closes excel
xl.Visible = True
xl.Quit

# Opens csv file in excel - Modify Kal
wb = xl.WorkBooks.open(FileKalinput) #works

# imports and runs excel macro
xl.VBE.ActiveVBProject.VBComponents.Import('c:\A_GIS\ScriptFiles\KalMod.bas') # Import and Run Macro
xl.Run('KalMod') #works

# saves as output csv file
wb.SaveAs(FileInput + '/Kal_GISinput.csv')
wb.Close(SaveChanges=0) #works

print ('Kal File Prep Complete')

# closes excel
xl.Visible = False
xl.Quit

# Opens csv file in excel - Modify BCID
wb = xl.WorkBooks.open(FileInput + '/AllenBCID.csv') #works

# imports and runs excel macro
xl.VBE.ActiveVBProject.VBComponents.Import('c:\A_GIS\ScriptFiles\DTR_BCID2_7c.bas') # Import and Run Macro
xl.Run('DTR_BCID2_7c') #works

# imports and runs excel macro
xl.VBE.ActiveVBProject.VBComponents.Import('c:\A_GIS\ScriptFiles\AllenTable3.bas') # Import and Run Macro
xl.Run('AllenTable') #works

# imports and runs excel macro
xl.VBE.ActiveVBProject.VBComponents.Import('c:\A_GIS\ScriptFiles\AllenRows.bas') # Import and Run Macro
xl.Run('AllenRows') #works

# saves as output csv file
wb.SaveAs(FileInput + '/BCID_GISinput.csv')
wb.Close(SaveChanges=0) #works

print ('Allen File Prep Complete')

# closes excel
xl.Visible = False
xl.Quit



# #### Start GIS ####
# Create workspace
arcpy.env.workspace = "c:\A_GIS\ScriptFiles\Workstation"

# Set Overwrite option
arcpy.env.overwriteOutput = True

# unQulaify Field Names
arcpy.env.qualifiedFieldNames = False

# Add Toolbox
arcpy.AddToolbox(r"C:\A_GIS\ScriptFiles\AAToolbox.tbx")

try: # Add XY to Layer
    in_Table = (FileInput + '/GPS_GISinput.csv')
    in_x = "LAT"
    in_y = "LONG"
    out_Layer = "GPSinput"
    save_Layer = (FileInput + '/GPSinput.lyr')
    coordsys = r"C:\A_GIS\Coordinate System\WGS 1984.prj" # Specific to computer. Individual users will need to modify.

    arcpy.MakeXYEventLayer_management(in_Table, in_y, in_x, out_Layer, coordsys)
    print "Added XY Data"

    arcpy.SaveToLayerFile_management(out_Layer, save_Layer, "RELATIVE")
    print "Save XY as Layer"

except:
    print arcpy.GetMessages()
    print "Didn't Add XY Data"
    sys.exit

try: # Add FID to XY Layer
    Save_route = (FileInput + '/GPSinput.shp')
    Route_Saved = (FileInput + '/SavedRoute.shp')

    arcpy.Select_analysis(save_Layer, Save_route)
    print "Add FID to XY Layer"

    arcpy.CopyFeatures_management(Save_route, Route_Saved)
    print "Saved XY with FID as Shapefile"

except:
    print arcpy.GetMessages()
    print "Didn't add FID"
    sys.exit

try:
    arcpy.MakeFeatureLayer_management(Route_Saved, 'SavedRoute')
    arcpy.MakeFeatureLayer_management(Route_Saved, 'SavedRouteKal')
    arcpy.MakeFeatureLayer_management(Route_Saved, 'SavedRouteAllen')
    print 'Convert XY Shapefile to Layer'

    KalTable = (FileInput + '/Kal_GISinput.csv')
    arcpy.MakeTableView_management(KalTable, 'KalTable')
    print 'Make Kal Tableview'

    AllenTable = (FileInput + '/BCID_GISinput.csv')
    arcpy.MakeTableView_management(AllenTable, 'AllenTable')
    print 'Make Allen Tableview'

except:
    print arcpy.GetMessages()
    print 'Didnt make Tableview'
    sys.exit()

# ####  Kal Calls ####
try: # Join Call File to XY Layer
    RouteField = "Call_ID"
    KalField = "K_ID"

    arcpy.AddJoin_management('SavedRouteKal', RouteField, 'KalTable', KalField, "KEEP_ALL")
    print "Join Kal to XY Layer"

except:
    print arcpy.GetMessages()
    print "Didn't add to XY"
    sys.exit()

try: # Save selected joins
    JoinSavedKal1 = (FileInput + '/JoinSavedKal1.shp')

    arcpy.Select_analysis('SavedRouteKal', JoinSavedKal1)
    print "Saved Kal Join"
except:
    print arcpy.GetMessages()
    print "Didn't Save Kal Join"
    sys.exit

try: # Save selected joins
    JoinSavedKal2 = (FileInput + '/AcousticRouteK.shp')
    KalRouteData = (FileInput + '/RouteDataK.lyr')

    arcpy.MakeFeatureLayer_management('SavedRouteKal', KalRouteData)
    print 'Made Acoustic Route a Feature Layer'

    arcpy.Select_analysis('SavedRouteKal', JoinSavedKal2)
    print ('Saved shapefile containing Acoustic Route with Kal information joined as: ')
    print (':---> ' + JoinSavedKal2)

except:
    print arcpy.GetMessages()
    print "Didn't Save Acoustic Route"
    sys.exit

try: # Make Route a Line
    Route_Line = (FileInput + '/RouteLine.shp')

    arcpy.PointsToLine_management(Route_Saved, Route_Line,"", "Call_ID")
    print ("Created Line Shapefile of Route")
    print (":--->" + Route_Line)
except:
    print arcpy.GetMessages()
    print "Didn't Make Route Line Shapefile"
    sys.exit

try: # Select Call Locations
     KalCalls = (FileInput + '/KalCalls.shp')

     arcpy.Select_analysis(JoinSavedKal2, KalCalls, '"KalNum" > 0')

     print ("Selected Calls")
     print ("Kal Calls Shapefile:-->" + KalCalls)

except:
     print arcpy.GetMessages()
     print "Didn't Make Selected Call Shapefiles"
     sys.exit


# #### BCID Calls ####

try: # Join Allen File to XY Layer
    RouteField = "Call_ID"
    AllenField = "A_ID"

    arcpy.AddJoin_management('SavedRouteAllen', RouteField, 'AllenTable', AllenField, "KEEP_ALL")
    print "Join Allen to XY Layer"

except:
    print arcpy.GetMessages()
    print "Didn't add Allen to XY"
    sys.exit()

try: # Save selected joins
    JoinSavedAllen1 = (FileInput + '/JoinSavedAllen.shp')

    arcpy.Select_analysis('SavedRouteAllen', JoinSavedAllen1)
    print "Saved Allen Join"

except:
    print arcpy.GetMessages()
    print "Didn't Save Allen Join"
    sys.exit

try: # Save selected joins
    JoinSavedAllen2 = (FileInput + '/AcousticRouteAllen.shp')
    AllenRouteData = (FileInput + '/RouteDataAllen.lyr')

    arcpy.MakeFeatureLayer_management('SavedRouteAllen', AllenRouteData)
    print 'Made Acoustic Route a Feature Layer'

    arcpy.Select_analysis('SavedRouteAllen', JoinSavedAllen2)
    print ('Saved shapefile containing Acoustic Route with BCID information joined as: ')
    print (':---> ' + JoinSavedAllen2)
except:
    print arcpy.GetMessages()
    print "Didn't Save Acoustic Route"
    sys.exit

try: # Select Call Locations
     AllenCalls = (FileInput + '/AllenCalls.shp')

     arcpy.Select_analysis(JoinSavedAllen2, AllenCalls, '"ID" > 1')

     print ("Selected Calls")
     print ("Allen Calls Shapefile:-->" + AllenCalls)

except:
     print arcpy.GetMessages()
     print "Didn't Make Selected Call Shapefiles"
     sys.exit


# All Calls Shapefile ##

try: # Join All Calls to XY Layer
    RouteField = "Call_ID"
    AllenField = "A_ID"
    KalField = "K_ID"

    arcpy.AddJoin_management('SavedRoute', RouteField,'AllenTable' , AllenField, "KEEP_ALL")
    print "Join Allen to XY Layer"

    arcpy.AddJoin_management('SavedRoute', RouteField, 'KalTable', KalField, "KEEP_ALL")
    print "Join Kal to XY Layer"

except:
    print arcpy.GetMessages()
    print "Didn't add All to XY"
    sys.exit()

try: # Save selected joins
    AllJoinSaved = (FileInput + '/AllJoinSaved.shp')

    arcpy.Select_analysis('SavedRoute', AllJoinSaved)
    print "Saved All Join"

except:
    print arcpy.GetMessages()
    print "Didn't Save All Join"
    sys.exit

try: # Save selected joins
    AllJoinSaved2 = (FileInput + '/AllCallsRoute.shp')
    AllJoinRouteData = (FileInput + '/RouteDataAll.lyr')

    arcpy.MakeFeatureLayer_management('SavedRoute', AllJoinRouteData)
    print 'Made Acoustic Route All a Feature Layer'

    arcpy.Select_analysis('SavedRoute', AllJoinSaved2)
    print ('Saved shapefile containing Acoustic Route All information joined as: ')
    print (':---> ' + AllJoinSaved2)
except:
    print arcpy.GetMessages()
    print "Didn't Save Acoustic Route All"
    sys.exit

try: # Select Call Locations
     AllCalls = (FileInput + '/AllCalls.shp')

     arcpy.Select_analysis(AllJoinSaved2, AllCalls, '"ID" > 1 OR "KalNum" > 0')

     print ("Selected Calls")
     print ("All Calls Shapefile:-->" + AllCalls)

except:
     print arcpy.GetMessages()
     print "Didn't Make Selected Call Shapefiles"
     sys.exit



try: # Delete Temp Files
    os.remove(FileInput + '/AllenBCID.csv')
    os.remove(FileInput + '/mod1gps.txt')
    os.remove(FileInput + '/mod0gps.txt')
    os.remove(FileInput + '/mod2gps.txt')
    os.remove(FileInput + '/mod3gps.csv')

    #arcpy.Delete_management(JoinSaved)
    #arcpy.Delete_management(JoinSaved1)
    #arcpy.Delete_management(JoinSaved2)
    #arcpy.Delete_management(JoinSaved3)
    #arcpy.Delete_management(Save_route)
    #arcpy.Delete_management(save_Layer)
    os.remove(FileInput + '/schema.ini')

except:
    print arcpy.GetMessages()
    print "Didn't Delete"
    sys.exit


# All done
sys.exit
