#Author Donutnz
#Description Dump favourited parameters for each configuration row into a CSV.

import adsk.core, adsk.fusion, adsk.cam, traceback
import csv

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface


        design=adsk.fusion.Design.cast(app.activeProduct)

        # Check if this is a configured design.
        if not design.isConfiguredDesign:
            ui.messageBox("Error: The current design is not configured! This script can only run on configured designs.")
            return

        outputFolderPath=""
        folderDlg=ui.createFolderDialog()
        folderDlg.title="Choose Output Folder"

        dlgResult=folderDlg.showDialog()
        if(dlgResult == adsk.core.DialogResults.DialogOK):
            outputFolderPath=folderDlg.folder
        else:
            return
        
        csvTargetFilePath="{}/{}.csv".format(outputFolderPath,app.activeDocument.name)
        
        favParams=[]

        # Get fav paramaters
        for p in design.allParameters:
            if p.isFavorite:
                favParams.append(p.name)

        rowParamValues=[]

        topTable=design.configurationTopTable

        # Get part number and description column IDs
        partNumberCOLID=""
        descriptionCOLID=""

        for col in topTable.columns:
            if col.title == "Part Number":
                partNumberCOLID = col.id
            elif col.title == "Description":
                descriptionCOLID = col.id

        # Extract parameters
        for confRow in topTable.rows:
            app.log("Starting Row: {}".format(confRow.name))
            confRow.activate()

            rowFParams={}
            rowFParams["partName"]=confRow.name
            rowFParams["partNumber"]=confRow.getCellByColumnId(partNumberCOLID)
            rowFParams["description"] = confRow.getCellByColumnId(descriptionCOLID)

            for pName in favParams:
                rowFParams[pName]=design.allParameters.itemByName(pName)

            rowParamValues.append(rowFParams)

        # Add properties for use in header
        favParams.append("partName")
        favParams.append("partNumber")
        favParams.append("description")

        # Write rows to CSV file
        with open(csvTargetFilePath, 'w', newline='') as csvFile:
            writer=csv.DictWriter(csvFile, fieldnames = favParams)

            writer.writeheader()
            writer.writerows(rowParamValues)

        ui.messageBox("Done! Wrote {} rows to {}".format(rowParamValues.count(), csvTargetFilePath))
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
