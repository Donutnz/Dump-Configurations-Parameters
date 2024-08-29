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

        app.log("Params Dump Script Started...")
        
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
        
        csvTargetFilePath="{}/{}-PARAMS.csv".format(outputFolderPath,app.activeDocument.name)

        app.log("Dumping to: {}".format(csvTargetFilePath))
        
        # Get favourited paramaters and make headers list
        favParams=[]
        headersList=["Part Name", "Part Number", "Description"]

        for p in design.allParameters:
            if p.isFavorite:
                favParams.append(p.name)
                headersList.append(p.name)

        topTable=design.configurationTopTable

        # Get part number and description column IDs
        partNumberCOLID=""
        descriptionCOLID=""

        for col in topTable.columns:
            if col.title == "Part Number":
                partNumberCOLID = col.id
            elif col.title == "Description":
                descriptionCOLID = col.id

        rowsOutputData=[]

        # Extract parameters
        for confRow in topTable.rows:
            app.log("Starting Row: {}".format(confRow.name))
            confRow.activate()

            rowFParams={}
            rowFParams["Part Name"] = confRow.name
            rowFParams["Part Number"] = confRow.getCellByColumnId(partNumberCOLID).value
            rowFParams["Description"] = confRow.getCellByColumnId(descriptionCOLID).value

            for pName in favParams:
                rowFParams[pName]=adsk.fusion.Parameter.cast(design.allParameters.itemByName(pName)).value

            rowsOutputData.append(rowFParams)

        # Write rows to CSV file
        with open(csvTargetFilePath, 'w', newline='') as csvFile:
            writer=csv.DictWriter(csvFile, fieldnames = headersList)

            writer.writeheader()
            writer.writerows(rowsOutputData)

        ui.messageBox("Done! Wrote {} rows to {}".format(len(rowsOutputData), csvTargetFilePath))
        app.log("Done")

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
