import csv
import openpyxl
import digiKeyInterface
import farnell_interface
import openpyxl
from openpyxl.styles import Font, Border, Side



def execute_main(inputFileName, inputMasterEPN, optGetFarnellCodes, optGenerateSpecSheets, optGenerateEagleScript, optGenerateBOM, progressbar, printToConsole):

    try:
        with open(inputFileName) as f:
            reader = csv.reader(f, delimiter=';')
            data = list(reader)
    except:
        return "Cannot open file: "+inputFileName

    def col(header):
        for x in range(0, len(data[0])-1):
            if data[0][x] == header:
                return x
        return -1

    colParts = col("Parts")
    colDigikey = col("DIGIKEY_PARTNUM")
    colQty = col("Qty")
    colValue = col("Value")
    colDevice = col("Device")
    colDesc = col("Description")
    colPackage = col("Package")
    colMf = col("MANUFACTURER_NAME")
    colEPN = col("EPN")

    if colDigikey == -1:
        return "[ERROR] No attribute 'DIGIKEY_PARTNUM' found. Add this attribute to some of your components in Eagle."

    if colEPN == -1:
        return "[ERROR] No attribute 'EPN' found. Add this attribute to some of your components in Eagle."    

    numThroughHoleComponents = 0
    numSurfaceMountComponents = 0
    numUniqueSurfaceMountComponents = 0
    numUniqueParts = 0

    if optGenerateEagleScript:
        eagleFile = open("output/eagle.txt", "w")
    
    if optGenerateBOM:
        wb = openpyxl.Workbook()
        sheet = wb.active

        sheet.append(["BOM"])
        sheet.append(['EPN','Qty','Value','Device','Package','Parts','Description','DIGIKEY_PARTNUM','FARNELL_OC','MANUFACTURER','MANUFACTURER_PARTNUM','SUPPLIER','TECH'])


    for x in range(1,len(data)):

        progressbar["value"]=(x/len(data))*100
        progressbar.update()

        components = str(data[x][colParts]).split(",")
        digiKeyCode = str(data[x][colDigikey])
        
        if digiKeyCode == "":
            if optGenerateBOM:
                sheet.append([data[x][colEPN], data[x][colQty], data[x][colValue] ,data[x][colDevice], data[x][colPackage], data[x][colParts], data[x][colDesc], "-", "-", data[x][colMf], "-", "-", "-"])

            continue

        digiKeyData = digiKeyInterface.getDigiKeyReference(digiKeyCode, optGenerateSpecSheets, data[x][colEPN], printToConsole, data[x][colEPN] )
        farnellCode = "Farnell"

        if optGetFarnellCodes:
            farnellCode = farnell_interface.getFarnell(digiKeyData[1])

        if optGenerateBOM:
            sheet.append([data[x][colEPN], data[x][colQty], data[x][colValue] ,data[x][colDevice], data[x][colPackage], data[x][colParts], digiKeyData[3], data[x][colDigikey], farnellCode, digiKeyData[0], digiKeyData[1], "-", digiKeyData[4]])

        numUniqueParts = numUniqueParts + 1

        if digiKeyData[4] == "TH":
            numThroughHoleComponents = numThroughHoleComponents + len(components)
        if digiKeyData[4] == "SMD":
            numSurfaceMountComponents = numSurfaceMountComponents + len(components)
            numUniqueSurfaceMountComponents = numUniqueSurfaceMountComponents + 1

        if optGenerateEagleScript:

            for c in components:
                
                partName = c.strip()
                eagleFile.write("ATTRIBUTE "+partName+" MANUFACTURER_NAME '"+digiKeyData[0]+"'; \n")
                eagleFile.write("CHANGE DISPLAY OFF; \n")
                eagleFile.write("ATTRIBUTE "+partName+" MANUFACTURER_PART_NUMBER '"+digiKeyData[1]+"'; \n")
                eagleFile.write("CHANGE DISPLAY OFF; \n")
                eagleFile.write("ATTRIBUTE "+partName+" DK_AVAILABILITY '"+digiKeyData[2]+"'; \n")
                eagleFile.write("CHANGE DISPLAY OFF; \n")
                eagleFile.write("ATTRIBUTE "+partName+" DK_DESCRIPTION '"+digiKeyData[3]+"'; \n")
                eagleFile.write("CHANGE DISPLAY OFF; \n")

                if optGetFarnellCodes:
                    eagleFile.write("ATTRIBUTE "+partName+" OC_FARNELL '"+farnellCode+"'; \n")
                    eagleFile.write("CHANGE DISPLAY OFF; \n")


    if optGenerateEagleScript:
        eagleFile.close()


    ### Generate Assembly Info Table

    

    if optGenerateBOM:
        dims = {}

        for row in sheet.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
        for col, value in dims.items():
            sheet.column_dimensions[col].width = max(value,5)

        fontStyle = Font(size = "18")
        sheet.cell(row = 1, column = 1, value = inputMasterEPN+' BOM').font = fontStyle

        # define main table style
        mediumStyle = openpyxl.worksheet.table.TableStyleInfo(name='TableStyleMedium4',
                                                            showRowStripes=True)

        # create main table
        table = openpyxl.worksheet.table.Table(ref="A2:M"+str(len(data)+1),
                                            displayName='BOM',
                                            tableStyleInfo=mediumStyle)
        # add the table to the worksheet
        sheet.add_table(table)

        # merge main table header
        sheet.merge_cells("A1:M1")

        # create assembly info table header
        sheet.cell(row = len(data)+3, column = 1, value = 'Assembly Info').font = fontStyle

        sheet.cell(row = len(data)+4, column = 1, value = 'Assembly Options')
        sheet.cell(row = len(data)+5, column = 1, value = 'Unique Part Count')
        sheet.cell(row = len(data)+6, column = 1, value = 'SMD Unique Part Count')
        sheet.cell(row = len(data)+7, column = 1, value = 'SMD Total Part Count')
        sheet.cell(row = len(data)+8, column = 1, value = 'Through Hole Part Count')

        sheet.cell(row = len(data)+5, column = 4, value = numUniqueParts)
        sheet.cell(row = len(data)+6, column = 4, value = numUniqueSurfaceMountComponents)
        sheet.cell(row = len(data)+7, column = 4, value = numSurfaceMountComponents)
        sheet.cell(row = len(data)+8, column = 4, value = numThroughHoleComponents)

        def allRoundBorder():
            return Border(top = Side(border_style='thin', color='FF000000'),    
                              right = Side(border_style='thin', color='FF000000'), 
                              bottom = Side(border_style='thin', color='FF000000'),
                              left = Side(border_style='thin', color='FF000000'))


        for y in range(3,9):
            for x in range(1,5):
                sheet.cell(row=len(data)+y, column=x).border = allRoundBorder() 




        # format assembly info table
        #table = openpyxl.worksheet.table.Table(ref="A"+str(len(data)+4)+":D"+str(len(data)+8),
        #                                    displayName='AssemblyInfo',
        #                                    tableStyleInfo=mediumStyle,
        #                                    h)
        # add the table to the worksheet
        #sheet.add_table(table)

        # merge assembly info table header
        sheet.merge_cells("A"+str(len(data)+3)+":D"+str(len(data)+3))

        # merge assembly info table body
        sheet.merge_cells("A"+str(len(data)+4)+":C"+str(len(data)+4))
        sheet.merge_cells("A"+str(len(data)+5)+":C"+str(len(data)+5))
        sheet.merge_cells("A"+str(len(data)+6)+":C"+str(len(data)+6))
        sheet.merge_cells("A"+str(len(data)+7)+":C"+str(len(data)+7))
        sheet.merge_cells("A"+str(len(data)+8)+":C"+str(len(data)+8))


        wb.save("output/"+inputMasterEPN+".xlsx")

    
    return "Complete."