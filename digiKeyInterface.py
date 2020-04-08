
import pdfkit
import requests
from bs4 import BeautifulSoup
import os
import math

errors = []

headers = {
  'Cookie': 'name=value;'
}

options = {
    'cookie': [
        ("name","value"),
    ]
}


def getAlternativeResistor( resistance, package, printToConsole ):


    mul = -1

    if "k" in resistance or "K" in resistance:
        res = resistance.split("k")
        mul = 1000

    if "R" in resistance:
        res = resistance.split("R")
        mul = 1

    if "meg" in resistance.lower():
        res = resistance.lower().split("meg")
        mul = 1000000
    else:
        if "M" in resistance:
            res = resistance.split("M")
            mul = 1000000
        else:
            if "m" in resistance:
                res = resistance.split("m")
                mul = 0.001

    if mul == -1:
        res = []
        res.append(resistance)
        mul = 1
        
    
    resistanceValue = float(res[0])*mul
    if len(res)>1:
        if res[1] is not '':
            #print(len(res))
            resistanceValue += float(res[1])/math.pow(10,len(str(res[1])))*mul

    #print(resistanceValue)

    minR = float(resistanceValue)
    maxR = float(resistanceValue)
    pv = ""

    if "0402" in package:
        pv = "&pv16=39158"

    if "0603" in package:
        pv = "&pv1291=39245"

    if "0805" in package:
        pv = "&pv1291=39328"

    if "1206" in package:
        pv = "&pv1291=83708"

    if pv == "":
        printToConsole("[ERROR] Unsupported Package: "+package)

    for i in range(0,10):

        req = "https://www.digikey.com/products/en/resistors/chip-resistor-surface-mount/52?k=&stock=1&v=13&pv7=2&pv1989=0&ColumnSort=1000011&umin2085="+str(minR)+"&umax2085="+str(maxR)+"&rfu2085=Ohms&page=1&pageSize=500&pv3=731&pv3=1131"+pv
        #print(req)
        response = requests.request("GET", req, headers=headers)
        
        soup = BeautifulSoup(response.text, "lxml")

        try:
            noResults = soup.find("div", { "id" : "noResults" }).prettify()
            printToConsole("Searching for more similar resistors...")

            minR = minR * 0.99
            maxR = maxR * 1.01
        except:
            try:
                #productTable = soup.find("table", class_="productTable").prettify()
                tableEntry = soup.find_all("td", class_="tr-dkPartNumber")
                dk = tableEntry[0].find("a").getText().strip()
                printToConsole("Found alternative resistor: "+dk)
                return dk
            except:
                tab = soup.find("table", { "id" : "product-overview" }).find_all("th")

                for t in tab:
                    if "Digi-Key Part Number" in t.getText():
                        dk = t.parent.find("td").getText().strip()
                        printToConsole("Found alternative resistor: "+dk)
                        return dk






def getDigiKeyReference( digikey_code, optGenerateSpecSheets, printToConsole, epn_code, regenerateEPN ):


    # Check if Digikey Code was supplied
    if digikey_code == "-" or digikey_code=="" or digikey_code is None:

        # No code - return nothing
        printToConsole("[WARNING] No Digikey code was specified for this component.")
        return ["-","-","-","-","-","-"]

    # Begin researching component
    printToConsole("Requesting component '"+digikey_code+"' from DigiKey")

    link = "https://www.digikey.ie/products/en?keywords="+digikey_code

    # Try get request by direct search
    response = requests.request("GET", link, headers=headers)
    soup = BeautifulSoup(response.text, "lxml")

    try:

        # Check if the page is a selection of products
        p = soup.find("table", class_="productTable").prettify()

        # Since no exception occurred then a product table exists.
        printToConsole("This page lists multiple products. Redirecting.")

        # Read the product table and redirect to the first link
        tableEntry = soup.find_all("td", class_="tr-dkPartNumber")
        link = "https://www.digikey.com"+tableEntry[0].find("a").get("href")
        response = requests.request("GET", link, headers=headers)
        soup = BeautifulSoup(response.text, "lxml")

    except:

        # The page had no product table.
        try:

            # Check if the page contains an exact component table
            p = soup.find("table", class_="exactPart")
            link = "https://www.digikey.com"+p.find("a").get("href")

            # Exact component table confirmed - redirect to first link
            response = requests.request("GET", link, headers=headers)
            soup = BeautifulSoup(response.text, "lxml")            

        except:

            # The page already contains component data, no redirect necessary.
            pass


    try:

        # Get details element and replace components to prepare them for PDF generation.
        details = soup.find("div", class_="product-details-overview").prettify().replace("href=\"/","href=\"https://www.digikey.com/").replace("Report an Error", "").replace("View Similar","").replace("//media.digikey.com", "https://media.digikey.com")
    
    except:
        
        # The details element could not be located for some reason.
        details = None
        printToConsole("[ERROR] Could not find component "+digikey_code+" !")

        # Since no details were retrieved, the search is worthless - return.
        return ["-","-","-","-","-","-"]

    
    try:

        # Get photo element and replace components to prepare them for PDF generation.
        photo = soup.find("a", class_="product-photo-large").prettify().replace("//media.digikey.com", "https://media.digikey.com")

    except:

        # The photo element could not be located for some reason. No need to quit because we have component details.
        photo = None
        printToConsole("[WARNING] Could not load image for "+digikey_code)


    # Prepare attributes that are expected for return
    manufacturer = "-"
    manufacturerPartNumber = "-"
    availability = "-"
    description = "-"
    height = "-"
    package = "-"
    resistance = "ERROR"

    try:

        # Determine Component Mounting Type
        tab = soup.find("table", { "id" : "product-attribute-table" }).prettify()

        # Determine if component is Through Hole or Surface Mount
        if "Through Hole" in tab:
            mountingType = "TH"
        else:
            if "Surface Mount" in tab:
                mountingType = "SMD"
            else:
                mountingType = "-"
        
        # Determine Component Attributes from the Long Table
        tab = soup.find("table", { "id" : "product-attribute-table" }).find_all("th")

        # Extract the components
        for t in tab:
            if t.getText().strip() == "Manufacturer":
                manufacturer = t.parent.find("td").getText().strip()
            if "Height" in t.getText().strip():
                height = t.parent.find("td").getText().strip()
            if "Resistance" in t.getText().strip():
                resistance = t.parent.find("td").getText().strip().replace(" MOhms","Meg").replace(" kOhms","k").replace(" Ohms","R")
            if "Supplier Device Package" in t.getText().strip():
                package = t.parent.find("td").getText().strip()

        ### Determine Component Attributes from the Short Table
        tab = soup.find("table", { "id" : "product-overview" }).find_all("th")

        # Extract the components
        for t in tab:
            if "Manufacturer Part Number" in t.getText():
                manufacturerPartNumber = t.parent.find("td").getText().strip()
            if "Description" in t.getText():
                description = t.parent.find("td").getText().strip()     


    except:

        # Fatal error - not expected to fail here, but for edge cases allow exceptions.
        return("Error with Digikey webpage.")


    try:

        # Determine Component Availability
        availability = soup.find("div", class_="quantity-message").getText().replace("\n"," ").strip()

    except:

        # Availability element was not found on this page.
        printToConsole("[WARNING] This product does not seem to be available for purchase!")
        availability = "None"


    # Check if the component changed and determine if a new EPN is required (resistors only)
    if regenerateEPN:
        finalEPN = generateResistorEPN( description, package, resistance, printToConsole )
    else:
        finalEPN = epn_code

    if finalEPN == "":
        newName=digikey_code
        printToConsole("[WARNING] No EPN code was specified for component "+digikey_code)
    else:
        newName=finalEPN


    if optGenerateSpecSheets:
        
        # Build specification document for this component
        style = open("resources/style.html", "r")
        file = open("output/"+newName+".html","w")
        file.write("<html><head>")
        file.write( style.read().replace('\n', '') )
        if photo is not None: file.write(photo)
        if details is not None:
            details = details.replace("</h1>", "</h1><a style=\"font-size: smaller;\">&emsp;&emsp;EPN: "+finalEPN+"</a>")
            file.write(details)
        file.write("</body></html")
        file.close()
        style.close()

        # Turn off console 
        optionsPDF = {
            'quiet': ''
        }

        pdfkit.from_file("output/"+newName+".html", "output/"+newName+".pdf", options=optionsPDF)
        os.remove("output/"+newName+".html")


    return [manufacturer,manufacturerPartNumber,availability,description,mountingType,resistance,finalEPN,link]






def generateResistorEPN( description, package, resistance, printToConsole ):

    if "0.1%" in description:
        prefix = "RP"
    else:
        if "1%" in description:
            prefix = "RQ"
        else:
            if "5%" in description:
                prefix = "RR"
            else:
                prefix = "RS"

    suffix = "X1"
    if "0805" in package:
        suffix = "R1"
    if "1206" in package:
        suffix = "R2"
    if "0603" in package:
        suffix = "R6"
    if "0402" in package:
        suffix = "R7"

    if "k" in resistance:
        res = resistance.split("k")
        mul = 1000
    else:
        if "Meg" in resistance:
            res = resistance.split("Meg")
            mul = 1000000
        else:
            if "GOhms" in resistance:
                res = resistance.split(" GOhms")
                mul = 1000000000
            else:
                if "R" in resistance:
                    res = resistance.split("R")
                else:
                    printToConsole("[ERROR] Can't determine resistor value and EPN.")
                    return " "


    resistanceValue = float(res[0])
    
    while str(resistanceValue).split(".")[1] is not "0":
        resistanceValue = resistanceValue * 10
        mul = mul / 10

    mult = "X"
    if mul == 1: mult = "0"
    if mul == 10: mult = "1"
    if mul == 100: mult = "2"
    if mul == 0.1: mult = "A"
    if mul == 0.01: mult = "B"
    if mul == 0.001: mult = "C"
    if mul == 1000: mult = "3"
    if mul == 10000: mult = "4"
    if mul == 100000: mult = "5"
    if mul == 1000000: mult = "6"
    if mul == 0.0001: mult = "D"
    if mul == 0.00001: mult = "E"

    return prefix+"-"+str(resistanceValue).split(".")[0]+mult+"-"+suffix