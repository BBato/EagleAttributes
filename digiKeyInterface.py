
import pdfkit
import requests
from bs4 import BeautifulSoup
import os

errors = []

headers = {

}

options = {
    
}


def getDigiKeyReference( digikey_code, create_spec_sheets, newName, printToConsole, epn_code ):

    if digikey_code == "-" or digikey_code.strip()=="":
        printToConsole("[WARNING] No Digikey code was specified for this component.")
        return ["-","-","-","-","-"]

    if newName == "": newName=digikey_code

    if epn_code == "":
        epn_code = "Not specified"
        printToConsole("[WARNING] No EPN code was specified for this component.")

    #print("Requesting component '"+digikey_code+"' from DigiKey")
    printToConsole("Requesting component '"+digikey_code+"' from DigiKey")
    response = requests.request("GET", "https://www.digikey.ie/products/en?keywords="+digikey_code, headers=headers)
    
    soup = BeautifulSoup(response.text, "lxml")

    ## Check if the page is a selection of products
    productTable = soup.find("table", class_="productTable")

    try:
        p = productTable.prettify()
        #print("This page lists multiple products. Redirecting.")
        printToConsole("This page lists multiple products. Redirecting.")
        tableEntry = soup.find_all("td", class_="tr-dkPartNumber")
        link = "https://www.digikey.com"+tableEntry[0].find("a").get("href")
        #print("Going to "+link)
        response = requests.request("GET", link, headers=headers)
        soup = BeautifulSoup(response.text, "lxml")

    except:
        print("Success.")
        printToConsole("This page lists multiple products. Redirecting.")

    try:
        details = soup.find("div", class_="product-details-overview").prettify().replace("href=\"/","href=\"https://www.digikey.com/").replace("Report an Error", "").replace("View Similar","").replace("//media.digikey.com", "https://media.digikey.com")
        details = details.replace("</h1>", "</h1><a style=\"font-size: smaller;\">&emsp;&emsp;EPN: "+epn_code+"</a>")
    except:
        details = None
        #print("[ERROR] Could not find component "+digikey_code+" !")
        printToConsole("[ERROR] Could not find component "+digikey_code+" !")
        return ["-","-","-","-","-"]

    
    try:
        photo = soup.find("a", class_="product-photo-large").prettify().replace("//media.digikey.com", "https://media.digikey.com")
    except:
        photo = None
        #print("[WARNING] Could not load image for "+digikey_code)
        printToConsole("[WARNING] Could not load image for "+digikey_code)

#    try:
#        pageText = soup.find("td", itemprop="description").getText()
#        print(pageText.strip())
#    except:
#        print("[ERROR] Failed to load product summary!")

    

    manufacturer = "-"
    manufacturerPartNumber = "-"
    availability = "-"
    description = "-"
    height = "-"

    try:
        ### Determine Component Mounting Type

        tab = soup.find("table", { "id" : "product-attribute-table" }).prettify()

        if "Through Hole" in tab:
            mountingType = "TH"
        else:
            if "Surface Mount" in tab:
                mountingType = "SMD"
            else:
                mountingType = "-"
        
        
        ### Determine Component Attributes - Long Table

        tab = soup.find("table", { "id" : "product-attribute-table" }).find_all("th")

        for t in tab:
            if t.getText().strip() == "Manufacturer":
                manufacturer = t.parent.find("td").getText().strip()
            if "Height" in t.getText().strip():
                height = t.parent.find("td").getText().strip()


        ### Determine Component Attributes - Short Table

        tab = soup.find("table", { "id" : "product-overview" }).find_all("th")

        for t in tab:
            if "Manufacturer Part Number" in t.getText():
                manufacturerPartNumber = t.parent.find("td").getText().strip()
            if "Description" in t.getText():
                description = t.parent.find("td").getText().strip()        
    except:
        return("Error with Digikey webpage.")


    ### Determine Component Availability

    try:
        availability = soup.find("div", class_="quantity-message").getText().replace("\n"," ").strip()
    except:
        #print("[WARNING] This product does not seem to be available for purchase!")
        printToConsole("[WARNING] This product does not seem to be available for purchase!")
        availability = "None"



    if create_spec_sheets:
        
        style = open("resources/style.html", "r")
        file = open("output/"+newName+".html","w")
        file.write("<html><head>")
        file.write( style.read().replace('\n', '') )
        if photo is not None: file.write(photo)
        if details is not None: file.write(details)
        file.write("</body></html")
        file.close()
        style.close()

        pdfkit.from_file("output/"+newName+".html", "output/"+newName+".pdf")
        os.remove("output/"+newName+".html")

    return [manufacturer,manufacturerPartNumber,availability,description,mountingType]

