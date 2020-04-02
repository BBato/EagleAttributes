
import pdfkit
import requests
import json
import os
import time

errors = []


def getFarnell( digikey_code ):

    if digikey_code == "-" or digikey_code.strip()=="": return ""

    print("Requesting component '"+digikey_code+"' from Farnell...")
    response = requests.request("GET", "http://api.element14.com//catalog/products?term=any%3A"+digikey_code+"&storeInfo.id=ie.farnell.com&resultsSettings.offset=0&resultsSettings.numberOfResults=1&resultsSettings.refinements.filters=&amp;resultsSettings.responseGroup=small&callInfo.omitXmlSchema=false&callInfo.callback=&callInfo.responseDataFormat=json&callinfo.apiKey=")
    response_dict = json.loads(response.content)
    
    try:
        code = response_dict["keywordSearchReturn"]["products"][0]["sku"]
        print("Success.")
    except:
        code = "Not Found"
        print("[ERROR] "+response.text)

    time.sleep(0.5)
    return code