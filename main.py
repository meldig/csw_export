import json
import xml.etree.ElementTree as ET
import xlsxwriter
import requests

# Constantes
NOM_FICHIER_DIG = "catalogue_interne_dig.xlsx"
NOM_FICHIER_AGENTS = "catalogue_agents_mel.xlsx"


with open("D:\\Documents\\MISSIONS\\Mission_1\\config.json") as jsonFile:
    jsonObject = json.load(jsonFile)
    jsonFile.close()


def getRecordById(id):
    r = requests.get(jsonObject["RECORDBYID"] + id, verify=False)
    return r


def getRecords():
    with open(jsonObject["PATHS"]["REQUEST"]) as xml:
        responseCatalogueDIG = requests.post(jsonObject["DIG"]["RECORDS"], data=xml.read(),
                                             headers={"Content-Type": "text/xml"}, verify=False)

    with open(jsonObject["PATHS"]["REQUEST"]) as xml:
        responseCatalogueAgents = requests.post(jsonObject["AGENTS"]["RECORDS"], data=xml.read(),
                                                headers={"Content-Type": "text/xml"}, verify=False)

    rootCatalogueDIG = ET.fromstring(responseCatalogueDIG.content.decode("utf-8"))
    rootCatalogueAgents = ET.fromstring(responseCatalogueAgents.content.decode("utf-8"))

    recordsCatalogueDIG = rootCatalogueDIG.findall(jsonObject["PATHS"]["RECORD_TAG_PATH"])
    recordsCatalogueAgents = rootCatalogueAgents.findall(jsonObject["PATHS"]["RECORD_TAG_PATH"])

    WriteInExcelFile(recordsCatalogueDIG, NOM_FICHIER_DIG)
    WriteInExcelFile(recordsCatalogueAgents, NOM_FICHIER_AGENTS)


def WriteInExcelFile(records, fileName):
    row = 1
    workbook = xlsxwriter.Workbook(jsonObject["PATHS"]["EXCEL"] + fileName)
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, 'Titre', workbook.add_format({'bold': True}))
    worksheet.write(0, 1, 'URL', workbook.add_format({'bold': True}))
    worksheet.write(0, 2, 'Export XML', workbook.add_format({'bold': True}))
    worksheet.set_column(0, 0, 140)
    worksheet.set_column(1, 1, 120)
    worksheet.set_column(2, 2, 120)

    if fileName == NOM_FICHIER_DIG:
        url = jsonObject["DIG"]["FICHE"]
    elif fileName == NOM_FICHIER_AGENTS:
        url = jsonObject["AGENTS"]["FICHE"]

    for record in records:
        id = record.find(jsonObject["PATHS"]["IDENTIFIER"])
        title = record.find(jsonObject["PATHS"]["TITLE"])

        worksheet.write(row, 0, title.text)
        worksheet.write(row, 1, url + id.text.split(':')[-1])

        row += 1

    workbook.close()


getRecords()