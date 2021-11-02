import json
import xml.etree.ElementTree as ET
import xlsxwriter
import requests

# Ouverture et chargement du fichier de configuration json contenant les Urls qui sera utilisé via la variable jsonObject
with open("chemin absolu du fichier de configuration") as jsonFile:
    jsonObject = json.load(jsonFile)
    jsonFile.close()


# Fonction qui permet de récupérer toutes les fiches des deux catalogues ainsi que leurs informations
def getRecords(url):
    with open(jsonObject[1]["paths"]["nom_du_fichier"]) as xml:
        response = requests.post(url, data=xml.read(),
                                             headers={"Content-Type": "text/xml"}, verify=False)

    root = ET.fromstring(response.content.decode("utf-8"))

    return root.findall(jsonObject[2]["tags"]["nom_balise"])


# Fonction qui permet de créer les fichiers excel et d'écrire les informations sur les fiches dedans
def writeInExcelFile(records, filePath, url):
    row = 1
    workbook = xlsxwriter.Workbook(filePath)
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, 'Titre', workbook.add_format({'bold': True}))
    worksheet.write(0, 1, 'URL', workbook.add_format({'bold': True}))
    worksheet.set_column(0, 0, 140)
    worksheet.set_column(1, 1, 120)

    for record in records:
        id = record.find(jsonObject[2]["tags"]["nom_balise"])
        title = record.find(jsonObject[2]["tags"]["nom_balise"])

        worksheet.write(row, 0, title.text)
        worksheet.write(row, 1, url + id.text.split(':')[-1])

        row += 1

    workbook.close()


records = getRecords(jsonObject[0]["urls"]["records"])
writeInExcelFile(records, jsonObject[2]["paths"]["nom_du_fichier"], jsonObject[0]["urls"]["fiche"])
