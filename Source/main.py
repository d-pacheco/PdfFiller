from Config import Config
import argparse
import sys
import re
import openpyxl
import pdfrw
from PathFinder import *
from rich import print
from os import listdir
from os.path import isfile, join


CURRENT_VERSION = 1.0

def init():
    parser = argparse.ArgumentParser(description="Description here")
    parser.add_argument('-c', '--config', dest='configPath', default='./config/config.yaml', help='Path to a custom config file')
    parser.add_argument('-p', '--templates', dest='templatesPath', default='./templates', help='Path to all stored templates')
    parser.add_argument('-d', '--data', dest='dataPath', default='./data', help='Path to all stored data')
    args = parser.parse_args()
    config = Config(args.configPath)
    templatePath = findTemplates(args.templatesPath)
    dataPath = findTemplates(args.dataPath)
    return config, templatePath, dataPath


def main(config: Config, templatePath: str, dataPath: str):
    templateName = GetTemplate(templatePath)
    if templateName != None:
        pdf_path = "{template_path}/{template_name}".format(template_path=templatePath, template_name=templateName)
        fill_pdf_form(pdf_path)

def GetTemplate(templatePath):
    templateFiles = [f for f in listdir(templatePath) if isfile(join(templatePath, f))]
    if len(templateFiles) != 0:
        print("Select which file you would like to generate from template:")
        index = 1
        for templateFile in templateFiles:
            print("\t{index}. {fileName}".format(index=index, fileName=templateFile))
            index += 1

        validInput = False
        while not validInput:
            userInput = input("Select an index of a template file: ")
            try:
                index = int(userInput)
                selectedTemplate = templateFiles[index - 1]
                validInput = True
            except Exception:
                print(f"[red]The selection you made was invalid. Please try again.")
        return selectedTemplate
    else:
        print(f"[red]There are no template files available to select from")
        return None

def fill_pdf_form(pdf_path):
    template_pdf = pdfrw.PdfReader(pdf_path)
    wb = openpyxl.load_workbook('../Data/Portfolio Creator Redemption Calculator.xlsm', data_only=True)
    
    annotations = template_pdf.pages[0]['/Annots']
    for annotation in annotations:
        if annotation['/Subtype'] == '/Widget':
            field = checkAnnotation(annotation)
            if field != None:
                cellValue = GetValueForField(annotation, field, wb)
                annotation.update(pdfrw.PdfDict(V='{}'.format(cellValue)))
            else:
                continue

    output_pdf_path = "output.pdf"  # Path to the output filled-in PDF file
    pdfrw.PdfWriter().write(output_pdf_path, template_pdf)
    

def checkAnnotation(annotation):
    for field in config.GetFieldsForFile("Redemption Form"):
        if annotation['/TU'] != None and field in annotation['/TU']:
            return field
    return None


def GetValueForField(annotation, field, wb):
    sheetName, cell = config.GetSheetAndCell("Redemption Form", field)
    sheet = wb[sheetName]
    index = re.findall(r'\d+', annotation['/T'])[0]
    if "-" in cell:
        cell = GetCellFromRange(cell, int(index))
    return sheet[cell].value


def GetCellFromRange(rangeStr, index):
    startCell, endCell = rangeStr.split('-')
    column = startCell[:1]
    startRow = int(startCell[1:])
    endRow = int(endCell[1:])
    offset = index % (endRow - startRow + 1)
    row = startRow + offset
    return f"{column}{row}"

        

if __name__ == '__main__':
    log = None
    try:
        config, templatePath, dataPath = init()
        main(config, templatePath, dataPath)
    except (KeyboardInterrupt, SystemExit):
        print('Exiting...')
        sys.exit()
