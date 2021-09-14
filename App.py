from re import template
import openpyxl
from openpyxl.worksheet import worksheet
from Person import Person
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle


class App:
    def __init__(self, inputFile, outputFile, row, columns, worksheetNum):
        self.inputFile = inputFile + '.xlsx'
        self.outputFile = outputFile + '.pdf'
        self.row = row
        self.columns = columns
        self.worksheetNum = worksheetNum
        self.people = []

    def getExcelData(self):
        wb = openpyxl.load_workbook(filename=self.inputFile)
        ws = wb.worksheets[self.worksheetNum]

        curRow = self.row
        while (ws.cell(column=self.columns[0], row=curRow).value != None):
            temp = []
            for col in self.columns:
                temp.append(ws.cell(column=col, row=curRow).value)
            self.people.append(temp)
            curRow += 1
            print(temp)

    def generatePDF(self):
        doc = SimpleDocTemplate(self.outputFile, pagesize=letter)
        elements = []
        t = Table(self.people)
        t.setStyle(TableStyle([('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                               ('BACKGROUND', (0, 0), (-1, 0), colors.gray),
                               ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black)]))
        elements.append(t)
        doc.build(elements)

    def run(self):
        self.getExcelData()
        self.generatePDF()
