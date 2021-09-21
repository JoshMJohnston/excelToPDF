############################################################
#   Author: Josh M. Johnston
#   Date: 09/16/2021
#   
#
############################################################

import sys
import openpyxl
from openpyxl.worksheet import worksheet
from reportlab.platypus.flowables import PageBreak
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate
from reportlab.platypus import Paragraph

QUESTIONS_PER_PAGE = 3

class App:
    def __init__(self, inputFile, outputFile, worksheetNum):
        self.inputFile = inputFile + '.xlsx'
        self.outputFile = outputFile + '.pdf'
        self.worksheetNum = worksheetNum
        self.questions = []

        try:
            self.wb = openpyxl.load_workbook(filename=self.inputFile)
        except:
            print("Error: \"Questions.xlsx\" not found.")
            sys.exit()
        
        self.ws = self.wb.worksheets[self.worksheetNum]
        self.invalidQuestions = ["<Unanswered>", "na", "n/a", "not for right now", "not right now", "no", "not at the moment.", "not right now"]

    def getQuestionColumnNumber(self):
        curCol = 1
        curCellData = self.ws.cell(column=curCol, row=1).value
        while (curCellData != None):
            if curCellData == "Answer 11":
                return curCol
            else:
                curCol += 1
                curCellData = self.ws.cell(column=curCol, row=1).value

    def getQuestions(self, col):
        curRow = 2
        curCellData = self.ws.cell(column=col, row=curRow).value
        while (curCellData != None):
            lowercase = curCellData.lower()
            if lowercase not in self.invalidQuestions:
                curCellData = curCellData.replace('<br>', '')
                curCellData = curCellData.replace('</>', '')
                if (curCellData != None):
                    self.questions.append(curCellData)
            curRow += 1
            curCellData = self.ws.cell(column=col, row=curRow).value

    def generatePDF(self):
        doc = SimpleDocTemplate(self.outputFile, pagesize=letter)
        elements = []
        count = 0
        for question in self.questions:
            count += 1
            p = Paragraph(question)
            elements.append(p)
            if (count < QUESTIONS_PER_PAGE):
                n = Paragraph("<br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/>")
                elements.append(n)
            if (count % QUESTIONS_PER_PAGE == 0):
                elements.append(PageBreak())
                count = 0
            
        doc.build(elements)

    def run(self):
        colNum = self.getQuestionColumnNumber()
        self.getQuestions(colNum)
        self.generatePDF()
