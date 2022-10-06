import os
import win32com.client

documentos = os.listdir('Asignaciones')

# print(documentos)

wdFormatPDF = 17

for doc in documentos:

    inputFile = os.path.abspath("Asignaciones/" + doc)
    outputFile = os.path.abspath(doc + ".pdf")
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(inputFile)
    doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

