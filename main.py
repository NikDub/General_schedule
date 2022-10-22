import re
import docx
from docxtpl import DocxTemplate

context = {}
replaceName = []
scheduleList = []

fileLectorName = open("templ/Имена преподователей.txt", "r", 256, "utf-8")
lectorList = fileLectorName.read().splitlines()
fileLectorName.close()

fileFileName = open("templ/Файлы.txt", "r", 256, "utf-8")
fileList = fileFileName.read().splitlines()
fileFileName.close()

print("Ожидайте обрабатываем файлы ...")
for fileName in fileList:
    doc = docx.Document('docx/'+fileName[:fileName.find(" ")])
    table = doc.tables[int(fileName[fileName.find(" "):].strip())]
    previousString = ' '
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            if cell.text != "":
                for lectorName in lectorList:
                    if cell.text.lower().find(lectorName[:4].lower()) != -1:
                        if cell.text.lower().find(lectorName[-4:].lower()) != -1:
                            resultString = table.rows[i].cells[0].text[0:3] + table.rows[i].cells[1].text[0:2].replace(":", "")
                            match = re.compile('w:fill=\"(\S*)\"').search(cell._tc.xml)
                            if match:
                                if match.group(1) in ['FFFFFF', 'auto']:
                                    resultString += 'Б' + lectorName[:4] + lectorName[-4:].replace(".", "")
                                else:
                                    resultString += 'З' + lectorName[:4] + lectorName[-4:].replace(".", "")

                            resultString += ' ' + table.rows[0].cells[j].text + ' '

                            resultString += cell.text.replace(".", "").replace(lectorName.replace(".", ""), "")

                            resultString = resultString.replace("\n", " ").replace("   ", " ").replace("  ", " ")

                            if previousString != resultString:
                                previousString = resultString
                                if len(table.rows) >= i+2:
                                    if table.rows[i].cells[1].text == table.rows[i + 1].cells[1].text and \
                                            table.rows[i + 1].cells[j].text == cell.text or \
                                            table.rows[i].cells[1].text == table.rows[i - 1].cells[1].text and \
                                            table.rows[i - 1].cells[j].text == cell.text:
                                        scheduleList.append(resultString.replace("Б"+lectorName[:4], "З"+lectorName[:4], 1))

                                    elif table.rows[i - 1].cells[j].text != cell.text and \
                                            table.rows[i + 1].cells[j].text != cell.text and \
                                            table.rows[i - 1].cells[1].text != table.rows[i].cells[1].text and \
                                            table.rows[i + 1].cells[1].text != table.rows[i].cells[1].text:
                                        scheduleList.append(resultString.replace("Б"+lectorName[:4], "З"+lectorName[:4], 1))
                                else:
                                    if table.rows[i].cells[1].text == table.rows[i - 1].cells[1].text and \
                                            table.rows[i - 1].cells[j].text == cell.text:
                                        scheduleList.append(resultString.replace("Б"+lectorName[:4], "З"+lectorName[:4], 1))

                                    elif table.rows[i - 1].cells[j].text != cell.text and \
                                            table.rows[i - 1].cells[1].text != table.rows[i].cells[1].text:
                                        scheduleList.append(resultString.replace("Б"+lectorName[:4], "З"+lectorName[:4], 1))

                                scheduleList.append(resultString)

print("Все файлы обработаны, записываем итоговый файл ...")
doc = DocxTemplate("templ/ШаблонРасписания.docx")
for i in scheduleList:
    context[i[0:12].strip()] = i[12:].strip()
doc.render(context)
doc.save("Расписание.docx")

stop = input("Итоговый файл готов, [Нажмите Enter для завершения]")