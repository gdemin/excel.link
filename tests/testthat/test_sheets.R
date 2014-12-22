context("sheets")
xl.workbook.add()
sheets <- xl.sheets()
xl.sheet.add("Second")
xl.sheet.add("First",before="Second")
for (sheet in sheets) xl.sheet.delete(sheet) # only 'First' and 'Second' exist in workbook now
xl.sheet.activate("Second") #last sheet activated 
xl.workbook.close()