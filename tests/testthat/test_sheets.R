context("sheets")
xl.workbook.add()
sheets = xl.sheets()
xl.sheet.add("Second")
xl.sheet.add("First",before="Second")
for (sheet in sheets) xl.sheet.delete(sheet) # only 'First' and 'Second' exist in workbook now
expect_identical(xl.sheets(), c("First","Second"))
xl.sheet.activate("Second") #last sheet activated 
xl.sheet.duplicate()
expect_identical(xl.sheets(), c("First", "Second", "Second (2)"))
new_sheet_name = xl.sheet.duplicate(before = "First")
expect_identical(new_sheet_name, "Second (3)")
expect_identical(xl.sheets(), c("Second (3)", "First", "Second", "Second (2)"))
xl.workbook.close()
