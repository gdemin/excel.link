context("workbooks")
xl.workbook.add()
xlrc[a1] <- iris
xl.workbook.save("iris.xlsx")
xl.workbook.add()
xlrc[a1] <- cars
xl.workbook.save("cars.xlsx")
xl.workbook.activate("iris")
xl.workbook.close("cars")
xl.workbook.open("cars.xlsx")
for (wb in xl.workbooks()) xl.workbook.close(wb)
unlink("iris.xlsx")
unlink("cars.xlsx")