context("workbooks")
data(iris)
data(cars)
workbooks = xl.workbooks()
xl.workbook.add()
xlrc[a1] = iris
xl.workbook.save("iris.xlsx")
xl.workbook.add()
xlrc[a1] = cars
xl.workbook.save("cars.xlsx")
for (wb in workbooks) xl.workbook.close(wb)
expect_identical(xl.workbooks(), c("iris.xlsx","cars.xlsx"))
xl.workbook.activate("iris")
xl.workbook.close("cars")
xl.workbook.close()

xls = xl.get.excel()
xls$quit()

xl.workbook.open("cars.xlsx")

books = xl.workbooks()
expect_equal(length(books), 1)
rownames(cars) = as.character(rownames(cars))
expect_identical(cars,xl.current.region("a1",row.names=TRUE,col.names=TRUE))
xl.workbook.open("iris.xlsx")
rownames(iris) = as.character(rownames(iris))
iris$Species = as.character(iris$Species)
expect_identical(all(iris==xl.current.region("a1",row.names=TRUE,col.names=TRUE)),TRUE)
for (wb in xl.workbooks()) xl.workbook.close(wb)
unlink("iris.xlsx")
unlink("cars.xlsx")


####

xls = xl.get.excel()
xls$quit()
xl.workbook.add()
books = xl.workbooks()
expect_equal(length(books), 1)
xls = xl.get.excel()
xls$quit()
