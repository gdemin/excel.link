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
xl[z1] = 1
xl.workbook.activate("cars")
expect_identical(colnames(crrc[a1]), c("speed", "dist"))
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

#################
context("xl.workbook.open with passwords")

data(iris)
test_iris = iris
test_iris$Species = as.character(test_iris$Species)
rownames(test_iris) = as.character(1:150)
xl.workbook.add()
xlrc[a1] = test_iris
xl.workbook.save("iris.xlsx", password = "read_password")
xl.workbook.close()
xl.workbook.open("iris.xlsx", password = "read_password")
new_iris = crrc[a1]
expect_identical(test_iris, new_iris)
xl.workbook.save("iris.xlsx", password = "read_password", write.res.password = "edit_password")
xl.workbook.close()
xl.workbook.open("iris.xlsx", password = "read_password", write.res.password = "edit_password")
new_iris = crrc[a1]
expect_identical(test_iris, new_iris)
xl.workbook.save("iris.xlsx", password = "", write.res.password = "edit_password")
xl.workbook.close()
xl.workbook.open("iris.xlsx", write.res.password = "edit_password")
new_iris = crrc[a1]
expect_identical(test_iris, new_iris)
xl.workbook.save("iris.xlsx", write.res.password = "")
xl.workbook.close()
xl.workbook.open("iris.xlsx")
new_iris = crrc[a1]
expect_identical(test_iris, new_iris)
xl.workbook.close()
unlink("iris.xlsx")
#######################
context("xl.workbook.open with passwords xlsb")

data(iris)
test_iris = iris
test_iris$Species = as.character(test_iris$Species)
rownames(test_iris) = as.character(1:150)
xl.workbook.add()
xlrc[a1] = test_iris
xl.workbook.save("iris.xlsb", password = "read_password", file.format = xl.constants$xlExcel12)
xl.workbook.close()
xl.workbook.open("iris.xlsb", password = "read_password")
new_iris = crrc[a1]
expect_identical(test_iris, new_iris)
xl.workbook.save("iris.xlsb", password = "read_password", write.res.password = "edit_password", file.format = xl.constants$xlExcel12)
xl.workbook.close()
xl.workbook.open("iris.xlsb", password = "read_password", write.res.password = "edit_password")
new_iris = crrc[a1]
expect_identical(test_iris, new_iris)
xl.workbook.save("iris.xlsb", password = "", write.res.password = "edit_password", file.format = xl.constants$xlExcel12)
xl.workbook.close()
xl.workbook.open("iris.xlsb", write.res.password = "edit_password")
new_iris = crrc[a1]
expect_identical(test_iris, new_iris)
xl.workbook.save("iris.xlsb", write.res.password = "", file.format = xl.constants$xlExcel12)
xl.workbook.close()
xl.workbook.open("iris.xlsb")
new_iris = crrc[a1]
expect_identical(test_iris, new_iris)
xl.workbook.close()
unlink("iris.xlsb")