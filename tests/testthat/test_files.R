
context("files")

# setwd("c:/temp")
data(iris)
rownames(iris)=as.character(rownames(iris))
iris$Species=as.character(iris$Species)
filename=paste0(tempfile(),".xlsx")
xl.save.file(iris,filename)

#####
xl.iris=xl.read.file(filename)
expect_identical(iris,xl.iris) 
unlink(filename)
####

xl.save.file(iris,"iris.xlsx")
xl.iris=xl.read.file("iris.xlsx")
expect_equal(iris,xl.iris) 
unlink("iris.xlsx")
################

xl.save.file(list(t(seq_along(iris)),iris),filename)
xl.iris=xl.read.file(filename,top.left.cell="A2")
expect_identical(iris,xl.iris)
unlink(filename)
######
xl.save.file(list(t(seq_along(iris)),iris),filename,top.left.cell="d24")
xl.iris=xl.read.file(filename,top.left.cell="d25")
expect_identical(iris,xl.iris) 
unlink(filename)
######
xl.save.file(list(t(seq_along(iris)),iris),filename,xl.sheet="iris",top.left.cell="d24")
xl.iris=xl.read.file(filename,xl.sheet="iris",top.left.cell="d25")
expect_identical(iris,xl.iris) 


xl.iris=xl.read.file(filename,xl.sheet=1,top.left.cell="d25")
expect_identical(iris,xl.iris) 
unlink(filename) 

#######
xl.save.file(iris,filename)
xl.iris=xl.read.file(filename,header=FALSE,top.left.cell="b2")
expect_equal(all(iris==xl.iris),TRUE) 
unlink(filename)

#######
xl.save.file(list(t(seq_along(iris)),iris),filename,xl.sheet="iris",top.left.cell="d24")
xl.iris=xl.read.file(filename,header=FALSE,xl.sheet="iris",top.left.cell="e26")
expect_equal(all(iris==xl.iris),TRUE) 
unlink(filename)

######################
context("xl.read.file/xl.save.file with passwords")
filename = "iris.xlsx"
data(iris)
test_iris = iris
test_iris$Species = as.character(test_iris$Species)
rownames(test_iris) = as.character(1:150)

xl.save.file(test_iris,filename, password = "read_password")
new_iris = xl.read.file(filename, password = "read_password")
expect_identical(test_iris, new_iris)

xl.save.file(test_iris,filename, password = "read_password", write.res.password = "edit_password")
new_iris = xl.read.file(filename, password = "read_password", write.res.password = "edit_password")
expect_identical(test_iris, new_iris)

xl.save.file(test_iris,filename, write.res.password = "edit_password")
new_iris = xl.read.file(filename, write.res.password = "edit_password")
expect_identical(test_iris, new_iris)

unlink("iris.xlsx")

###################################
context("xl.read.file with hidden sheet")

data(iris)
data(cars)
workbooks = xl.workbooks()
xl.workbook.add()
xl.sheet.add("iris")
xlrc[a1] = iris
xl.sheet.add("cars")
xlrc[a1] = cars
xl.sheet.hide("iris")
expect_false(xl.sheet.visible("iris"))
expect_error(xl.sheet.activate("iris"))
xl.workbook.save("hidden.xlsx")
xl.workbook.close()

new_cars = xl.read.file("hidden.xlsx", row.names = TRUE, col.names = TRUE, xl.sheet = "cars")
new_iris = xl.read.file("hidden.xlsx", row.names = TRUE, col.names = TRUE, xl.sheet = "iris")


rownames(cars) = as.character(rownames(cars))
rownames(iris) = as.character(rownames(iris))
iris$Species = as.character(iris$Species)

expect_identical(cars, new_cars)
expect_identical(iris, new_iris)

unlink("hidden.xlsx")

