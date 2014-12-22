context("files")

setwd("c:/temp")
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