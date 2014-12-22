# not for automatic testing
help(package=excel.link)
library(excel.link)
data(iris)
library(RDCOMClient)
##### sample session #####
library(excel.link)
xl.workbook.add()
xl.sheet.add("Iris dataset",before=1)
xlrc[a1] <- iris
xl.iris <- xl.connect.table("a1",row.names=TRUE,col.names=TRUE)
dists <- dist(xl.iris[,1:4])
clusters <- hclust(dists,method="ward")
xl.iris$clusters <- cutree(clusters,3)
plot(clusters)
pl.clus <- current.graphics()
cross <- table(xl.iris$Species,xl.iris$clusters)
plot(cross)
pl.cross <- current.graphics()
xl.sheet.add("Results",before=2)
xlrc$a1 <- list("Crosstabulation",cross,pl.cross,"Dendrogram",pl.clus)





#### workbooks ##########
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

##### sheets #######
xl.workbook.add()
sheets <- xl.sheets()
xl.sheet.add("Second")
xl.sheet.add("First",before="Second")
for (sheet in sheets) xl.sheet.delete(sheet) # only 'First' and 'Second' exist in workbook now
xl.sheet.activate("Second") #last sheet activated 
xl.workbook.close()




############# xl ##################

xl.sheet.add("Datasets examples")
data.sets <- list("Iris dataset",iris,"Cars dataset",cars,"Titanic dataset",as.data.frame(Titanic))
xlrc[a1] <- data.sets

#### current.graphics #############
xl.workbook.add()
plot(sin)
xl[a1]=current.graphics()
plot(cos)
cos.plot=current.graphics()
xl.sheet.add()
xl[a1]=list("Cosine plotting",cos.plot,"End of cosine plotting")





###### files ########################
setwd("c:/temp")

 ######
 
dists <- dist(iris[,1:4])
clusters <- hclust(dists,method="ward.D")
iris$clusters <- cutree(clusters,3)
plot(clusters)
pl.clus <- current.graphics()
cross <- table(iris$Species,iris$clusters)
plot(cross)
pl.cross <- current.graphics()
output=list("Iris",pl.clus,cross,pl.cross,"Data:","",iris)
xl.save.file(output,filename)
xl.workbook.open(filename)
xl.workbook.close()
unlink(filename)

######
filename=paste0(tempfile(),".xlsx")
dists <- dist(iris[,1:4])
clusters <- hclust(dists,method="ward")
iris$clusters <- cutree(clusters,3)
png("1.png")
plot(clusters)
dev.off()
pl.clus <- current.graphics(filename="1.png")
cross <- table(iris$Species,iris$clusters)
png("2.png")
plot(cross)
dev.off()
pl.cross <- current.graphics(filename="2.png")
output=list("Iris",pl.clus,cross,pl.cross,"Data:","",iris)
xl.save.file(output,filename)
xl.workbook.open(filename)
xl.workbook.close()
unlink(filename)

######
dists <- dist(iris[,1:4])
clusters <- hclust(dists,method="ward")
iris$clusters <- cutree(clusters,3)
png("1.png")
plot(clusters)
dev.off()
pl.clus <- current.graphics(filename="1.png")
cross <- table(iris$Species,iris$clusters)
png("2.png")
plot(cross)
dev.off()
pl.cross <- current.graphics(filename="2.png")
output=list("Iris",pl.clus,cross,pl.cross,"Data:","",iris)
xl.save.file(output,"output.xls")
xl.workbook.open("output.xls")
# xl.workbook.close() # close workbook
# unlink("output.xls") # delete file


###### move to automatic tests in the future ######## 
setwd("c:/temp")
xl.workbook.add()
xlrc[a1] = iris
xl[a1] = 'dfsdf'
xl.workbook.save("iris.xlsx")
xl.workbook.close() 
# debug(xl.read.file)
xl.iris=xl.read.file("iris.xlsx")
str(xl.iris)
xl.iris=xl.read.file("iris.xlsx",row.names=TRUE,col.names=TRUE)
str(xl.iris)
xl.workbook.open("iris.xlsx")
xl[a1]="        "
xl.workbook.save("iris.xlsx")
xl.workbook.close() 
xl.iris=xl.read.file("iris.xlsx")
str(xl.iris)
unlink("iris.xlsx")

###################

png("1.png")
plot(sin)
dev.off()
sin.plot = current.graphics(filename = "1.png")
png("2.png")
plot(cos)
dev.off()
cos.plot = current.graphics(filename = "2.png")
output = list("Cosine plotting",cos.plot,"Sine plotting",sin.plot)
xl.workbook.add()
xl[a1] = output