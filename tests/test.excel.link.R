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


###### xl.connect.table #####
data(iris)
rownames(iris)=as.character(rownames(iris))
iris$Species=as.character(iris$Species)
xl.workbook.add()


xlrc[a1]=iris
xl.iris=xl.connect.table("a1",row.names=TRUE,col.names=TRUE)
identical(xl.iris[],iris)

iris=iris[order(iris$Sepal.Length),]
sort(xl.iris,column="Sepal.Length")
identical(xl.iris[],iris)

sort(xl.iris,column="rownames")
iris=iris[order(rownames(iris)),]
identical(xl.iris[],iris)

identical(xl.iris[,1:3],iris[,1:3])
identical(xl.iris[,3],iris[,3])
identical(xl.iris[26,1:3],iris[26,1:3])
identical(xl.iris[-26,1:3],iris[-26,1:3])
identical(xl.iris[50,],iris[50,])
identical(xl.iris$Species,iris$Species)
identical(xl.iris[,'Species',drop=FALSE],iris[,'Species',drop=FALSE])
identical(xl.iris[c(TRUE,FALSE),'Sepal.Length'],iris[c(TRUE,FALSE),'Sepal.Length'])



xl.iris[,'group']=xl.iris$Sepal.Length>mean(xl.iris$Sepal.Length)
iris[,'group']=iris$Sepal.Length>mean(iris$Sepal.Length)
identical(xl.iris[],iris)

xl.iris$temp=c('aa','bb')
iris$temp=c('aa','bb')
identical(xl.iris[],iris)

xl.iris[,"temp"]=NULL
iris[,"temp"]=NULL
identical(xl.iris[],iris)

xl.iris[xl.iris$Sepal.Length>6,"Sepal.Length"]=xl.iris$Petal.Length[xl.iris$Sepal.Length>6]
iris[iris$Sepal.Length>6,"Sepal.Length"]=iris$Petal.Length[iris$Sepal.Length>6]
identical(xl.iris[],iris)
xl.iris[xl.iris$Sepal.Length>6,"dummy"]=xl.iris$Petal.Length[xl.iris$Sepal.Length>6]
xl.iris[xl.iris$Species=="setosa",c("a","b")]=xl.iris[xl.iris$Species=="setosa",c('Species','Petal.Length')]
iris[iris$Sepal.Length>6,"dummy"]=iris$Petal.Length[iris$Sepal.Length>6]
iris[iris$Species=="setosa","a"]=iris[iris$Species=="setosa",'Species']
iris[iris$Species=="setosa","b"]=iris[iris$Species=="setosa",'Petal.Length']
identical(xl.iris[],iris)


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



### dimnames ######
data(iris)
xl.workbook.add()
xl[d4]=iris
test1=xl.connect.table("d4",row.names=FALSE,col.names=FALSE)
has.colnames(test1)==FALSE 
has.rownames(test1)==FALSE 
dimnames(test1)
rownames(test1)
colnames(test1)

xl.sheet.add()
xlrc[d4]=iris
test2=xl.connect.table("d4",row.names=TRUE,col.names=TRUE)
has.colnames(test2) # TRUE
has.rownames(test2) # TRUE
dimnames(test2)
rownames(test2)
colnames(test2)

############# xl ##################
data(iris)
rownames(iris)=as.character(rownames(iris))
iris$Species=as.character(iris$Species)
xl.workbook.add()
xlrc$a1 <- iris
xl.iris <- xl.current.region("a1",row.names=TRUE,col.names=TRUE)
identical(xl.iris,iris)

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


####### na ################################
set.seed(1)
xl.workbook.add()
aaa=matrix(rnorm(100000),ncol=100)
xl[a1]=aaa
aaa[aaa<0]=NA
xl[a1]=aaa
test=as.matrix(xl[a1:cv1000])
dimnames(test)=NULL
all.equal(aaa,test)
aaa=matrix(sample(c(TRUE,FALSE,NA),100000,replace=TRUE),ncol=100)
xl[a1]=aaa  #### долго 
test=as.matrix(xl[a1:cv1000])
dimnames(test)=NULL
all.equal(aaa,test)

####### na, na='na' ################################
set.seed(1)
xl.workbook.add()
aaa=matrix(rnorm(100000),ncol=100)
aaa[aaa<0]=NA
xl[a1, na='na']=aaa
test=as.matrix(xl[a1:cv1000, na='na'])
dimnames(test)=NULL
all.equal(aaa,test)
aaa=matrix(sample(c(TRUE,FALSE,NA),100000,replace=TRUE),ncol=100)
xl[a1, na='na']=aaa
test=as.matrix(xl[a1:cv1000, na='na'])
dimnames(test)=NULL
all.equal(aaa,test)



##### multi-column element of data.frame ######
xl.workbook.add()
test=data.frame(a=letters[1:3],b=I(matrix(1:9,3)),d=LETTERS[1:3])
str(test)
xlrc[a1]=test 
test2=data.frame(a=letters[1:3],b=(matrix(1:9,3)),d=LETTERS[1:3])
all(test2==xlrc[a1:f4])


###### files ########################
setwd("c:/temp")
data(iris)
rownames(iris)=as.character(rownames(iris))
iris$Species=as.character(iris$Species)
filename=paste0(tempfile(),".xlsx")
xl.save.file(iris,filename)

#####
xl.iris=xl.read.file(filename)
identical(iris,xl.iris) # Shoud be TRUE
unlink(filename)
####

xl.save.file(iris,"iris.xlsx")
xl.iris=xl.read.file("iris.xlsx")
all(iris==xl.iris) # Shoud be TRUE
unlink("iris.xlsx")
################

xl.save.file(list(t(seq_along(iris)),iris),filename)
xl.iris=xl.read.file(filename,top.left.cell="A2")
str(xl.iris)
identical(iris,xl.iris)
unlink(filename)
######
xl.save.file(list(t(seq_along(iris)),iris),filename,top.left.cell="d24")
xl.iris=xl.read.file(filename,top.left.cell="d25")
str(xl.iris)
identical(iris,xl.iris) 
unlink(filename)
######
xl.save.file(list(t(seq_along(iris)),iris),filename,xl.sheet="iris",top.left.cell="d24")
xl.iris=xl.read.file(filename,xl.sheet="iris",top.left.cell="d25")
str(xl.iris)
identical(iris,xl.iris) 


xl.iris=xl.read.file(filename,xl.sheet=1,top.left.cell="d25")
str(xl.iris)
identical(iris,xl.iris) 
unlink(filename) 

#######
xl.save.file(iris,filename)
xl.iris=xl.read.file(filename,header=FALSE,top.left.cell="b2")
str(xl.iris)
all(iris==xl.iris) # Shoud be TRUE
unlink(filename)

#######
xl.save.file(list(t(seq_along(iris)),iris),filename,xl.sheet="iris",top.left.cell="d24")
xl.iris=xl.read.file(filename,header=FALSE,xl.sheet="iris",top.left.cell="e26")
str(xl.iris)
all(iris==xl.iris) 
unlink(filename)
 ######
 
dists <- dist(iris[,1:4])
clusters <- hclust(dists,method="ward")
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


######
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