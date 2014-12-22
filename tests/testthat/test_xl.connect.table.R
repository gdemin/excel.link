context("xl.connect.table")
data(iris)
rownames(iris)=as.character(rownames(iris))
iris$Species=as.character(iris$Species)
xl.workbook.add()


xlrc[a1]=iris
xl.iris=xl.connect.table("a1",row.names=TRUE,col.names=TRUE)
expect_equal(xl.iris[],iris)

iris=iris[order(iris$Sepal.Length),]
sort(xl.iris,column="Sepal.Length")
expect_equal(xl.iris[],iris)

sort(xl.iris,column="rownames")
iris=iris[order(rownames(iris)),]
expect_equal(xl.iris[],iris)

expect_equal(xl.iris[,1:3],iris[,1:3])
expect_equal(xl.iris[,3],iris[,3])
expect_equal(xl.iris[26,1:3],iris[26,1:3])
expect_equal(xl.iris[-26,1:3],iris[-26,1:3])
expect_equal(xl.iris[50,],iris[50,])
expect_equal(xl.iris$Species,iris$Species)
expect_equal(xl.iris[,'Species',drop=FALSE],iris[,'Species',drop=FALSE])
expect_equal(xl.iris[c(TRUE,FALSE),'Sepal.Length'],iris[c(TRUE,FALSE),'Sepal.Length'])



xl.iris[,'group']=xl.iris$Sepal.Length>mean(xl.iris$Sepal.Length)
iris[,'group']=iris$Sepal.Length>mean(iris$Sepal.Length)
expect_equal(xl.iris[],iris)

xl.iris$temp=c('aa','bb')
iris$temp=c('aa','bb')
expect_equal(xl.iris[],iris)

xl.iris[,"temp"]=NULL
iris[,"temp"]=NULL
expect_equal(xl.iris[],iris)

xl.iris[xl.iris$Sepal.Length>6,"Sepal.Length"]=xl.iris$Petal.Length[xl.iris$Sepal.Length>6]
iris[iris$Sepal.Length>6,"Sepal.Length"]=iris$Petal.Length[iris$Sepal.Length>6]
expect_equal(xl.iris[],iris)
xl.iris[xl.iris$Sepal.Length>6,"dummy"]=xl.iris$Petal.Length[xl.iris$Sepal.Length>6]
xl.iris[xl.iris$Species=="setosa",c("a","b")]=xl.iris[xl.iris$Species=="setosa",c('Species','Petal.Length')]
iris[iris$Sepal.Length>6,"dummy"]=iris$Petal.Length[iris$Sepal.Length>6]
iris[iris$Species=="setosa","a"]=iris[iris$Species=="setosa",'Species']
iris[iris$Species=="setosa","b"]=iris[iris$Species=="setosa",'Petal.Length']
expect_equal(xl.iris[],iris)

xl.workbook.close()


##### sample session #####
library(excel.link)
xl.workbook.add()
xl.sheet.add("Iris dataset",before=1)
xlrc[a1] <- iris
xl.iris <- xl.connect.table("a1",row.names=TRUE,col.names=TRUE)
dists <- dist(xl.iris[,1:4])
clusters <- hclust(dists,method="ward.D")
xl.iris$clusters <- cutree(clusters,3)
plot(clusters)
pl.clus <- current.graphics()
cross <- table(xl.iris$Species,xl.iris$clusters)
plot(cross)
pl.cross <- current.graphics()
xl.sheet.add("Results",before=2)
xlrc$a1 <- list("Crosstabulation",cross,pl.cross,"Dendrogram",pl.clus)
