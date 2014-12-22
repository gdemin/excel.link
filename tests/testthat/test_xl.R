context("xl (long test)")

data(iris)
rownames(iris)=as.character(rownames(iris))
iris$Species=as.character(iris$Species)
xl.workbook.add()
xlrc$a1 <- iris
xl.iris <- xl.current.region("a1",row.names=TRUE,col.names=TRUE)
expect_identical(xl.iris,iris)
xl.workbook.close()

## NA

set.seed(1)
xl.workbook.add()
aaa=matrix(rnorm(100000),ncol=100)
xl[a1]=aaa
aaa[aaa<0]=NA
xl[a1]=aaa
test=as.matrix(xl[a1:cv1000])
dimnames(test)=NULL
expect_equal(aaa,test)
aaa=matrix(sample(c(TRUE,FALSE,NA),100000,replace=TRUE),ncol=100)
xl[a1]=aaa  #### долго 
test=as.matrix(xl[a1:cv1000])
dimnames(test)=NULL
expect_equal(aaa,test)
xl.workbook.close()

#### NA na='na'

set.seed(1)
xl.workbook.add()
aaa=matrix(rnorm(100000),ncol=100)
aaa[aaa<0]=NA
xl[a1, na='na']=aaa
test=as.matrix(xl[a1:cv1000, na='na'])
dimnames(test)=NULL
expect_equal(aaa,test)
aaa=matrix(sample(c(TRUE,FALSE,NA),100000,replace=TRUE),ncol=100)
xl[a1, na='na']=aaa
test=as.matrix(xl[a1:cv1000, na='na'])
dimnames(test)=NULL
expect_equal(aaa,test)

xl.workbook.close()

##### multi-column element of data.frame ######
xl.workbook.add()
test=data.frame(a=letters[1:3],b=I(matrix(1:9,3)),d=LETTERS[1:3])
xlrc[a1]=test 
test2=data.frame(a=letters[1:3],b=(matrix(1:9,3)),d=LETTERS[1:3],stringsAsFactors = FALSE)
expect_equal(all(test2==xlrc[a1:f4]),TRUE)
expect_identical(colnames(test2),colnames(xlrc[a1:f4]))

xl.workbook.close()