context("current.region")

data(iris)
rownames(iris)=as.character(rownames(iris))
iris$Species=as.character(iris$Species)
xl.workbook.add()
xlrc$a1 <- iris
xl.iris <- crrc$a1
expect_identical(xl.iris,iris)
xl.workbook.close()

##Clear range


xl.workbook.add()
aaa=matrix(1:16,ncol=4)
cr[a1]=aaa
bbb = matrix(1:9,ncol=3)
cr[a1] = bbb
expect_equal(all(bbb==cr$a1),TRUE)
new_bbb = data.frame(rbind(bbb,NA),NA)
colnames(new_bbb) = letters[1:4]
rownames(new_bbb) = NULL
expect_equal(xl[a1:d4],new_bbb)
xl.workbook.close()

