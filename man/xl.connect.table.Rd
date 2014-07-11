\name{xl.connect.table}
\alias{xl.connect.table}
\alias{[.excel.range}
\alias{$.excel.range}
\alias{[<-.excel.range}
\alias{$<-.excel.range}
\alias{sort.excel.range}

\title{
Interactive work with data in Microsoft Excel 
}
\description{
'xl.connect.table' returns object of 'excel.range' class which represent data on Excel sheet. This object can be treated similar to data.frame. So you can assign values, delete columns/rows and so on. For more information see details and examples. 

'sort' sorts Excel range by single column (multiple columns currently not supported).
}
\usage{
xl.connect.table(str.rng="A1",row.names=TRUE,col.names=TRUE,na="")

\method{[}{excel.range}(x, i, j, drop = if (missing(i)) TRUE else !missing(j) && (length(j) == 1))
\method{$}{excel.range}(x,value)

\method{[}{excel.range}(x,i,j) <- value 
\method{$}{excel.range}(x,j) <- value

\method{sort}{excel.range}(x,decreasing=FALSE,column,\dots)
}
\arguments{
\item{x}{An object of 'excel.range' class}
\item{str.rng}{string which represents Excel range}
\item{i}{rows to extract or replace.  These are 'numeric' or 'character' or 'logical'}
\item{j}{columns to extract or replace.  These are 'numeric' or 'character' or 'logical'}
\item{drop}{logical.  If TRUE the result is coerced to the lowest possible dimension}
\item{row.names}{a logical value indicating whether the Excel range contains the rows names as its first column}
\item{col.names}{a logical value indicating whether the Excel range contains the columns names as its first row}
\item{na}{character. NA representaion in Excel. By default it is empty string}
\item{value}{a suitable replacement value}
\item{decreasing}{logical.  Should the sort be increasing or decreasing?}
\item{column}{number or string. Column by which sort}
\item{\dots}{arguments to be passed to or from methods or (for the default methods and objects without a class)}
}
%- maybe also 'usage' for other objects documented here.
\details{
Connected range is 'current region', e. g. selection wich can be obtained by pressing Ctrl+Shift+* when selected 'str.rng' (or top-left cell of this range is active).

Indices are 'numeric' or 'character' vectors or empty (missing).  Numeric values are coerced to integer as by 'as.integer' (and hence truncated towards zero).  Character vectors will be matched to the 'colnames' of the object (or Excel column names if has.colnames==FALSE). For extraction form if column name doesn't exist error will be generated. For replacement form new column will be created.

'i', 'j' can be logical vectors, indicating elements/slices to select.  Such vectors are recycled if necessary to match the corresponding extent. 'i', 'j' can also be negative integers, indicating elements/slices to leave out of the selection.

Matrix indexing is not supported. 

There is special value for 'column' argument in 'sort' - 'rownames'. In this case 'x' will be sorted by row names if it has it.

%%  ~~ If necessary, more details than the description above ~~
}
\value{
For 'xl.connect.table' object of 'excel.range' class.

For '[' a data frame or a single column (the latter only when dimensions have been dropped).  If non-existent column selected error will be generated.

For '$', a column.

For '[<-', '[[<-' and '$<-', object of 'excel.range' class.

'sort' invisibly return NULL.

%%  ~Describe the value returned
%%  If it is a LIST, use
%%  \item{comp1 }{Description of 'comp1'}
%%  \item{comp2 }{Description of 'comp2'}
%% ...
}

\author{
Gregory Demin <excel.link.feedback@gmail.com>
}


\examples{

\dontrun{
### session example ###
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

### completly senseless actions #####
data(iris)
rownames(iris) <- as.character(rownames(iris))
iris$Species <- as.character(iris$Species)
xl.workbook.add()


xlrc[a1] <- iris
xl.iris <- xl.connect.table("a1",row.names=TRUE,col.names=TRUE)
identical(xl.iris[],iris)

iris <- iris[order(iris$Sepal.Length),]
sort(xl.iris,column="Sepal.Length")
identical(xl.iris[],iris)

sort(xl.iris,column="rownames")
iris <- iris[order(rownames(iris)),]
identical(xl.iris[],iris)

identical(xl.iris[,1:3],iris[,1:3])
identical(xl.iris[,3],iris[,3])
identical(xl.iris[26,1:3],iris[26,1:3])
identical(xl.iris[-26,1:3],iris[-26,1:3])
identical(xl.iris[50,],iris[50,])
identical(xl.iris$Species,iris$Species)
identical(xl.iris[,'Species',drop=FALSE],iris[,'Species',drop=FALSE])
identical(xl.iris[c(TRUE,FALSE),'Sepal.Length'],iris[c(TRUE,FALSE),'Sepal.Length'])

xl.iris[,'group'] <- xl.iris$Sepal.Length>mean(xl.iris$Sepal.Length)
iris[,'group'] <- iris$Sepal.Length>mean(iris$Sepal.Length)
identical(xl.iris[],iris)

xl.iris$temp <- c('aa','bb')
iris$temp <- c('aa','bb')
identical(xl.iris[],iris)

xl.iris[,"temp"] <- NULL
iris[,"temp"] <- NULL
identical(xl.iris[],iris)


}
  
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{ IO }
\keyword{ connection }