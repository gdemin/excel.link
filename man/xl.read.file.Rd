\name{xl.read.file, xl.save.file}
\alias{xl.read.file}
\alias{xl.save.file}
\title{
Functions for saving and reading data to/from Excel file.
}
\description{
Functions for saving and reading data to/from Excel file.
}
\usage{
xl.read.file(filename, header = TRUE, row.names = NULL, col.names = NULL, 
             xl.sheet = NULL, top.left.cell = "A1", na = "", excel.visible = FALSE)
			 
xl.save.file(r.obj, filename, row.names = TRUE, col.names = TRUE, 
             xl.sheet = NULL, top.left.cell = "A1", na = "", excel.visible = FALSE)
}
\arguments{
\item{filename}{a character string naming a file}
\item{r.obj}{R object}
\item{header}{a logical value indicating whether the file contains the names of the variables as its first line. If TRUE and top-left corner is empty cell, first column is considered as row names. Ignored if row.names or col.names is not NULL.}
\item{row.names}{a logical value indicating whether the row names of r.obj are to be read/saved along with r.obj}
\item{col.names}{a logical value indicating whether the column names of r.obj are to be read/saved along with r.obj}
\item{xl.sheet}{character. Name of Excel sheet where data is located/will be saved. By default it is NULL and data will be read/saved from/to active sheet.}
\item{top.left.cell}{character. Top-left corner of data in Excel sheet. By default is 'A1'.}
\item{na}{character. NA representation in Excel. By default it is empty string}
\item{excel.visible}{a logical value indicating will Excel visible during this operations. FALSE by default.}
}
%- maybe also 'usage' for other objects documented here.
\details{
xl.read.file reads only rectangular data set. It is highly recommended to have all column names and ids in data set. Orphaned rows/columns located apart from the main data will be ignored.  xl.save.file can save all objects for which xl.write method exists - see example.

%%  ~~ If necessary, more details than the description above ~~
}
\value{
xl.read.file always return data.frame.
xl.save.file return NULL.
%%  ~Describe the value returned
%%  If it is a LIST, use
%%  \item{comp1 }{Description of 'comp1'}
%%  \item{comp2 }{Description of 'comp2'}
%% ...
}

\author{
Gregory Demin <excel.link.feedback@gmail.com>
}

\seealso{xl.write, xl.workbook.save, xl.workbook.open, current.graphics}

\examples{

\dontrun{
data(iris)
xl.save.file(iris,"iris.xlsx")
xl.iris = xl.read.file("iris.xlsx")
all(iris == xl.iris) # Shoud be TRUE
unlink("iris.xlsx")

# Save to file list with different data types 
dists = dist(iris[,1:4])
clusters = hclust(dists,method="ward")
iris$clusters = cutree(clusters,3)
png("1.png")
plot(clusters)
dev.off()
pl.clus = current.graphics(filename="1.png")
cross = table(iris$Species,iris$clusters)
png("2.png")
plot(cross)
dev.off()
pl.cross = current.graphics(filename="2.png")
output = list("Iris",pl.clus,cross,pl.cross,"Data:","",iris)
xl.save.file(output,"output.xls")
xl.workbook.open("output.xls")
# xl.workbook.close() # close workbook
# unlink("output.xls") # delete file

}
  
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
% \keyword{ ~xl }
\keyword{ IO }
\keyword{ connection }