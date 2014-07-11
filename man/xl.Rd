\name{xl}
\alias{xl}
\alias{xlc}
\alias{xlr}
\alias{xlrc}
\alias{xl.selection}
\alias{xl.current.region}
\alias{[.xl}
\alias{[<-.xl}
\alias{[[.xl}
\alias{[[<-.xl}
\alias{$.xl}
\alias{$<-.xl}

\title{
Object for working with Microsoft Excel data
}
\description{
Object aimed to work with data in Microsoft Excel ranges. 
}
\usage{
xl.selection(drop=TRUE,na="",row.names=FALSE,col.names=FALSE)

xl.current.region(str.rng,drop=TRUE,na="",row.names=FALSE,col.names=FALSE)

\method{[}{xl}(x,str.rng,drop=!(has.rownames(x) | has.colnames(x)),na="")
\method{[}{xl}(x,str.rng,na="") <- value

\method{[[}{xl}(x,str.rng,drop=!(has.rownames(x) | has.colnames(x)),na="")
\method{[[}{xl}(x,str.rng,na="") <- value

\method{$}{xl}(x,str.rng)
\method{$}{xl}(x,str.rng) <- value
}
\arguments{
\item{xl}{object representing Microsoft Excel application}
\item{x}{xl, xlc, xlr or xlrc. xl - read/write without column and row names, "r" - with rownames, "c" - with colnames}
\item{str.rng}{string which represents Excel range. For single bracket operations it can be without quotes in almost all cases}
\item{drop}{logical.  If TRUE the result is coerced to the lowest possible dimension.  The default is to drop if there are no columns and rows names}
\item{row.names}{a logical value indicating whether the Excel range contains the row names as its first column}
\item{col.names}{a logical value indicating whether the Excel range contains the column names as its first row}
\item{na}{character. NA representation in Excel. By default it is empty string}
\item{value}{a suitable replacement value. It recycled to fill excel range only if value is object of length 1. In other cases size of excel range is ignored}
}
%- maybe also 'usage' for other objects documented here.
\details{
For convenient interactive usage arguments can be given without quotes in most cases (e. g. xl[a1]=5 or xl[u2:u85]="Hi" or xl[MyNamedRange]=42, but xl["Sheet1!A1"]=42). When it used in functions or you need to use variable as argument it is recommended apply double brackets notation:(e. g. xl[["a1"]]=5 or xl[["u2:u85"]]="Hi" or xl[["MyNamedRange"]]=42).

Difference between 'xl', 'xlc', 'xlrc' and 'xlr' is 'xl' ignore row and column names, 'xlc' suppose read and write to Excel with column names, 'xlrc' - with column and row names and so on. There is argument 'drop' which is TRUE by default for 'xl' and FALSE by default for other options.

'xl.selection' returns data.frame with data from current selection in Excel.

'xl.current.region' returns data.frame with data from current region (range which can be selected by pressing Ctrl+Shift+*) in Excel.
%%  ~~ If necessary, more details than the description above ~~
}
\value{
Return appropriate dataset from Excel. Excel datetime type currently not supported.
}

\author{
Gregory Demin <excel.link.feedback@gmail.com>
}

\examples{

\dontrun{
data(iris)
rownames(iris) <- as.character(rownames(iris))
iris$Species <- as.character(iris$Species)
xl.workbook.add()
xlrc$a1 <- iris
xl.iris <- xl.current.region("a1",row.names=TRUE,col.names=TRUE)
identical(xl.iris,iris)

xl.sheet.add("Datasets examples")
data.sets <- list("Iris dataset",iris,"Cars dataset",cars,"Titanic dataset",as.data.frame(Titanic))
xlrc[a1] <- data.sets
}
  
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
% \keyword{ ~xl }
\keyword{ IO }
\keyword{ connection }