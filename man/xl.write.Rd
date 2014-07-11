\name{xl.write}
\alias{xl.write}
\alias{xl.write.default}
\alias{xl.write.data.frame}
\alias{xl.write.matrix}
\alias{xl.write.factor}
\alias{xl.write.list}
\alias{xl.write.table}
\alias{xl.write.current.graphics}

\title{
Methods for writing data to Excel sheet
}
\description{
Methods for writing data to Excel sheet
}
\usage{
xl.write(r.obj,xl.rng,na="",\dots)
\method{xl.write}{default}(r.obj,xl.rng,na="",row.names=TRUE,\dots)
\method{xl.write}{data.frame}(r.obj,xl.rng,na="",row.names=TRUE,col.names=TRUE,\dots)
\method{xl.write}{matrix}(r.obj,xl.rng,na="",row.names=TRUE,col.names=TRUE,\dots)
\method{xl.write}{factor}(r.obj,xl.rng,na="",row.names=TRUE,\dots)
\method{xl.write}{list}(r.obj,xl.rng,na="",\dots)
\method{xl.write}{table}(r.obj,xl.rng,na="",\dots)
\method{xl.write}{current.graphics}(r.obj,xl.rng,na="",delete.file=FALSE,\dots)
}
\arguments{
\item{r.obj}{R object}
\item{xl.rng}{An object of class 'COMIDispatch' (as used in RDCOMClient package) - reference to Excel range}
\item{na}{character. NA representaion in Excel. By default it is empty string}
\item{row.names}{a logical value indicating whether the row names of r.obj are to be written along with r.obj}
\item{col.names}{a logical value indicating whether the column names of r.obj are to be written along with r.obj}
\item{delete.file}{a logical value indicating whether delete file with graphic after insertion in Excel}
\item{\dots}{arguments for further processing}
}
%- maybe also 'usage' for other objects documented here.
\details{
'xl.rng' should be COM-reference to Excel range, not string. Method invisibly returns number of columns and rows occupied by 'r.obj' on Excel sheet. It's useful for multiple objects writing to prevent their overlapping.

It is more convenient to use 'xl' object. 'xl.write' aimed mostly for programming purposes, not for interactive usage. 
%%  ~~ If necessary, more details than the description above ~~
}
\value{
c(rows,columns)

Invisibly returns rows and columns number ocuppied by r.obj on Excel sheet. 
%%  ~Describe the value returned
%%  If it is a LIST, use
%%  \item{comp1 }{Description of 'comp1'}
%%  \item{comp2 }{Description of 'comp2'}
%% ...
}

\author{
Gregory Demin <excel.link.feedback@gmail.com>
}

\seealso{xl, xlr,xlc,xlrc, current.graphics}

\examples{

\dontrun{
xls=xl.get.excel()
xl.workbook.add()
rng=xls[["Activesheet"]]$Cells(1,1)
nxt=xl.write(iris,rng,row.names=TRUE,col.names=TRUE)
rng=rng$Offset(nxt[1]+1,0)
nxt=xl.write(cars,rng,row.names=TRUE,col.names=TRUE)
rng=rng$Offset(nxt[1]+1,0)
nxt=xl.write(as.data.frame(Titanic),rng,row.names=TRUE,col.names=TRUE)

xl.sheet.add()
rng=xls[["Activesheet"]]$Cells(1,1)
data.sets=list("Iris dataset",iris,"Cars dataset",cars,"Titanic dataset",as.data.frame(Titanic))
xl.write(data.sets,rng,row.names=TRUE,col.names=TRUE)

}
  
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
% \keyword{ ~xl }
\keyword{ IO }
\keyword{ connection }