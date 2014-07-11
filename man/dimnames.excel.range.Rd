\name{dimnames.excel.range}
\alias{dimnames.excel.range}
\alias{dim.excel.range}
\alias{has.colnames}
\alias{has.rownames}
\alias{has.colnames.default}
\alias{has.rownames.default}
\alias{has.colnames<-}
\alias{has.rownames<-}
\alias{has.colnames.excel.range}
\alias{has.rownames.excel.range}

\title{
Operations with column and row names of 'excel.range' and 'xl' classes.
}
\description{
Operations with column and row names of 'excel.range' and 'xl' classes
}
\usage{
\method{dimnames}{excel.range}(x)
\method{dim}{excel.range}(x)

has.colnames(x)
has.rownames(x)
\method{has.colnames}{default}(x)
\method{has.rownames}{default}(x)
\method{has.colnames}{excel.range}(x)
\method{has.rownames}{excel.range}(x)
has.colnames(x) <- value
has.rownames(x) <- value
}
\arguments{
\item{x}{xl or excel.range object}
\item{value}{logical value}
}
%- maybe also 'usage' for other objects documented here.
\details{
'dimnames' read column/row names if xl/excel.range object has it. In opposite case it returns column/row names as in Excel (e. g. "a","b", "c"). 

'has.colnames' and 'has.rownames' sets/gets apropriate attributes of excel.range and xl objects. They are intended for internal usage.
%%  ~~ If necessary, more details than the description above ~~
}
\value{
'dimnames' returns row and column names.

'dim' returns object row and column numbers 

'has.rownames'/'has.colnames' - logical value. TRUE if excel.range/xl object has rownames/colnames.
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
xl.workbook.add()
xl[d4]=iris
test1=xl.connect.table("d4",row.names=FALSE,col.names=FALSE)
has.colnames(test1) # FALSE
has.rownames(test1) # FALSE
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

}
  
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
% \keyword{ ~xl }
\keyword{ IO }
\keyword{ connection }
\keyword{ utilities }
