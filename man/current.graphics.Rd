\name{current.graphics}
\alias{current.graphics}

\title{
Auxilary function for export graphics to Microsoft Excel 
}
\description{
If argument 'type' provided this function will save graphics from windows plotting device to temporary file and return path to this file. In other case (provided filename) 'type' will be ignored and resulted value is path to file  with class attribute 'current.graphics'. So it could be used with expressions such xl[a1] = current.graphics(filename="plot.png").
}
\usage{
current.graphics(type="emf",filename=NULL,\dots)
}
\arguments{
\item{type}{file type. Ignored if argument 'filename' provided.}
\item{filename}{filename (or full path) of file with graphics.}
\item{\dots}{arguments for internally used savePlot function}
}
%- maybe also 'usage' for other objects documented here.

\value{
Path to file with saved graphics with class attribute 'current.graphics'. If used with argument 'type' than result has attribute 'temp.file'=TRUE

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
plot(sin)
xl[a1] = current.graphics()
plot(cos)
cos.plot = current.graphics()
xl.sheet.add()
xl[a1] = list("Cosine plotting",cos.plot,"End of cosine plotting")

# the same thing without graphic windows 
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

}
  
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
% \keyword{ ~xl }
\keyword{ utilities }