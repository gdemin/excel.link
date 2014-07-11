\name{xl.workbooks}
\alias{xl.workbooks}
\alias{xl.workbook.add}
\alias{xl.workbook.open}
\alias{xl.workbook.close}
\alias{xl.workbook.save}
\alias{xl.workbook.activate}

\title{
Basic operations with Excel workbooks
}
\description{
Functions for basic operations with Excel workbooks
}
\usage{
xl.workbooks()
xl.workbook.add(filename=NULL)
xl.workbook.open(filename)
xl.workbook.close(xl.workbook.name=NULL)
xl.workbook.save(filename)
xl.workbook.activate(xl.workbook.name)
}
\arguments{
\item{filename}{Excel workbook filename}
\item{xl.workbook.name}{Excel workbook name}
}
%- maybe also 'usage' for other objects documented here.
\details{
'xl.workbook.close' closes workbook with given name. If name doesn't submitted it closed active workbook.  It doesn't prompt about saving so if you don't save changes before closing all changes will be lost. 

'xl.workbook.save' saves active workbook. If only filename submitted it saves in the working directory. If name of workbook is omitted than new workbook is saved under its default name in the current working directory. It doesn't prompt about overwriting if file already exists.

'xl.workbook.add' adds new workbook and invisibly returns name of this newly created workbook. Added workbook become active. If 'filename' argument provided then Excel workbook 'filename' will be used as template.

'xl.workbook.activate' activates workbook with given name. If workbook with this name doesn't exists error will be generated. 

'xl.workbooks' returns character vector of names of all opened workbooks. 

%%  ~~ If necessary, more details than the description above ~~
}
\value{
'xl.workbooks' returns character vector of open workbooks.
'xl.workbook.save' invisibly returns path to the saved workbook
'xl.workbook.add'/'xl.workbook.activate'/'xl.workbook.activate' invisibly returns name of created/activated workbook.
%%  ~Describe the value returned
%%  If it is a LIST, use
%%  \item{comp1 }{Description of 'comp1'}
%%  \item{comp2 }{Description of 'comp2'}
%% ...
}

\author{
Gregory Demin <excel.link.feedback@gmail.com>
}

\seealso{xl.sheets}

\examples{

\dontrun{

## senseless actions
xl.workbook.add()
xlrc[a1] <- iris
xl.workbook.save("iris.xlsx")
xl.workbook.add()
xlrc[a1] <- cars
xl.workbook.save("cars.xlsx")
xl.workbook.activate("iris")
xl.workbook.close("cars")
xl.workbook.open("cars.xlsx")
xl.workbooks()
for (wb in xl.workbooks()) xl.workbook.close(wb)
unlink("iris.xlsx")
unlink("cars.xlsx")

}
  
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
% \keyword{ ~xl }
\keyword{ IO }
\keyword{ connection }