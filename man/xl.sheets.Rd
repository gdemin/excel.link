\name{xl.sheets}
\alias{xl.sheets}
\alias{xl.sheet.add}
\alias{xl.sheet.delete}
\alias{xl.sheet.activate}

\title{
Basic operations with worksheets
}
\description{
Functions for basic operations with worksheets
}
\usage{
xl.sheets()
xl.sheet.add(xl.sheet.name=NULL,before=NULL)
xl.sheet.delete(xl.sheet=NULL)
xl.sheet.activate(xl.sheet)
}
\arguments{
\item{xl.sheet.name}{sheet name in active workbook}
\item{before}{sheet name or sheet number in active workbook before which new sheet will be added}
\item{xl.sheet}{sheet name or sheet number in active workbook}
}
%- maybe also 'usage' for other objects documented here.
\details{
'xl.sheet.delete' deletes sheet with given name/number. If name doesn't submitted it delete active sheet.
  
'xl.sheet.add' adds new sheet with given name and invisibly returns name of this newly added sheet. Added sheet become active. If 'xl.sheet.name' is missing default name will be used. If 'before' argument is missing, sheet will be added at the last position. If sheet with given name already exists error will be generated. 

'xl.sheet.activate' activates sheet with given name/number. If sheet with this name doesn't exist error will be generated. 
%%  ~~ If necessary, more details than the description above ~~
}
\value{
'xl.sheets' returns vector of sheet names in active workbook.

'xl.sheet.add'/'xl.sheet.activate' invisibly returns name of created/activated sheet.
  
%%  ~Describe the value returned
%%  If it is a LIST, use
%%  \item{comp1 }{Description of 'comp1'}
%%  \item{comp2 }{Description of 'comp2'}
%% ...
}

\author{
Gregory Demin <excel.link.feedback@gmail.com>
}

\seealso{xl.workbooks}

\examples{

\dontrun{
xl.workbook.add()
sheets <- xl.sheets()
xl.sheet.add("Second")
xl.sheet.add("First",before="Second")
for (sheet in sheets) xl.sheet.delete(sheet) # only 'First' and 'Second' exist in workbook now
xl.sheet.activate("Second") #last sheet activated 

}
  
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
% \keyword{ ~xl }
\keyword{ IO }
\keyword{ connection }