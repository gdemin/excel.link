% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/current.graphics.R
\name{current.graphics}
\alias{current.graphics}
\title{Auxiliary function for export graphics to Microsoft Excel}
\usage{
current.graphics(
  type = c("png", "jpeg", "bmp", "tiff"),
  filename = NULL,
  picname = NULL,
  ...
)
}
\arguments{
\item{type}{file type. Ignored if argument 'filename' provided.}

\item{filename}{character. filename (or full path) of file with graphics.}

\item{picname}{character. Picture name in Excel.}

\item{...}{arguments for internally used \code{\link{dev.copy}} function}
}
\value{
Path to file with saved graphics with class attribute 
  'current.graphics'. If used with argument \code{type} than result has 
  attribute \code{temp.file = TRUE}.
}
\description{
Auxiliary function for export graphics to Microsoft Excel
}
\details{
If argument \code{type} provided this function will save graphics 
  from windows plotting device to temporary file and return path to this 
  file. Argument \code{filename} is intended to transfer plots to Excel from 
  file-based graphics devices (see Examples) or just insert into Excel file 
  with graphics. If argument \code{filename} is provided argument \code{type}
  will be ignored and returned value is path to file \code{filename} with 
  class attribute 'current.graphics'. So it could be used with expressions 
  such \code{xl[a1] = current.graphics(filename="plot.png")}. If 
  \code{picname} is provided then picture will be inserted in Excel with this
  name. If picture \code{picname} already exists in Excel it will be deleted.
  This argument is useful when we need to change old picture in Excel instead
  of adding new picture. \code{picname} will be automatically prepended by
  "_" to avoid conflicts with Excel range names.
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
