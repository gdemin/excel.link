excel.link
==========

[![CRAN\_Status\_Badge](http://www.r-pkg.org/badges/version/excel.link)](https://cran.r-project.org/package=excel.link)
[![](http://cranlogs.r-pkg.org/badges/excel.link)](http://cran.rstudio.com/web/packages/excel.link/index.html)

❗ Microsoft Windows and Microsoft Excel are required for this package.

### Convenient Data Exchange with Microsoft Excel
Allows access to data in running instance of Microsoft Excel (e. g. `xl[a1] =
xl[b2]*3` and so on). Graphics can be transferred with `xl[a1] =
current.graphics()`. so on). Graphics can be transferred with 'xl[a1] =
current.graphics()'. Additionally there are function for reading/writing 'Excel'
files - 'xl.read.file'/'xl.save.file'. They are not fast but able to read/write
'*.xlsb'-files and password-protected files. There is an Excel workbook with
examples of calling R from Excel in the 'doc' folder. It tries to keep things as
simple as possible - there are no needs in any additional installations besides
R, only 'VBA' code in the Excel workbook. Microsoft Excel is required for this
package.

The excel.link package mainly consists of two rather independent parts: one
is for transferring data/graphics to running instance of Excel, another part - work with data table in Excel in similar way as with usual data.frame.

##### Transferring data

 Package provided family of objects:  `xl`, `xlc`, `xlr` and `xlrc`. You don't need to initialize these objects or to do any other preliminary actions. Just after execution `library(excel.link)` you can transfer data to Excel active sheet by simple assignment, for example: `xlrc[a1] = iris`. In this notation 'iris' dataset will be written with column and row names. If you doesn't need column/row names just remove 'r'/'c' letters (`xlc[a1] = iris` - with column names but without row names). To read Excel data just type something like this: `xl[a1:b5]`. You will get data.frame with values from range a1:a5 without column and row names. It is possible to use named ranges (e. g. `xl[MyNamedRange]`). To transfer graphics use `xl[a1] = current.graphics()`.
 
##### Live connection

For example we put iris datasset to Excel sheet:
 `xlc[a1] = iris`. After that we connect Excel range with R object: `xl_iris = xl.connect.table("a1",row.names = FALSE, col.names = TRUE)`. 
So we can: 
- get data from this Excel range: `xl_iris$Species` 
- add new data to this Excel range: `xl_iris$new_column = 42`
- sort this range: `sort(xl_iris,column = "Sepal.Length")` 
- and more...

# Aknowledgements

To comply CRAN policy includes source code from RDCOMClient package (http://www.omegahat.net/RDCOMClient) by Duncan Temple Lang (duncan at wald.ucdavis.edu).
