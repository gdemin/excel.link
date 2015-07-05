#' @name xl.current.region
#' @title Read/write from/to Excel current region.
#'   
#' @description Current region is a region that will be selected by pressing 
#'   \code{Ctrl+Shift+*} in Excel. The current region is a range bounded by any 
#'   combination of blank rows and blank columns. \code{cr}, \code{crc}, 
#'   \code{crr}, \code{crrc} objects are already defined in the package. It 
#'   doesn't need to create or init them.
#'   
#' @param x One of \code{cr}, \code{crc}, \code{crr}, \code{crrc} objects. 
#'   \code{cr} - read/write with/without column and row names, "r" - with 
#'   rownames, "c" - with colnames
#' @param str.rng character Excel range. For single bracket operations it can be
#'   without quotes in almost all cases.
#' @param drop logical. If TRUE the result is coerced to the lowest possible 
#'   dimension. By default dimensions will be dropped if there are no columns 
#'   and rows names.
#' @param row.names logical value indicating whether the Excel range contains 
#'   the row names as its first column.
#' @param col.names logical value indicating whether the Excel range contains 
#'   the column names as its first row.
#' @param na character. NA representation in Excel. By default it is empty 
#'   string.
#' @param value suitable replacement value. All data will be placed in Excel
#'   sheet starting from top-left cell of current region. Current region will be
#'   cleared before writing.
#'   
#' @details \code{cr} object represents Microsoft Excel application. For 
#'   convenient interactive usage arguments can be given without quotes in most 
#'   cases (e. g. \code{cr[a1] = 5} or \code{cr[u2:u85] = "Hi"} or 
#'   \code{cr[MyNamedRange] = 42}, but \code{cr["Sheet1!A1"] = 42}). When it 
#'   used in your own functions or you need to use variable as argument it is 
#'   recommended apply double brackets notation: \code{cr[["a1"]] = 5} or 
#'   \code{cr[["u2:u85"]] = "Hi"} or \code{cr[["MyNamedRange"]] = 42}. 
#'   Difference between \code{cr}, \code{crc}, \code{crrc} and \code{crr} is 
#'   \code{cr} ignore row and column names, \code{crc} suppose read and write to
#'   Excel with column names, \code{crrc} - with column and row names and so on.
#'   There is argument \code{drop} which is \code{TRUE} by default for \code{cr}
#'   and \code{FALSE} by default for other options. 
#'   All these functions never coerce characters to factors
#'   
#' @return Returns appropriate dataset from Excel.
#' @aliases cr crrc crc crr
#' @seealso
#' \code{\link{xl}}
#'    
#' @examples
#' 
#' \dontrun{ 
#' data(iris)
#' data(mtcars)
#' xl.workbook.add()
#' xlc$a1 = iris
#' identical(crc[a1],xlc[a1:e151]) # should be TRUE
#' identical(crc$a1,xlc[a1:e151]) # should be TRUE
#' identical(crc$a1,xlc[a1]) # should be FALSE
#' 
#' # current region will be cleared before writing - no parts of iris dataset
#' crrc$a1 = mtcars 
#' identical(crrc$a1,xlrc[a1:l33]) # should be TRUE
#' 
#' }
#' @export
xl.current.region = function(str.rng,drop = TRUE,na = "",row.names = FALSE,col.names = FALSE)
    # return current region from Microsoft Excel (region selected when pressing Ctrl+Shift+*)
{
    ex = xl.get.excel()
    xl.rng = ex$range(str.rng)
    xl.read.range(xl.rng[["CurrentRegion"]],drop = drop,na = na,row.names = row.names,col.names = col.names)
} 





#' @export
cr = function()
{
    # run Excel if it's not running and
    # return reference to Microsoft Excel
    xl.get.excel()
}

# set class for usage '.[', '.[ = ' etc operators
class(cr) = c('cr','xl',class(cr))

#' @export
crrc = cr

#' @export
crc = cr

#' @export
crr = cr

has.rownames(cr) = FALSE 
has.colnames(cr) = FALSE 

has.rownames(crc) = FALSE 
has.colnames(crc) = TRUE 

has.rownames(crr) = TRUE 
has.colnames(crr) = FALSE 

has.rownames(crrc) = TRUE 
has.colnames(crrc) = TRUE 


#' @export
#' @rdname xl.current.region
'[[.cr' = function(x,str.rng,drop = !(has.rownames(x) | has.colnames(x)),na = "")
    ### return current region from Microsoft Excel. range.name is character string in form of standard
    ### Excel reference, e. g. ['A1:B5'], ['Sheet1!F8'], ['[Book3]Sheet7!B1'] or range name 
    ### The difference with '[' is that value should be quoted string. It's intended to use in user define functions
    ### or in cases where value is string variable with Excel range 
{
    xl.rng = x()$Range(str.rng)$CurrentRegion()
    xl.read.range(xl.rng,drop = drop,row.names = has.rownames(x),col.names = has.colnames(x),na = na)
}




#' @export
#' @rdname xl.current.region
'[[<-.cr' = function(x,str.rng,na = "",value)
{
    xl.rng = x()$Range(str.rng)$CurrentRegion()
    xl.rng$Clear()
    xl.write(value,xl.rng$Cells(1,1),row.names = has.rownames(x),col.names = has.colnames(x),na = na)
    x
}











