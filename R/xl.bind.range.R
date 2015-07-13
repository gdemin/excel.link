# Idea by Stefan Fritsch (https://github.com/gdemin/excel.link/issues/1) 


#' Active bindings to Excel ranges
#' 
#' \code{xl.bind.range} and \code{xl.bind.current.region} create \code{sym} in 
#' environment \code{env} so that getting the value of \code{sym} return bound 
#' Excel range, and assigning to \code{sym} will write the value to be assigned 
#' to Excel range. In case of \code{xl.bind.range}  range will be updated after 
#' each assignment accordingly to the size of the assigned value.
#' \code{xl.bind.current.region} always returns data from current region 
#' (Ctrl+Shift+* in Excel) of bound range. 
#' \code{\%<-xl\%} etc are shortcuts for \code{xl.bind.range} and 
#' \code{xl.bind.current.region}. "r" means with row names, "c" means with
#' column names. Range in most cases can be provided without quotes: \code{a1
#' \%<-xl\% a1:b100}. Functions with '='  and with '<-' in the names do the same
#' things - they are just for those who prefer '=' assignment and for those who
#' prefer '<-' assignment.
#' Assignment and reading may be slow because these functions always read/write
#' entire dataset.
#' 
#' @usage 
#' xl.bind.range(sym, str.range, drop = TRUE, na = "",
#'      row.names = FALSE, col.names = FALSE, env = parent.frame()) 
#' 
#' xl.bind.current.region(sym, str.range, drop = TRUE, na = "",
#'      row.names = FALSE, col.names = FALSE, env = parent.frame()) 
#'  
#' sym \%<-xl\% value 
#' 
#' sym \%<-xlr\% value
#' 
#' sym \%<-xlc\% value
#' 
#' sym \%<-xlrc\% value
#' 
#' sym \%<-cr\% value 
#' 
#' sym \%<-crr\% value
#' 
#' sym \%<-crc\% value
#' 
#' sym \%<-crrc\% value
#' 
#' @param sym character/active binding.
#' @param value character Excel range address. It can be without quotes in many cases.
#' @param str.range character Excel range.
#' @param drop logical. If TRUE the result is coerced to the lowest possible 
#'   dimension. By default dimensions will be dropped if there are no columns 
#'   and rows names.
#' @param row.names logical value indicating whether the Excel range contains 
#'   the row names as its first column.
#' @param col.names logical value indicating whether the Excel range contains 
#'   the column names as its first row.
#' @param na character. NA representation in Excel. By default it is empty 
#'   string.
#' @param env an environment.
#'   
#' @return \code{xl.binding.address} returns list with three components about 
#'   bound Excel range: \code{address}, \code{rows} - number of rows, 
#'   \code{columns} - number of columns. All other functions don't return 
#'   anything but create active binding to Excel range in the environment.
#'   
#' @aliases xl.bind.current.region %<-xl% %<-xlc% %<-xlr% %<-xlrc%  %<-cr% %<-crc% %<-crr% %<-crrc%
#' @seealso \code{\link{xl}}, \code{\link{xlr}}, \code{\link{xlc}}, 
#'   \code{\link{xlrc}}
#'   
#' @author Idea by Stefan Fritsch
#'   (\url{https://github.com/gdemin/excel.link/issues/1})
#'   
#' @examples 
#' \dontrun{
#'  xl.workbook.add()
#'  range_a1 %=xl% a1 # binding range_a1 to cell A1 on active sheet
#'  range_a1 # should be NA
#'  range_a1 = 42 # value in Excel should be changed
#'  identical(range_a1, 42) 
#'  cr_a1 %=cr% a1 # binding cr_a1 to current region around cell A1 on active sheet
#'  identical(cr_a1, range_a1)
#'  # difference between 'cr' and 'xl':  
#'  xl[a2] = 43
#'  range_a1 # 42
#'  xl.binding.address(range_a1)
#'  xl.binding.address(cr_a1)
#'  cr_a1 # identical to 42:43
#'  # make cr and xl identical: 
#'  range_a1 = 42:43
#'  identical(cr_a1, range_a1)
#'  
#'  xl_iris %=crc% a1 # bind current region A1 on active sheet with column names
#'  xl_iris = iris # put iris dataset to Excel sheet
#'  identical(xl_iris$Sepal.Width, iris$Sepal.Width) # should be TRUE
#'  
#'  xl_iris$new_col = xl_iris$Sepal.Width*xl_iris$Sepal.Length # add new column on Excel sheet
#'  
#' }
#' @export
xl.bind.range = function(sym, str.range, drop = TRUE, na = "", row.names = FALSE, col.names = FALSE, env = parent.frame())
{
    
    if (exists(sym, env)) remove(list=sym, envir=env)
    
    
    xl = xl.get.excel()
    xl.rng = xl$Range(str.range)
    assignment = function(value=NULL){
        if (missing(value)) {
            xl.read.range(xl.rng, drop=drop, row.names=row.names, col.names=col.names, na = na)
        } else {
            if(is.null(value)){
                cat(paste0(xl.rng$address(External = TRUE),
                    "\n", xl.rng$rows()$count(),
                    "\n", xl.rng$columns()$count()))        
            } else {
                xl.rng$clear()
                if(is.atomic(value) && length(value)<2 && is.null(attributes(value))){
                    res = xl.write(value,xl.rng$cells(1,1), na = na, row.names = row.names, col.names = col.names)
                } else {
                    res = xl.write(value,xl.rng, na = na, row.names = row.names, col.names = col.names)
                } 
                if (res[1]>0) res[1] = res[1] - 1
                if (res[2]>0) res[2] = res[2] - 1
                xl.rng <<- xl$range(xl.rng$cells(1,1),xl.rng$cells(1,1)$offset(res[1],res[2]))
            }
        }
    }
    # assign active binding:   
    makeActiveBinding(
        sym,
        assignment,
        env
    )
    
}

#' @export
xl.bind.current.region = function(sym, str.range, drop = TRUE, na = "", row.names = FALSE, col.names = FALSE, env = parent.frame())
{
    
    if (exists(sym, env)) remove(list=sym, envir=env)
    xl = xl.get.excel()
    xl.rng = xl$Range(str.range)
    assignment = function(value){
        curr.rng = xl.rng$CurrentRegion()
        if (missing(value)) {
            xl.read.range(curr.rng, drop=drop, row.names=row.names, col.names=col.names, na = na)
        } else 
        {
            if(is.null(value)){
                
                cat(paste0(xl.rng$currentregion()$address(External = TRUE),
                    "\n",xl.rng$currentregion()$rows()$count(),
                    "\n",xl.rng$currentregion()$columns()$count()))
            } else {
                curr.rng$clear()
                xl.write(value,curr.rng$cells(1,1), na = na, row.names = row.names, col.names = col.names)
                
            }
        }
    }
    
    # assign active binding:   
    makeActiveBinding(
        sym,
        assignment,
        env
    )
    
}

bind.generator = function(row.names, col.names, fun) {
    function(sym, value){
        sym = deparse(substitute(sym))
        value = substitute(value)
        if (!is.character(value)) value = deparse(value)
        env = parent.frame()
        fun(sym, value, row.names = row.names, col.names = col.names, env = env)
        
    }
}

#' @export
#' @rdname xl.bind.range
"%=xl%" = bind.generator(row.names = FALSE, col.names = FALSE, fun = xl.bind.range)
#' @export
#' @rdname xl.bind.range
"%=xlr%" = bind.generator(row.names = TRUE, col.names = FALSE, fun = xl.bind.range)
#' @export
#' @rdname xl.bind.range
"%=xlc%" = bind.generator(row.names = FALSE, col.names = TRUE, fun = xl.bind.range)
#' @export
#' @rdname xl.bind.range
"%=xlrc%" = bind.generator(row.names = TRUE, col.names = TRUE, fun = xl.bind.range)

#' @export
#' @rdname xl.bind.range
"%=cr%" = bind.generator(row.names = FALSE, col.names = FALSE, fun = xl.bind.current.region)
#' @export
#' @rdname xl.bind.range
"%=crr%" = bind.generator(row.names = TRUE, col.names = FALSE, fun = xl.bind.current.region)
#' @export
#' @rdname xl.bind.range
"%=crc%" = bind.generator(row.names = FALSE, col.names = TRUE, fun = xl.bind.current.region)
#' @export
#' @rdname xl.bind.range
"%=crrc%" = bind.generator(row.names = TRUE, col.names = TRUE, fun = xl.bind.current.region)



#' @export
`%<-cr%` = `%=cr%`
#' @export
`%<-crr%` = `%=crr%`
#' @export
`%<-crc%` = `%=crc%`
#' @export
`%<-crrc%` = `%=crrc%`

#' @export
`%<-xl%` = `%=xl%`
#' @export
`%<-xlr%` = `%=xlr%`
#' @export
`%<-xlc%` = `%=xlc%`
#' @export
`%<-xlrc%` = `%=xlrc%`

#' @export
#' @rdname xl.bind.range
xl.binding.address = function(sym){
    sym = substitute(sym)
    if (!is.character(sym)) sym = deparse(sym)
    res = eval(parse(text = paste0("capture.output(",sym,"<-NULL)")),envir = parent.frame())
    res = strsplit(res, split = "\n")
    names(res) = c("address","rows","columns")
    res$rows = as.integer(res$rows)
    res$columns = as.integer(res$columns)
    res
}

