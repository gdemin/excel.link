#' @name xl
#' @title Data exchange with running Microsoft Excel instance.
#'   
#' @description 
#' \code{xl}, \code{xlc}, \code{xlr}, \code{xlrc} objects are
#' already defined in the package. It doesn't need to create or init them. Just
#' after attaching package one can write something like this: \code{xl[a1] =
#' "Hello, world!"} and this text should appears in \code{A1} cell on active
#' sheet of active Excel workbook.
#' 
#' @param x One of \code{xl}, \code{xlc}, \code{xlr}, \code{xlrc} objects. 
#'   \code{xl} - read/write with/without column and row names, "r" - with 
#'   rownames, "c" - with colnames
#' @param str.rng character Excel range. For single bracket operations it can be
#'   without quotes in almost all cases.
#' @param drop a logical. If TRUE the result is coerced to the lowest possible 
#'   dimension. By default dimensions will be dropped if there are no columns 
#'   and rows names.
#' @param row.names a logical value indicating whether the Excel range contains 
#'   the row names as its first column.
#' @param col.names a logical value indicating whether the Excel range contains 
#'   the column names as its first row.
#' @param na character. NA representation in Excel. By default it is empty 
#'   string.
#' @param value a suitable replacement value. It will be recycled to fill excel 
#'   range only if it is object of length 1. In other cases size of excel range 
#'   is ignored - all data will be placed in Excel sheet starting from top-left 
#'   cell of submitted range.
#'   
#' @details \code{xl} object represents Microsoft Excel application. For 
#'   convenient interactive usage arguments can be given without quotes in most 
#'   cases (e. g. \code{xl[a1] = 5} or \code{xl[u2:u85] = "Hi"} or 
#'   \code{xl[MyNamedRange] = 42}, but \code{xl["Sheet1!A1"] = 42}). When it 
#'   used in your own functions or you need to use variable as argument it is 
#'   recommended apply double brackets notation: \code{xl[["a1"]] = 5} or 
#'   \code{xl[["u2:u85"]] = "Hi"} or \code{xl[["MyNamedRange"]] = 42}. 
#'   Difference between \code{xl}, \code{xlc}, \code{xlrc} and \code{xlr} is 
#'   \code{xl} ignore row and column names, \code{xlc} suppose read and write to
#'   Excel with column names, \code{xlrc} - with column and row names and so on.
#'   There is argument \code{drop} which is \code{TRUE} by default for \code{xl}
#'   and \code{FALSE} by default for other options. \code{xl.selection} returns 
#'   data.frame with data from current selection in Excel. 
#'   \code{xl.current.region} returns data.frame with data from current region 
#'   (range which can be selected by pressing \code{Ctrl+Shift+*}) in Excel. All
#'   these functions never coerce characters to factors
#'   
#' @return Returns appropriate dataset from Excel. Excel datetime type currently
#'   not supported.
#' @aliases xl xlrc xlc xlr
#'   
#' @examples
#' 
#' \dontrun{ 
#' data(iris)
#' rownames(iris) <- as.character(rownames(iris))
#' iris$Species <- as.character(iris$Species)
#' xl.workbook.add()
#' xlrc$a1 <- iris
#' xl.iris <- xl.current.region("a1",row.names=TRUE,col.names=TRUE)
#' identical(xl.iris,iris)
#' 
#' xl.sheet.add("Datasets examples")
#' data.sets <- list("Iris dataset",iris,"Cars dataset",cars,"Titanic dataset",as.data.frame(Titanic))
#' xlrc[a1] <- data.sets
#' 
#' }
#' @export
'[.xl' = function(x,str.rng,drop = !(has.rownames(x) | has.colnames(x)),na = "")
    ### return range from Microsoft Excel. range.name is character string in form of standard
    ### Excel reference, quotes can be omitted, e. g. [A1:B5], [Sheet1!F8], [[Book3]Sheet7!B1] or range name 
    ### Function is intended to use in interactive environement 
{
    # str.rng = as.character(sys.call())[3]
    str.rng = as.character(as.expression(substitute(str.rng)))
    x[[str.rng,drop = drop,na = na]]
}



#' @export
xl = function()
    {
        # run Excel if it's not running and
        # return reference to Microsoft Excel
       xl.get.excel()
}

# set class for usage '.[', '.[ = ' etc operators
class(xl) = c('xl',class(xl))

#' @export
xlrc = xl

#' @export
xlc = xl

#' @export
xlr = xl


has.rownames(xl) = FALSE 
has.colnames(xl) = FALSE 

has.rownames(xlc) = FALSE 
has.colnames(xlc) = TRUE 

has.rownames(xlr) = TRUE 
has.colnames(xlr) = FALSE 

has.rownames(xlrc) = TRUE 
has.colnames(xlrc) = TRUE 


#' @export
#' @rdname xl
'[[.xl' = function(x,str.rng,drop = !(has.rownames(x) | has.colnames(x)),na = "")
    ### return range from Microsoft Excel. range.name is character string in form of standard
    ### Excel reference, e. g. ['A1:B5'], ['Sheet1!F8'], ['[Book3]Sheet7!B1'] or range name 
    ### The difference with '[' is that value should be quoted string. It's intended to use in user define functions
    ### or in cases where value is string variable with Excel range 
{
    xl.rng = x()$Range(str.rng)
    xl.read.range(xl.rng,drop = drop,row.names = has.rownames(x),col.names = has.colnames(x),na = na)
}

#' @export
#' @rdname xl
'$.xl' = function(x,str.rng)
    ### return range from Microsoft Excel. range.name is character string in form of standard
    ### Excel reference, e. g. xl$'A1:B5', xl$'Sheet1!F8', xl$'[Book3]Sheet7!B1', xl$h3 or range name 
    ### The difference with '[' is that value should be quoted string. It's intended to use in user define functions
    ### or in cases where value is string variable with Excel range 
{
    x[[str.rng]]
}


#' @export
#' @rdname xl
'[[<-.xl' = function(x,str.rng,na = "",value)
{
    xl.rng = x()$Range(str.rng)
    xl.write(value,xl.rng,row.names = has.rownames(x),col.names = has.colnames(x),na = na)
    x
}


#' @export
#' @rdname xl
'$<-.xl' = function(x,str.rng,value)
{
    x[[str.rng]] = value
    x
}

#' @export
#' @rdname xl
'[<-.xl' = function(x,str.rng,na = "",value)
{
    str.rng = as.character(as.expression(substitute(str.rng)))
    x[[str.rng,na = na]] = value
    x
}

#' @export
#' @rdname xl
xl.selection = function(drop = TRUE,na = "",row.names = FALSE,col.names = FALSE)
    # return current selection from Microsoft Excel
{
    ex = xl.get.excel()
    xl.rng = ex[['Selection']]
    xl.read.range(xl.rng,drop = drop,na = na,row.names = row.names,col.names = col.names)
}


#' @export
#' @rdname xl
xl.current.region = function(str.rng,drop = TRUE,na = "",row.names = FALSE,col.names = FALSE)
    # return current region from Microsoft Excel (region selected when pressing Ctrl+Shift+*)
{
    ex = xl.get.excel()
    xl.rng = ex$range(str.rng)
    xl.read.range(xl.rng[["CurrentRegion"]],drop = drop,na = na,row.names = row.names,col.names = col.names)
}


xl.read.range = function(xl.rng,drop = TRUE,row.names = FALSE,col.names = FALSE,na = "")
    # return matrix/data.frame/vector from excel from given range
{
    if (col.names && (xl.rng[["rows"]][["count"]]<2)) col.names = FALSE
    if (row.names && (xl.rng[["columns"]][["count"]]<2)) row.names = FALSE
    data.list = xl.rng[['Value']]
    if (!is.list(data.list)) data.list = list(list(data.list))
    data.list = lapply(data.list, function(each.col) {
        lapply(each.col, function(each.cell){
            
            if(is.null(each.cell) || length(each.cell) == 0 || each.cell == na) NA else each.cell
            
        })
        
    })
    

    # here we create logical matrix to keep type of excel cell
    # currently there are three type NA - NA, TRUE - ComDate, FALSE - all others types
    # these complexities are needed to deal with datetime columns. 
    # if we have column with datetime and any other type (FALSE)
    # we convert it to characters. If datetime mixed only with NA it will become POSIXct
    # Main idea - data should look like as in Excel - no strange modifications
    # such as integers become dates or dates become integers.
    
    types = lapply(data.list, function(each.col){
        unlist(lapply(each.col, function(each.cell){
            ifelse(class(each.cell) == "COMDate", TRUE, ifelse(is.na(each.cell),NA,FALSE))

        }))
        
        
    })
    
    type_matrix = do.call(cbind,types)

    if (col.names) type_matrix = type_matrix[-1,,drop = FALSE]
    if (row.names) type_matrix = type_matrix[,-1,drop = FALSE]
    
    
    if (col.names)    {
        colNames = lapply(data.list,function(each) {
            res = each[[1]]
            if (class(res) == "COMDate"){
                gsub(" UTC","",excel_datetime2POSIXct(res),fixed = TRUE)
                
            }   else res
            
        })
        if (row.names) colNames = colNames[-1]
        data.list = lapply(data.list,"[",-1)
    }
    if (row.names) {
        rowNames = unlist(lapply(data.list[[1]], function(each) {
            if (class(each) == "COMDate"){
                gsub(" UTC","",excel_datetime2POSIXct(each),fixed = TRUE)
                
            }  else each                
        }))
        data.list = data.list[-1]
    }	
    data.list = lapply(data.list,unlist)
    # classes = unique(sapply(data.list,class))
    
    final.matrix = do.call(data.frame,list(data.list,stringsAsFactors = FALSE))
    # make types
    for (each.col in seq_len(ncol(final.matrix))){
        if(any(type_matrix[,each.col] %in% TRUE)){
            # we have datetime
            if(any(type_matrix[,each.col] %in% FALSE)){
               # we have datetime mixed with other types -> convert to string 
                date_part = final.matrix[type_matrix[,each.col] %in% TRUE,each.col]
                date_part = excel_datetime2POSIXct(date_part) 
                final.matrix[type_matrix[,each.col] %in% TRUE,each.col] = gsub(" UTC","",date_part,fixed = TRUE) 
            } else {
                # we have only date.time/NA in this column -> convert to POSIXct
                final.matrix[,each.col] = excel_datetime2POSIXct(final.matrix[,each.col]) 
                
            }
            
        }
        
    }

    if (row.names && anyDuplicated(rowNames)) {
        row.names = FALSE
        warning("There are duplicated rownames. They will be ignored.")
    }	
    if (row.names) {
        rownames(final.matrix) = rowNames
    } else {
        rownames(final.matrix) = xl.rownames(xl.rng)[ifelse(col.names,-1,TRUE)]
    }
    if (col.names) {
        colnames(final.matrix) = colNames 
    } else {
        colnames(final.matrix) = xl.colnames(xl.rng)[ifelse(row.names,-1,TRUE)]
    }
    if (ncol(final.matrix)<2 & drop) final.matrix = final.matrix[,1]
    final.matrix
}

excel_datetime2POSIXct = function(value){
    as.POSIXct(as.numeric(value)*86400, origin="1899-12-30", tz="UTC") # 60*60*24 = 86400

}
