#' Basic operations with worksheets.
#' 
#' @param xl.sheet.name character. sheet name in active workbook
#' @param before character/numeric. sheet name or sheet number in active
#'   workbook before which new sheet will be added
#' @param xl.sheet character/numeric. sheet name or sheet number in active
#'   workbook
#'   
#' @details \itemize{ 
#' \item{\code{xl.sheet.add}}{ adds new sheet with given name and invisibly 
#' returns name of this newly added sheet. Added sheet become active. If 
#' \code{xl.sheet.name} is missing default name will be used. If \code{before}
#' argument is missing, sheet will be added at the last position. If sheet with
#' given name already exists error will be generated.}
#' \item{\code{xl.sheet.activate}}{ activates sheet with given name/number. If 
#' sheet with this name doesn't exist error will be generated.}
#' \item{\code{xl.sheet.delete}}{ deletes sheet with given
#' name/number. If name doesn't submitted it delete active sheet.} 
#' }
#' 
#' @return
#' \itemize{
#' \item{\code{xl.sheet.add}/\code{xl.sheet.activate}}{ invisibly return name of 
#' created/activated sheet.}
#' \item{\code{xl.sheets}}{ returns vector of sheet names in active workbook.}
#' \item{\code{xl.sheet.delete}}{ invisibly returns NULL.}
#' }
#' 
#' @seealso
#' \code{\link{xl.workbooks}}
#' 
#' @examples
#' 
#' \dontrun{ 
#' xl.workbook.add()
#' sheets <- xl.sheets()
#' xl.sheet.add("Second")
#' xl.sheet.add("First", before="Second")
#' for (sheet in sheets) xl.sheet.delete(sheet) # only 'First' and 'Second' exist in workbook now
#' xl.sheet.activate("Second") #last sheet activated 
#' 
#' }
#' @export
xl.sheet.add = function(xl.sheet.name = NULL,before = NULL)
    ### add new sheet to active workbook after the last sheet with given name and invisibily return reference to it 
{
    ex = xl.get.excel()
    sh.count = ex[['ActiveWorkbook']][['Sheets']][['Count']]
    sheets = tolower(xl.sheets())
    if (!is.null(xl.sheet.name) && (tolower(xl.sheet.name) %in% sheets)) stop ('sheet with name "',xl.sheet.name,'" already exists.')
    if (is.null(before)) {
        res = ex[['ActiveWorkbook']][['Sheets']]$Add(After = ex[['ActiveWorkbook']][['Sheets']][[sh.count]])
    } else {
        before = xl.sheet.exists(before,sheets)
        res = ex[['ActiveWorkbook']][['Sheets']]$Add(Before = ex[['ActiveWorkbook']][['Sheets']][[before]])
    }    
    if (!is.null(xl.sheet.name)) res[['Name']] = substr(xl.sheet.name,1,63)
    invisible(res[['Name']])
}

#' @export
#' @rdname xl.sheet.add
xl.sheets = function()
    ### Return worksheets names in the active workbook 
{
    ex = xl.get.excel()
    sh.count = ex[['ActiveWorkbook']][['Sheets']][['Count']]
    sapply(seq_len(sh.count), function(sh) ex[['ActiveWorkbook']][['Sheets']][[sh]][['Name']])
}

#' @export
#' @rdname xl.sheet.add
xl.sheet.activate = function(xl.sheet)
    ### activate sheet with given name (number) in active workbook 
{
    ex = xl.get.excel()
    #on.exit(ex[["DisplayAlerts"]] = TRUE)
    xl.sheet = xl.sheet.exists(xl.sheet)
    xl.sh = ex[['ActiveWorkbook']][['Sheets']][[xl.sheet]]
    #ex[["DisplayAlerts"]] = FALSE
    xl.sh$activate()
    invisible(xl.sh[['Name']])
}

xl.sheet.exists = function(xl.sheet,all.sheets = xl.sheets())
    ## check exsistense of xl.sheet in all.sheets and return xl.sheet position in all.sheets 
{
    UseMethod("xl.sheet.exists")
}



xl.sheet.exists.numeric = function(xl.sheet,all.sheets = xl.sheets())
{
    if (xl.sheet>length(all.sheets)) stop ("too large sheet number. In workbook only ",length(all.sheets)," sheet(s)." )
    xl.sheet
}



xl.sheet.exists.character = function(xl.sheet,all.sheets = xl.sheets())
{
    xl.sheet = which(tolower(xl.sheet) == tolower(all.sheets)) 
    if (length(xl.sheet) == 0) stop ("sheet ",xl.sheet," doesn't exist." )
    xl.sheet
}


#' @export
#' @rdname xl.sheet.add
xl.sheet.delete = function(xl.sheet = NULL)
    ### delete sheet with given name(number) in active workbook 
{
    ex = xl.get.excel()
    on.exit(ex[["DisplayAlerts"]] <- TRUE)
    if (is.null(xl.sheet)) {
        xl.sh = ex[['ActiveWorkbook']][["ActiveSheet"]]
    } else {
        xl.sheet = xl.sheet.exists(xl.sheet)
        xl.sh = ex[['ActiveWorkbook']][['Sheets']][[xl.sheet]]
    }	
    ex[["DisplayAlerts"]] = FALSE
    xl.sh$Delete()
    invisible(NULL)
}

